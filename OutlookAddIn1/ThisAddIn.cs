using System;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private Outlook.Items _calendarItems;
        private Outlook.MAPIFolder _calendarFolder; // store to release on shutdown
        private Microsoft.Office.Tools.CustomTaskPane _manageTimesheetPane;
        private System.Windows.Forms.Timer _paneWidthTimer;

        // PERFORMANCE: Cache current user email (avoid repeated COM calls)
        private string _cachedUserEmail = null;

        // PERFORMANCE: Once the "Timesheet Submitted" category is confirmed correct, skip COM scan
        private bool _timesheetCategoryEnsured = false;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Wire ItemSend immediately (lightweight, no MAPI access)
            this.Application.ItemSend += Application_ItemSend;

            // Defer heavy COM/MAPI work (GetDefaultFolder, Items, SMTP lookup)
            // to the first idle moment after Outlook finishes loading.
            // This keeps the add-in startup instant.
            System.Windows.Forms.Application.Idle += OnFirstIdle;
        }

        private void OnFirstIdle(object sender, EventArgs e)
        {
            // Unsubscribe immediately — we only need this once
            System.Windows.Forms.Application.Idle -= OnFirstIdle;

            try
            {
                _calendarFolder = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                _calendarItems = _calendarFolder.Items;

                _calendarItems.ItemAdd += CalendarItems_ItemAdd;
                _calendarItems.ItemChange += CalendarItems_ItemChange;

                // Pre-fetch user email (SMTP resolution via Exchange)
                _cachedUserEmail = GetCurrentUserSmtpFromSession();

                System.Diagnostics.Debug.WriteLine($"Deferred init complete. Email: {_cachedUserEmail}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Deferred init error: {ex.Message}");
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("CreateRibbonExtensibilityObject: Creating CalendarRibbon");
                return new CalendarRibbon();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CreateRibbonExtensibilityObject FAILED: {ex.Message}");
                return null;
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                if (_calendarItems != null)
                {
                    _calendarItems.ItemAdd -= CalendarItems_ItemAdd;
                    _calendarItems.ItemChange -= CalendarItems_ItemChange;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_calendarItems);
                    _calendarItems = null;
                }

                if (_calendarFolder != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_calendarFolder);
                    _calendarFolder = null;
                }

                // Clean up custom task pane
                if (_paneWidthTimer != null)
                {
                    _paneWidthTimer.Stop();
                    _paneWidthTimer.Dispose();
                    _paneWidthTimer = null;
                }
                if (_manageTimesheetPane != null)
                {
                    this.CustomTaskPanes.Remove(_manageTimesheetPane);
                    _manageTimesheetPane = null;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ThisAddIn_Shutdown error: {ex.Message}");
            }
        }

        private void CalendarItems_ItemAdd(object Item)
        {
            // Currently no action needed on ItemAdd — timesheet submission is manual
        }

        private void CalendarItems_ItemChange(object Item)
        {
            var appt = Item as Outlook.AppointmentItem;
            if (appt == null) return;

            // ✅ CRITICAL OPTIMIZATION: Get UserProperties ONCE and reuse
            Outlook.UserProperties ups = null;
            try
            {
                ups = appt.UserProperties;

                // Read all user properties in ONE batch (fastest!)
                var programCode = GetUPFromCollection(ups, "ProgramCode");

                // Skip if no timesheet data
                if (string.IsNullOrWhiteSpace(programCode))
                    return;

                // Skip draft meetings (not sent yet)
                var processOnSend = GetUPFromCollection(ups, "ProcessOnSend");
                if (string.Equals(processOnSend, "true", StringComparison.OrdinalIgnoreCase))
                    return;

                // ✅ OPTIMIZED: Batch ALL COM property reads together
                var meetingStatus = appt.MeetingStatus;
                var entryId = appt.EntryID ?? "";
                var subject = appt.Subject ?? "";
                var startUtc = appt.StartUTC;
                var endUtc = appt.EndUTC;
                var lastMod = appt.LastModificationTime.ToUniversalTime();
                var isRecurring = appt.IsRecurring;
                var activityCode = GetUPFromCollection(ups, "ActivityCode");
                var stageCode = GetUPFromCollection(ups, "StageCode");

                // Skip cancelled meetings
                if (meetingStatus == Outlook.OlMeetingStatus.olMeetingCanceled ||
                    meetingStatus == Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled)
                    return;

                // Get current user email
                var currentUserEmail = GetCurrentUserSmtpFromSession();
                if (!IsTti(currentUserEmail))
                    return;

                var source = "ItemChange";
                var globalIdFallback = entryId;

                System.Diagnostics.Debug.WriteLine($"{source}: {subject} - Start: {startUtc:yyyy-MM-dd HH:mm}");

                // Fire-and-forget: existence check + DB upsert on background thread
                _ = CalendarItems_ItemChange_BackgroundAsync(
                    source, entryId, globalIdFallback, subject, startUtc, endUtc,
                    programCode, activityCode, stageCode, currentUserEmail, lastMod, isRecurring);
            }
            finally
            {
                // ✅ CRITICAL: Release UserProperties collection
                if (ups != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ups);
                    ups = null;
                }
            }
        }

        // ✅ NEW: Helper to get user property from existing collection (no new COM calls!)
        private static string GetUPFromCollection(Outlook.UserProperties ups, string name)
        {
            Outlook.UserProperty up = null;
            try
            {
                up = ups?.Find(name);
                return (up != null && up.Value != null) ? up.Value.ToString() : "";
            }
            catch { return ""; }
            finally
            {
                if (up != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(up);
                    up = null;
                }
            }
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            try
            {
                // ✅ NEW: Detect meeting cancellations FIRST
                var meetingItem = Item as Outlook.MeetingItem;
                if (meetingItem != null && meetingItem.MessageClass == "IPM.Schedule.Meeting.Canceled")
                {
                    Outlook.AppointmentItem cancelAppt = null;
                    Outlook.UserProperties cancelUps = null;

                    try
                    {
                        cancelAppt = meetingItem.GetAssociatedAppointment(false);
                        if (cancelAppt != null)
                        {
                            var cancelGlobalId = Safe<string>(() => cancelAppt.GlobalAppointmentID) ?? "";
                            var cancelEntryId = cancelAppt.EntryID ?? "";
                            var cancelSubject = cancelAppt.Subject ?? "";

                            System.Diagnostics.Debug.WriteLine($"Meeting cancellation detected: {cancelSubject} (GlobalID: {cancelGlobalId})");

                            // ✅ CRITICAL: Remove metadata to prevent ItemChange from re-inserting
                            try
                            {
                                cancelUps = cancelAppt.UserProperties;
                                var programCodeProp = cancelUps.Find("ProgramCode");
                                var activityCodeProp = cancelUps.Find("ActivityCode");
                                var stageCodeProp = cancelUps.Find("StageCode");
                                var processOnSendProp = cancelUps.Find("ProcessOnSend");

                                if (programCodeProp != null) { programCodeProp.Delete(); System.Runtime.InteropServices.Marshal.ReleaseComObject(programCodeProp); }
                                if (activityCodeProp != null) { activityCodeProp.Delete(); System.Runtime.InteropServices.Marshal.ReleaseComObject(activityCodeProp); }
                                if (stageCodeProp != null) { stageCodeProp.Delete(); System.Runtime.InteropServices.Marshal.ReleaseComObject(stageCodeProp); }
                                if (processOnSendProp != null) { processOnSendProp.Delete(); System.Runtime.InteropServices.Marshal.ReleaseComObject(processOnSendProp); }

                                cancelAppt.Save();
                                System.Diagnostics.Debug.WriteLine("Removed metadata from cancelled meeting to prevent re-insertion");
                            }
                            catch (Exception metaEx)
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed to remove metadata: {metaEx.Message}");
                            }

                            // Fire-and-forget deletion for ALL users
                            _ = ItemSend_ProcessCancellationAsync(cancelGlobalId, cancelEntryId, cancelSubject);
                        }
                    }
                    finally
                    {
                        if (cancelUps != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(cancelUps);
                            cancelUps = null;
                        }
                        if (cancelAppt != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(cancelAppt);
                            cancelAppt = null;
                        }
                    }

                    // Don't process as a normal send
                    return;
                }

                // === Process normal meeting sends ===
                Outlook.AppointmentItem appt2 = Item as Outlook.AppointmentItem;
                if (appt2 == null)
                {
                    var mi = Item as Outlook.MeetingItem;
                    if (mi != null)
                        appt2 = mi.GetAssociatedAppointment(false);
                }
                if (appt2 == null) return;

                // ✅ OPTIMIZED: Get UserProperties ONCE
                Outlook.UserProperties ups2 = null;
                try
                {
                    ups2 = appt2.UserProperties;

                    // Read metadata in one batch
                    var programCode = GetUPFromCollection(ups2, "ProgramCode");
                    var activityCode = GetUPFromCollection(ups2, "ActivityCode");
                    var stageCode = GetUPFromCollection(ups2, "StageCode");

                    // ✅ CRITICAL FIX: Only proceed if ProgramCode exists AND this is NOT a new meeting
                    // New meetings have metadata but should NOT auto-submit on send
                    if (string.IsNullOrWhiteSpace(programCode))
                        return;

                    // ✅ NEW: Check if this is a new meeting (no prior submission)
                    // If there's no existing timesheet, this is a NEW meeting → DO NOT auto-submit
                    var entryId = appt2.EntryID ?? "";
                    var subject = appt2.Subject ?? "";
                    var startUtc = appt2.StartUTC;
                    var organizerEmail = GetCurrentUserSmtpFromSession();

                    if (!IsTti(organizerEmail))
                        return;

                    // Read all remaining COM properties NOW before handing off to background
                    var endUtc = appt2.EndUTC;
                    var lastMod = appt2.LastModificationTime.ToUniversalTime();
                    var initialGlobalId = Safe<string>(() => appt2.GlobalAppointmentID) ?? entryId;

                    // Clear ProcessOnSend flag synchronously (must happen before send completes)
                    try
                    {
                        var processOnSendProp = ups2.Find("ProcessOnSend");
                        if (processOnSendProp != null)
                        {
                            processOnSendProp.Delete();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(processOnSendProp);
                            System.Diagnostics.Debug.WriteLine("Cleared ProcessOnSend flag");
                        }
                    }
                    catch (Exception flagEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to clear flag: {flagEx.Message}");
                    }

                    // Fire-and-forget: existence check (was .Result) + retry loop + DB upsert
                    _ = ItemSend_ProcessSendAsync(
                        entryId, initialGlobalId, subject, startUtc, endUtc, lastMod,
                        programCode, activityCode, stageCode, organizerEmail);
                }
                finally
                {
                    if (ups2 != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ups2);
                        ups2 = null;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Application_ItemSend failed: " + ex.Message);
            }
        }

        // Ensures "Timesheet Submitted" category exists with the correct colour.
        // Result is cached after the first call so subsequent submits skip the COM scan.
        public void EnsureTimesheetCategory()
        {
            if (_timesheetCategoryEnsured) return;
            const string categoryName = "Timesheet Submitted";
            const Outlook.OlCategoryColor categoryColor = Outlook.OlCategoryColor.olCategoryColorPeach;

            Outlook.NameSpace session = null;
            Outlook.Categories categories = null;
            try
            {
                session    = this.Application.Session;
                categories = session.Categories;
                bool found = false;
                foreach (Outlook.Category c in categories)
                {
                    Outlook.Category captured = c; // capture for finally
                    try
                    {
                        if (c.Name == categoryName)
                        {
                            if (c.Color != categoryColor)
                            {
                                categories.Remove(categoryName);
                                categories.Add(categoryName, categoryColor);
                            }
                            found = true;
                            break;
                        }
                    }
                    finally { Marshal.ReleaseComObject(captured); }
                }
                if (!found)
                    categories.Add(categoryName, categoryColor);
                _timesheetCategoryEnsured = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"EnsureTimesheetCategory failed: {ex.Message}");
            }
            finally
            {
                if (categories != null) Marshal.ReleaseComObject(categories);
                if (session    != null) Marshal.ReleaseComObject(session);
            }
        }

        // Applies "Timesheet Submitted" category to the appointment and saves it.
        private void ApplyCategoryToAppointment(Outlook.AppointmentItem appt)
        {
            try
            {
                const string categoryName = "Timesheet Submitted";

                // PERF: Create/validate the category once; skip COM scan on repeat calls
                EnsureTimesheetCategory();

                // Merge with any non-timesheet categories already on the appointment
                if (!string.IsNullOrEmpty(appt.Categories))
                {
                    var existing = appt.Categories
                        .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(c => c.Trim())
                        .Where(c => !c.Equals("Timesheet Submitted", StringComparison.OrdinalIgnoreCase) &&
                                    !c.Equals("Timesheet Ignored",   StringComparison.OrdinalIgnoreCase))
                        .ToList();
                    existing.Add(categoryName);
                    appt.Categories = string.Join(", ", existing);
                }
                else
                {
                    appt.Categories = categoryName;
                }

                appt.Save();
                System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment: Applied '{categoryName}' to '{appt.Subject}'");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment failed: {ex.Message}\n{ex.StackTrace}");
            }
        }

        // NEW: Get current user SMTP from Application.Session (more reliable)
        // ✅ OPTIMIZED: Use cached email if available
        private string GetCurrentUserSmtpFromSession()
        {
            // Return cached value if available (99% of calls)
            if (!string.IsNullOrWhiteSpace(_cachedUserEmail))
                return _cachedUserEmail;

            Outlook.NameSpace session     = null;
            Outlook.Recipient currentUser = null;
            Outlook.AddressEntry ae       = null;
            try
            {
                session     = this.Application?.Session;
                currentUser = session?.CurrentUser;
                ae          = currentUser?.AddressEntry;
                if (ae != null)
                {
                    if ("EX".Equals(ae.Type, StringComparison.OrdinalIgnoreCase))
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        Outlook.PropertyAccessor pa = null;
                        try
                        {
                            pa = ae.PropertyAccessor;
                            var smtp = pa.GetProperty(PR_SMTP_ADDRESS) as string;
                            if (!string.IsNullOrWhiteSpace(smtp))
                            {
                                _cachedUserEmail = smtp;
                                return smtp;
                            }
                        }
                        finally { if (pa != null) Marshal.ReleaseComObject(pa); }
                    }
                    if (!string.IsNullOrWhiteSpace(ae.Address))
                    {
                        _cachedUserEmail = ae.Address;
                        return ae.Address;
                    }
                }
                var fallback = currentUser?.Name ?? string.Empty;
                _cachedUserEmail = fallback;
                return fallback;
            }
            catch { return string.Empty; }
            finally
            {
                if (ae          != null) Marshal.ReleaseComObject(ae);
                if (currentUser != null) Marshal.ReleaseComObject(currentUser);
                if (session     != null) Marshal.ReleaseComObject(session);
            }
        }

        private static T Safe<T>(Func<T> f)
        {
            try { return f(); } catch { return default(T); }
        }

        private static string GetCurrentUserSmtp(Outlook.AppointmentItem appt)
        {
            Outlook.NameSpace session     = null;
            Outlook.Recipient currentUser = null;
            Outlook.AddressEntry ae       = null;
            try
            {
                session     = appt?.Session;
                currentUser = session?.CurrentUser;
                ae          = currentUser?.AddressEntry;
                if (ae != null)
                {
                    if ("EX".Equals(ae.Type, StringComparison.OrdinalIgnoreCase))
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        Outlook.PropertyAccessor pa = null;
                        try
                        {
                            pa = ae.PropertyAccessor;
                            var smtp = pa.GetProperty(PR_SMTP_ADDRESS) as string;
                            if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                        }
                        finally { if (pa != null) Marshal.ReleaseComObject(pa); }
                    }
                    if (!string.IsNullOrWhiteSpace(ae.Address)) return ae.Address;
                }
                return currentUser?.Name ?? string.Empty;
            }
            catch { return string.Empty; }
            finally
            {
                if (ae          != null) Marshal.ReleaseComObject(ae);
                if (currentUser != null) Marshal.ReleaseComObject(currentUser);
                if (session     != null) Marshal.ReleaseComObject(session);
            }
        }

        // Get SMTP from a Recipient
        private static string GetSmtpFromRecipient(Outlook.Recipient recipient)
        {
            if (recipient == null) return string.Empty;
            Outlook.AddressEntry ae = null;
            try
            {
                ae = recipient.AddressEntry;
                if (ae != null)
                {
                    if ("EX".Equals(ae.Type, StringComparison.OrdinalIgnoreCase))
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        Outlook.PropertyAccessor pa = null;
                        try
                        {
                            pa = ae.PropertyAccessor;
                            var smtp = pa.GetProperty(PR_SMTP_ADDRESS) as string;
                            if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                        }
                        finally { if (pa != null) Marshal.ReleaseComObject(pa); }
                    }
                    if (!string.IsNullOrWhiteSpace(ae.Address)) return ae.Address;
                }
                return recipient.Address ?? string.Empty;
            }
            catch { return string.Empty; }
            finally { if (ae != null) Marshal.ReleaseComObject(ae); }
        }

        // Check if email is TTI domain
        private static bool IsTti(string email)
            => !string.IsNullOrWhiteSpace(email) &&
               email.EndsWith("@thetravellerinc.com", StringComparison.OrdinalIgnoreCase);

        // Helper method to get all recipients as semicolon-separated string
        private static string GetAllRecipients(Outlook.AppointmentItem appt)
        {
            if (appt == null) return string.Empty;

            // Check if this is a meeting (has recipients) or just a calendar appointment
            // For calendar appointments (no recipients), return empty string
            try
            {
                if (appt.MeetingStatus == Outlook.OlMeetingStatus.olNonMeeting)
                {
                    return string.Empty;  // Calendar appointment, no recipients
                }
            }
            catch
            {
                return string.Empty;  // If we can't determine meeting status, assume no recipients
            }

            var recipientEmails = new System.Collections.Generic.List<string>();

            Outlook.Recipients recipients = null;
            try
            {
                recipients = appt.Recipients;
                if (recipients != null && recipients.Count > 0)
                {
                    foreach (Outlook.Recipient recipient in recipients)
                    {
                        Outlook.Recipient captured = recipient;
                        try
                        {
                            var email = GetSmtpFromRecipient(captured);
                            if (!string.IsNullOrWhiteSpace(email))
                                recipientEmails.Add(email);
                        }
                        catch { }
                        finally { Marshal.ReleaseComObject(captured); }
                    }
                }

                // Add organizer if not already in list
                var organizerEmail = GetCurrentUserSmtp(appt);
                if (!string.IsNullOrWhiteSpace(organizerEmail) && !recipientEmails.Contains(organizerEmail))
                    recipientEmails.Insert(0, organizerEmail);
            }
            catch { }
            finally { if (recipients != null) Marshal.ReleaseComObject(recipients); }

            return string.Join("; ", recipientEmails);
        }

        // Public accessor so other classes (e.g. ManageTimesheetPane) can use the cached email
        // without making their own COM calls
        public string GetCachedUserEmail()
        {
            return !string.IsNullOrWhiteSpace(_cachedUserEmail)
                ? _cachedUserEmail
                : GetCurrentUserSmtpFromSession();
        }

        // ── Background helpers (replaces anonymous Task.Run lambdas) ──────────

        // Fix 4: Named async method — replaces Task.Run in CalendarItems_ItemChange
        private async System.Threading.Tasks.Task CalendarItems_ItemChange_BackgroundAsync(
            string source, string entryId, string globalIdFallback,
            string subject, DateTime startUtc, DateTime endUtc,
            string programCode, string activityCode, string stageCode,
            string currentUserEmail, DateTime lastMod, bool isRecurring)
        {
            var tempRec = new MeetingRecord
            {
                GlobalId = globalIdFallback,
                EntryId = entryId,
                StartUtc = startUtc,
                UserDisplayName = currentUserEmail
            };

            try
            {
                var existing = await DbWriter.GetExistingTimesheetAsync(tempRec).ConfigureAwait(false);
                if (existing == null)
                {
                    System.Diagnostics.Debug.WriteLine($"{source}: Skipping '{subject}' - no existing timesheet found");
                    return;
                }
                System.Diagnostics.Debug.WriteLine($"{source}: Existing timesheet found for '{subject}' - proceeding with update");
            }
            catch (Exception checkEx)
            {
                System.Diagnostics.Debug.WriteLine($"{source}: Failed to check existing timesheet: {checkEx.Message}");
                return;
            }

            var globalId = globalIdFallback;
            try
            {
                var ns = Globals.ThisAddIn.Application.Session;
                var bgAppt = ns.GetItemFromID(entryId) as Outlook.AppointmentItem;
                if (bgAppt != null)
                {
                    try   { globalId = Safe<string>(() => bgAppt.GlobalAppointmentID) ?? entryId; }
                    finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(bgAppt); }
                }
            }
            catch { /* use fallback */ }

            var rec = new MeetingRecord
            {
                Source      = source,
                EntryId     = entryId,
                GlobalId    = globalId,
                Subject     = subject,
                StartUtc    = startUtc,
                EndUtc      = endUtc,
                ProgramCode = programCode,
                ActivityCode = activityCode,
                StageCode   = stageCode,
                UserDisplayName = currentUserEmail,
                LastModifiedUtc = lastMod,
                IsRecurring = isRecurring,
                Recipients  = string.Empty
            };

            await DbWriter.UpsertAsync(rec).ConfigureAwait(false);
            System.Diagnostics.Debug.WriteLine($"{source}: Updated '{subject}' for {currentUserEmail}");
        }

        // Fix 4: Named async method — replaces Task.Run in meeting-cancellation path
        private async System.Threading.Tasks.Task ItemSend_ProcessCancellationAsync(
            string cancelGlobalId, string cancelEntryId, string cancelSubject)
        {
            try
            {
                var deleted = await DbWriter.DeleteAllTimesheetsByGlobalIdAsync(cancelGlobalId, cancelEntryId).ConfigureAwait(false);
                System.Diagnostics.Debug.WriteLine(deleted > 0
                    ? $"Deleted {deleted} timesheet record(s) for cancelled meeting: {cancelSubject}"
                    : $"No timesheet records found for cancelled meeting: {cancelSubject}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to delete timesheet on cancellation: {ex.Message}");
            }
        }

        // Fix 1 + 4: Named async method — replaces Task.Run in normal-send path AND
        //            moves the .Result existence-check here so the UI thread is never blocked
        private async System.Threading.Tasks.Task ItemSend_ProcessSendAsync(
            string entryId, string initialGlobalId, string subject,
            DateTime startUtc, DateTime endUtc, DateTime lastMod,
            string programCode, string activityCode, string stageCode,
            string organizerEmail)
        {
            try
            {
                // FIX 1: existence check is now fully async (was .Result on the UI thread)
                var tempRec = new MeetingRecord
                {
                    GlobalId        = initialGlobalId,
                    EntryId         = entryId,
                    StartUtc        = startUtc,
                    UserDisplayName = organizerEmail
                };

                var existing = await DbWriter.GetExistingTimesheetAsync(tempRec).ConfigureAwait(false);
                if (existing == null)
                {
                    System.Diagnostics.Debug.WriteLine($"ItemSend: Skipping auto-submit for '{subject}' - no prior submission found");
                    return;
                }

                System.Diagnostics.Debug.WriteLine($"ItemSend: Prior submission found for '{subject}' - proceeding with update");

                // Wait for the send to complete before re-accessing the appointment
                await System.Threading.Tasks.Task.Delay(500).ConfigureAwait(false);

                Outlook.AppointmentItem bgAppt = null;
                string globalId   = initialGlobalId;
                string recipients = "";
                int retryCount    = 0;
                const int maxRetries = 5;

                while (retryCount < maxRetries)
                {
                    try
                    {
                        var ns = Globals.ThisAddIn.Application.Session;
                        bgAppt = ns.GetItemFromID(entryId) as Outlook.AppointmentItem;
                        if (bgAppt != null)
                        {
                            globalId   = Safe<string>(() => bgAppt.GlobalAppointmentID) ?? entryId;
                            recipients = GetAllRecipients(bgAppt);
                            ApplyCategoryToAppointment(bgAppt);
                            System.Diagnostics.Debug.WriteLine($"ItemSend: Category applied after {retryCount} retries");
                            break;
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException comEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"ItemSend: COM exception retry {retryCount}: {comEx.Message}");
                        retryCount++;
                        if (retryCount < maxRetries)
                            await System.Threading.Tasks.Task.Delay(500 * retryCount).ConfigureAwait(false);
                        else
                            throw;
                    }
                    finally
                    {
                        if (bgAppt != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(bgAppt);
                            bgAppt = null;
                        }
                    }
                }

                var rec = new MeetingRecord
                {
                    Source          = "ItemSend",
                    EntryId         = entryId,
                    GlobalId        = globalId,
                    Subject         = subject,
                    StartUtc        = startUtc,
                    EndUtc          = endUtc,
                    ProgramCode     = programCode  ?? "",
                    ActivityCode    = activityCode ?? "",
                    StageCode       = stageCode    ?? "",
                    UserDisplayName = organizerEmail,
                    LastModifiedUtc = lastMod,
                    Recipients      = recipients
                };

                await DbWriter.UpsertAsync(rec).ConfigureAwait(false);
                System.Diagnostics.Debug.WriteLine($"ItemSend: Saved '{subject}' for {organizerEmail}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ItemSend_ProcessSendAsync failed: {ex.Message}");
            }
        }

        private const int PaneFixedWidth = 370;

        public void ShowManageTimesheetPane()
        {
            try
            {
                bool isNewPane = (_manageTimesheetPane == null);

                if (isNewPane)
                {
                    var paneControl = new ManageTimesheetPane();
                    _manageTimesheetPane = this.CustomTaskPanes.Add(paneControl, "Manage Timesheet");
                    _manageTimesheetPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                    _manageTimesheetPane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
                    _manageTimesheetPane.Width = PaneFixedWidth;

                    // Debounced resize guard: snap back to fixed width after user stops dragging.
                    // Using a timer avoids the re-entrant SizeChanged cascade that caused >1 min freeze.
                    _paneWidthTimer = new System.Windows.Forms.Timer { Interval = 150 };
                    _paneWidthTimer.Tick += (ts, te) =>
                    {
                        _paneWidthTimer.Stop();
                        if (_manageTimesheetPane != null && _manageTimesheetPane.Width != PaneFixedWidth)
                            _manageTimesheetPane.Width = PaneFixedWidth;
                    };

                    paneControl.Resize += (s, e) =>
                    {
                        // Restart timer on every drag tick — fires once user releases
                        _paneWidthTimer.Stop();
                        _paneWidthTimer.Start();
                    };
                }

                _manageTimesheetPane.Visible = true;

                // Load data AFTER Visible = true — at this point the control is fully
                // hosted inside the VSTO CustomTaskPane and its window handle exists,
                // so all Invoke calls inside LoadDataAsync are safe.
                var ctrl = _manageTimesheetPane.Control as ManageTimesheetPane;
                if (ctrl != null)
                    _ = ctrl.LoadDataAsync();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to show Manage Timesheet pane: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"Failed to open Manage Timesheet: {ex.Message}", "Error");
            }
        }
    }
}