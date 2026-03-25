using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using OutlookAddIn1;

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
            var appt = Item as Outlook.AppointmentItem;
            if (appt != null)
                ExportProgramMeta(appt, "ItemAdd");
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

                // ✅ OPTIMIZED: Fire-and-forget background update (fully non-blocking)
                System.Threading.Tasks.Task.Run(async () =>
                {
                    // ✅ CRITICAL FIX: Only update if timesheet already exists in database
                    // This prevents auto-submitting NEW meetings created via "New Online Meeting"
                    var tempRec = new MeetingRecord
                    {
                        GlobalId = globalIdFallback,
                        EntryId = entryId,
                        StartUtc = startUtc,
                        UserDisplayName = currentUserEmail
                    };

                    try
                    {
                        var existing = await DbWriter.GetExistingTimesheetAsync(tempRec);
                        if (existing == null)
                        {
                            System.Diagnostics.Debug.WriteLine($"{source}: Skipping '{subject}' - no existing timesheet found (new meeting must be manually submitted)");
                            return; // ✅ EXIT: Don't auto-submit new meetings!
                        }

                        System.Diagnostics.Debug.WriteLine($"{source}: Existing timesheet found for '{subject}' - proceeding with update");
                    }
                    catch (Exception checkEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"{source}: Failed to check existing timesheet: {checkEx.Message}");
                        return; // ✅ If we can't check, don't risk auto-submitting
                    }

                    // ✅ Get GlobalID in background (expensive for Teams/Zoom meetings)
                    var globalId = globalIdFallback;
                    try
                    {
                        var ns = Globals.ThisAddIn.Application.Session;
                        var bgAppt = ns.GetItemFromID(entryId) as Outlook.AppointmentItem;
                        if (bgAppt != null)
                        {
                            try
                            {
                                globalId = Safe<string>(() => bgAppt.GlobalAppointmentID) ?? entryId;
                            }
                            finally
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(bgAppt);
                            }
                        }
                    }
                    catch { /* use fallback */ }

                    var rec = new MeetingRecord
                    {
                        Source = source,
                        EntryId = entryId,
                        GlobalId = globalId,
                        Subject = subject,
                        StartUtc = startUtc,
                        EndUtc = endUtc,
                        ProgramCode = programCode,
                        ActivityCode = activityCode,
                        StageCode = stageCode,
                        UserDisplayName = currentUserEmail,
                        LastModifiedUtc = lastMod,
                        IsRecurring = isRecurring,
                        Recipients = string.Empty  // Skip recipients for performance
                    };

                    await DbWriter.UpsertAsync(rec);
                    System.Diagnostics.Debug.WriteLine($"{source}: Updated {subject} for {currentUserEmail}");
                });
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
                            System.Threading.Tasks.Task.Run(async () =>
                            {
                                try
                                {
                                    var deleted = await DbWriter.DeleteAllTimesheetsByGlobalIdAsync(cancelGlobalId, cancelEntryId);
                                    if (deleted > 0)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"Deleted {deleted} timesheet record(s) for cancelled meeting: {cancelSubject}");
                                    }
                                    else
                                    {
                                        System.Diagnostics.Debug.WriteLine($"No timesheet records found for cancelled meeting: {cancelSubject}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Failed to delete timesheet on cancellation: {ex.Message}");
                                }
                            });
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

                    // ✅ CRITICAL: Check if timesheet already exists in database
                    // Only auto-submit on ItemSend if the user has ALREADY submitted via "Submit Timesheet"
                    bool hasExistingTimesheet = false;
                    try
                    {
                        var tempRec = new MeetingRecord
                        {
                            GlobalId = Safe<string>(() => appt2.GlobalAppointmentID) ?? entryId,
                            EntryId = entryId,
                            StartUtc = startUtc,
                            UserDisplayName = organizerEmail
                        };

                        // ✅ Use Task.Result for synchronous call in non-async method
                        var existing = DbWriter.GetExistingTimesheetAsync(tempRec).Result;

                        hasExistingTimesheet = existing != null;

                        if (!hasExistingTimesheet)
                        {
                            System.Diagnostics.Debug.WriteLine($"ItemSend: Skipping auto-submit for new meeting '{subject}' - user must manually submit timesheet");
                            return; // ✅ EXIT: Don't auto-submit new meetings!
                        }

                        System.Diagnostics.Debug.WriteLine($"ItemSend: Existing timesheet found for '{subject}' - proceeding with update");
                    }
                    catch (Exception checkEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"ItemSend: Failed to check existing timesheet: {checkEx.Message}");
                        return; // ✅ If we can't check, don't risk auto-submitting
                    }

                    // ✅ FIX: Capture ONLY fast properties
                    var endUtc = appt2.EndUTC;
                    var lastMod = appt2.LastModificationTime.ToUniversalTime();

                    // ✅ CRITICAL: Clear ProcessOnSend flag IMMEDIATELY
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

                    // ✅ Store variables for background processing
                    var entryIdForBackground = entryId;
                    var subjectForLog = subject;
                    var programCodeCopy = programCode;
                    var activityCodeCopy = activityCode;
                    var stageCodeCopy = stageCode;

                    // ✅ FIX: Use async/await pattern with proper timing
                    System.Threading.Tasks.Task.Run(async () =>
                    {
                        try
                        {
                            // ✅ Wait for send to complete - check multiple times with increasing delays
                            await System.Threading.Tasks.Task.Delay(500); // Initial short delay

                            Outlook.AppointmentItem bgAppt = null;
                            string globalId = "";
                            string recipients = "";
                            int retryCount = 0;
                            const int maxRetries = 5;

                            while (retryCount < maxRetries)
                            {
                                try
                                {
                                    var ns = Globals.ThisAddIn.Application.Session;
                                    bgAppt = ns.GetItemFromID(entryIdForBackground) as Outlook.AppointmentItem;

                                    if (bgAppt != null)
                                    {
                                        // ✅ Check if appointment is in a stable state (has been saved after send)
                                        globalId = Safe<string>(() => bgAppt.GlobalAppointmentID) ?? entryIdForBackground;
                                        recipients = GetAllRecipients(bgAppt);

                                        // ✅ Apply category
                                        ApplyCategoryToAppointment(bgAppt);

                                        // Success - exit retry loop
                                        System.Diagnostics.Debug.WriteLine($"ItemSend: Category applied successfully after {retryCount} retries");
                                        break;
                                    }
                                }
                                catch (System.Runtime.InteropServices.COMException comEx)
                                {
                                    // ✅ COM exception might mean appointment not ready yet
                                    System.Diagnostics.Debug.WriteLine($"ItemSend: COM exception on retry {retryCount}: {comEx.Message}");
                                    retryCount++;

                                    if (retryCount < maxRetries)
                                    {
                                        await System.Threading.Tasks.Task.Delay(500 * retryCount); // Exponential backoff
                                    }
                                    else
                                    {
                                        throw; // Give up after max retries
                                    }
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

                            // ✅ Save to database (after category is applied)
                            var rec = new MeetingRecord
                            {
                                Source = "ItemSend",
                                EntryId = entryIdForBackground,
                                GlobalId = globalId,
                                Subject = subjectForLog,
                                StartUtc = startUtc,
                                EndUtc = endUtc,
                                ProgramCode = programCodeCopy ?? "",
                                ActivityCode = activityCodeCopy ?? "",
                                StageCode = stageCodeCopy ?? "",
                                UserDisplayName = organizerEmail,
                                LastModifiedUtc = lastMod,
                                Recipients = recipients
                            };

                            await DbWriter.UpsertAsync(rec);
                            System.Diagnostics.Debug.WriteLine($"ItemSend: Saved {subjectForLog} for {organizerEmail}");
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"ItemSend background failed: {ex.Message}");
                        }
                    });
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

        // ✅ NEW: Helper method to apply category (called in background)
        private void ApplyCategoryToAppointment(Outlook.AppointmentItem appt)
        {
            try
            {
                var categoryName = "Timesheet Submitted";
                var categoryColor = Outlook.OlCategoryColor.olCategoryColorPeach;

                System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): About to apply category");
                System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): appt.Subject={appt.Subject}");
                System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): appt.EntryID={appt.EntryID}");
                System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): Current appt.Categories='{appt.Categories}'");

                // Check if category exists, create if not
                var categories = this.Application.Session.Categories;
                var existingCategory = System.Linq.Enumerable.FirstOrDefault(
                    System.Linq.Enumerable.Cast<Outlook.Category>(categories),
                    c => c.Name == categoryName);

                if (existingCategory == null)
                {
                    System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): Category '{categoryName}' does not exist, creating with Peach color");
                    categories.Add(categoryName, categoryColor);
                    System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): Category created successfully");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): Category '{categoryName}' already exists with color {existingCategory.Color}");

                    // ✅ FIX: If category exists with wrong color, delete and recreate it
                    if (existingCategory.Color != categoryColor)
                    {
                        System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): Category has wrong color ({existingCategory.Color}), deleting and recreating with Peach");
                        categories.Remove(categoryName);
                        categories.Add(categoryName, categoryColor);
                        System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): Category recreated with Peach color");
                    }
                }

                // Remove any existing timesheet categories first
                if (!string.IsNullOrEmpty(appt.Categories))
                {
                    var existingCategories = appt.Categories.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(c => c.Trim())
                        .Where(c => !c.Equals("Timesheet Submitted", StringComparison.OrdinalIgnoreCase) &&
                                   !c.Equals("Timesheet Ignored", StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    existingCategories.Add(categoryName);
                    appt.Categories = string.Join(", ", existingCategories);
                    System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): Updated categories (had existing): '{appt.Categories}'");
                }
                else
                {
                    appt.Categories = categoryName;
                    System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): Set categories (was empty): '{appt.Categories}'");
                }

                appt.Save();
                System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): appt.Save() completed. Final categories='{appt.Categories}'");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): FAILED! Exception: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment (ItemSend): Stack trace: {ex.StackTrace}");
            }
        }

        // NEW: Get current user SMTP from Application.Session (more reliable)
        // ✅ OPTIMIZED: Use cached email if available
        private string GetCurrentUserSmtpFromSession()
        {
            // Return cached value if available (99% of calls)
            if (!string.IsNullOrWhiteSpace(_cachedUserEmail))
                return _cachedUserEmail;

            try
            {
                var session = this.Application?.Session;
                var ae = session?.CurrentUser?.AddressEntry;
                if (ae != null)
                {
                    if ("EX".Equals(ae.Type, StringComparison.OrdinalIgnoreCase))
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        var smtp = ae.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                        if (!string.IsNullOrWhiteSpace(smtp))
                        {
                            _cachedUserEmail = smtp; // Cache for future calls
                            return smtp;
                        }
                    }
                    if (!string.IsNullOrWhiteSpace(ae.Address))
                    {
                        _cachedUserEmail = ae.Address; // Cache for future calls
                        return ae.Address;
                    }
                }
                var fallback = session?.CurrentUser?.Name ?? string.Empty;
                _cachedUserEmail = fallback; // Cache for future calls
                return fallback;
            }
            catch { return string.Empty; }
        }

        private static void ExportProgramMeta(Outlook.AppointmentItem appt, string source)
        {
            if (appt == null) return;

            // Read your custom properties
            var program = GetUP(appt, "ProgramCode");
            var activity = GetUP(appt, "ActivityCode");
            var stage = GetUP(appt, "StageCode");

            // Fallback: parse from subject like "[KIA-0523] ..."
            if (string.IsNullOrWhiteSpace(program))
                program = ExtractSubjectCode(appt.Subject);

            //// Build a compact record
            //var record = new MeetingRecord
            //{
            //    Source = source,
            //    EntryId = appt.EntryID,                           // stable after save
            //    //GlobalId = Safe<string>(() => appt.GlobalAppointmentID),
            //    Subject = appt.Subject ?? "",
            //    StartUtc = appt.StartUTC,
            //    EndUtc = appt.EndUTC,
            //    ProgramCode = program ?? "",
            //    ActivityCode = activity ?? "",
            //    StageCode = stage ?? ""
            //    // LastModified = appt.LastModificationTime
            //};

            //// === Transport: CSV for POC ===
            //EventIo.LogCsv(record);


        }

        private static string GetUP(Outlook.AppointmentItem appt, string name)
        {
            try
            {
                var ups = appt.UserProperties;
                var up = ups != null ? ups.Find(name) : null;
                return up != null && up.Value != null ? up.Value.ToString() : "";
            }
            catch { return ""; }
        }

        private static string ExtractSubjectCode(string subject)
        {
            if (string.IsNullOrEmpty(subject)) return "";
            var m = Regex.Match(subject, @"\[(?<code>[^\[\]]+)\]");
            return m.Success ? m.Groups["code"].Value : "";
        }

        private static T Safe<T>(Func<T> f)
        {
            try { return f(); } catch { return default(T); }
        }

        //private static readonly string CsvPath = Path.Combine(
        //    Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
        //    "meeting_events.csv");

        private static string GetCurrentUserSmtp(Outlook.AppointmentItem appt)
        {
            try
            {
                var ae = appt?.Session?.CurrentUser?.AddressEntry;
                if (ae != null)
                {
                    if ("EX".Equals(ae.Type, StringComparison.OrdinalIgnoreCase))
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        var smtp = ae.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                        if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                    }
                    if (!string.IsNullOrWhiteSpace(ae.Address)) return ae.Address;
                }
                return appt?.Session?.CurrentUser?.Name ?? string.Empty;
            }
            catch { return string.Empty; }
        }

        // Get SMTP from a Recipient
        private static string GetSmtpFromRecipient(Outlook.Recipient recipient)
        {
            if (recipient == null) return string.Empty;
            try
            {
                var ae = recipient.AddressEntry;
                if (ae != null)
                {
                    if ("EX".Equals(ae.Type, StringComparison.OrdinalIgnoreCase))
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        var smtp = ae.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                        if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                    }
                    if (!string.IsNullOrWhiteSpace(ae.Address)) return ae.Address;
                }
                return recipient.Address ?? string.Empty;
            }
            catch { return string.Empty; }
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

            try
            {
                // Get all recipients from the meeting
                if (appt.Recipients != null && appt.Recipients.Count > 0)
                {
                    foreach (Outlook.Recipient recipient in appt.Recipients)
                    {
                        try
                        {
                            var email = GetSmtpFromRecipient(recipient);
                            if (!string.IsNullOrWhiteSpace(email))
                            {
                                recipientEmails.Add(email);
                            }
                        }
                        catch { }
                    }
                }

                // Add organizer if not already in list
                var organizerEmail = GetCurrentUserSmtp(appt);
                if (!string.IsNullOrWhiteSpace(organizerEmail) && !recipientEmails.Contains(organizerEmail))
                {
                    recipientEmails.Insert(0, organizerEmail); // Organizer first
                }
            }
            catch { }

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

        private const int PaneFixedWidth = 370;

        public void ShowManageTimesheetPane()
        {
            try
            {
                if (_manageTimesheetPane == null)
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
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to show Manage Timesheet pane: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"Failed to open Manage Timesheet: {ex.Message}", "Error");
            }
        }
    }
}