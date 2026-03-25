using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1
{
    [ComVisible(true)]
    public class CalendarRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        private string ExplorerTabsXml() => @"
  <tab idMso='TabCalendar'>
    <group id='grpTTIMeetings' label='TTI Specific' insertAfterMso='GroupHelp'>
      <button id='btnManageTimesheet' label='Manage Timesheet' size='large' getImage='GetManageTimesheetImage' onAction='OnManageTimesheet'/>
    </group>
  </tab>";

        private string ExplorerContextMenusXml() => @"
  <contextMenus>
    <contextMenu idMso='ContextMenuCalendarItem'>
      <button id='SubmitTimesheet' label='Edit Timesheet' onAction='OnSubmitTimesheet'/>
      <button id='CancelTimesheet' label='Cancel Timesheet Submission' onAction='OnCancelTimesheet'/>
    </contextMenu>
  </contextMenus>";

        public string GetCustomUI(string ribbonID)
        {
            System.Diagnostics.Debug.WriteLine($"GetCustomUI called with ribbonID='{ribbonID}'");

            if (!string.IsNullOrEmpty(ribbonID) &&
                ribbonID.StartsWith("Microsoft.Outlook.Explorer", StringComparison.Ordinal))
            {
                var xml = $@"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon><tabs>{ExplorerTabsXml()}</tabs></ribbon>
  {ExplorerContextMenusXml()}
</customUI>";
                System.Diagnostics.Debug.WriteLine($"GetCustomUI: Returning Explorer XML ({xml.Length} chars)");
                return xml;
            }

            return null;
        }


        public void OnRibbonLoad(Office.IRibbonUI ribbonUI) { _ribbon = ribbonUI; }

        public Bitmap GetManageTimesheetImage(Office.IRibbonControl control)
        {
            try
            {
                var asm = System.Reflection.Assembly.GetExecutingAssembly();
                using (var stream = asm.GetManifestResourceStream("OutlookAddIn1.Resources.ManageTimesheet.png"))
                {
                    if (stream != null)
                        return new Bitmap(stream);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetManageTimesheetImage failed: {ex.Message}");
            }
            return null;
        }

        public void OnManageTimesheet(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.ShowManageTimesheetPane();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to open Manage Timesheet: " + ex.Message, "Error");
            }
        }

        public async void OnSubmitTimesheet(Office.IRibbonControl control)
        {
            Outlook.Explorer exp = null;
            Outlook.Inspector insp = null;
            Outlook.AppointmentItem appt = null;

            try
            {
                var app = Globals.ThisAddIn.Application;

                exp = app.ActiveExplorer();

                if (exp != null && exp.Selection != null && exp.Selection.Count > 0)
                {
                    appt = exp.Selection[1] as Outlook.AppointmentItem;
                }

                if (appt == null)
                {
                    insp = app.ActiveInspector();
                    if (insp != null)
                        appt = insp.CurrentItem as Outlook.AppointmentItem;
                }

                if (appt == null)
                {
                    MessageBox.Show("No calendar item selected.", "Submit Timesheet");
                    return;
                }

                var email = GetCurrentUserEmailAddress();

                var tempRec = new MeetingRecord
                {
                    GlobalId = GetGlobalId(appt),
                    EntryId = appt.EntryID ?? "",
                    StartUtc = appt.StartUTC,
                    UserDisplayName = email
                };

                // ✅ Check if timesheet already exists and get ALL records for this meeting
                var existingTimesheets = await DbWriter.GetAllTimesheetsForMeetingAsync(tempRec);

                if (existingTimesheets != null && existingTimesheets.Count > 0)
                {
                    // ✅ Check if this is a multi-program submission
                    if (existingTimesheets.Count > 1)
                    {
                        // Multi-program timesheet exists
                        var torontoTime = existingTimesheets[0].LastModifiedTorontoTime;
                        var programList = string.Join("\n", existingTimesheets.Select(t =>
                            $"  • {t.ProgramCode} ({(t.HoursAllocated ?? (t.EndUtc - t.StartUtc).TotalHours):F2} hrs)"));

                        var dialogResult = MessageBox.Show(
                            $"Multi-program timesheet already submitted for this meeting.\n\n" +
                            $"Previously submitted:\n{programList}\n" +
                            $"Submitted: {torontoTime:yyyy-MM-dd HH:mm} (Toronto Time)\n\n" +
                            $"Would you like to re-submit this timesheet?\n" +
                            $"(This will DELETE the existing entries and allow you to create new ones)",
                            "Multi-Program Timesheet Exists",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (dialogResult != DialogResult.Yes)
                        {
                            return;
                        }

                        // ✅ CRITICAL: Delete all existing records for this meeting before re-submission
                        foreach (var existing in existingTimesheets)
                        {
                            await DbWriter.DeleteTimesheetAsync(existing);
                        }
                        System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Deleted {existingTimesheets.Count} existing records before re-submission");
                    }
                    else
                    {
                        // Single-program timesheet exists
                        var existingTimesheet = existingTimesheets[0];
                        var torontoTime = existingTimesheet.LastModifiedTorontoTime;

                        var dialogResult = MessageBox.Show(
                            $"Timesheet already submitted for this meeting.\n\n" +
                            $"Previously submitted:\n" +
                            $"Program:  {existingTimesheet.ProgramCode}\n" +
                            $"Activity: {existingTimesheet.ActivityCode}\n" +
                            $"Stage:    {existingTimesheet.StageCode}\n" +
                            $"Submitted: {torontoTime:yyyy-MM-dd HH:mm} (Toronto Time)\n\n" +
                            $"Would you like to edit this timesheet?",
                            "Timesheet Already Submitted",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (dialogResult != DialogResult.Yes)
                        {
                            return;
                        }

                        // ✅ CRITICAL: Delete existing record before re-submission to prevent duplicates
                        await DbWriter.DeleteTimesheetAsync(existingTimesheet);
                        System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Deleted existing record before re-submission");
                    }
                }

                // ✅ Get initial values for the dialog (from first existing record or UserProperties)
                var firstExistingTimesheet = existingTimesheets?.FirstOrDefault();
                string currProgram = firstExistingTimesheet?.ProgramCode ?? GetUP(appt, "ProgramCode");
                string currActivity = firstExistingTimesheet?.ActivityCode ?? GetUP(appt, "ActivityCode");
                string currStage = firstExistingTimesheet?.StageCode ?? GetUP(appt, "StageCode");

                // Calculate meeting duration in hours
                var duration = (appt.EndUTC - appt.StartUTC).TotalHours;

                // ✅ DECLARE VARIABLES FOR BOTH SINGLE AND MULTI-PROGRAM MODES
                var entryId = appt.EntryID ?? "";
                var globalId = GetGlobalId(appt);
                var subject = appt.Subject ?? "";
                var startUtc = appt.StartUTC;
                var endUtc = appt.EndUTC;
                var lastModified = appt.LastModificationTime.ToUniversalTime();
                var isRecurring = appt.IsRecurring;
                var source = firstExistingTimesheet != null ? "SubmitTimesheet_Update" : "SubmitTimesheet";

                using (var dlg = new ProgramPickerForm(currProgram, currActivity, currStage, duration))
                {
                    if (dlg.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    // ✅ CHECK IF MULTI-PROGRAM MODE IS ACTIVATED
                    // NEW: Now includes ORIGINAL program + ADDITIONAL programs
                    if (dlg.IsMultiProgram && dlg.ProgramAllocations.Count >= 1)
                    {
                        // MULTI-PROGRAM MODE: SUBMIT ORIGINAL + ADDITIONAL PROGRAMS
                        var allocations = new System.Collections.Generic.List<ProgramAllocation>();

                        // ✅ NEW: Calculate original program hours (remainder after additional programs)
                        double additionalHours = dlg.ProgramAllocations.Sum(p => p.Hours);
                        double originalProgramHours = duration - additionalHours;

                        // ✅ NEW: Add ORIGINAL program as FIRST allocation
                        allocations.Add(new ProgramAllocation
                        {
                            ProgramCode = dlg.ProgramCode,
                            ActivityCode = dlg.ActivityCode,
                            StageCode = dlg.StageCode,
                            Hours = originalProgramHours
                        });

                        // ✅ NEW: Add ADDITIONAL programs
                        allocations.AddRange(dlg.ProgramAllocations);

                        // Get recipients once for all allocations
                        string recipients = "";
                        try
                        {
                            recipients = GetAllRecipients(appt);
                        }
                        catch (Exception recipEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to get recipients: {recipEx.Message}");
                        }

                        foreach (var allocation in allocations)
                        {
                            var rec = new MeetingRecord
                            {
                                Source = source,
                                EntryId = entryId,
                                GlobalId = globalId,
                                Subject = subject,
                                StartUtc = startUtc,               // ✅ Keep actual meeting start
                                EndUtc = endUtc,                   // ✅ Keep actual meeting end
                                HoursAllocated = allocation.Hours, // ✅ NEW: Explicit allocation
                                ProgramCode = allocation.ProgramCode,
                                ActivityCode = allocation.ActivityCode,
                                StageCode = allocation.StageCode,
                                UserDisplayName = email,
                                LastModifiedUtc = lastModified,
                                IsRecurring = isRecurring,
                                Recipients = recipients,           // ✅ FIXED: Include recipients
                                Status = "submitted"
                            };

                            await DbWriter.UpsertAsync(rec);
                        }

                        System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Multi-program DbWriter.UpsertAsync completed successfully for {allocations.Count} programs (1 original + {dlg.ProgramAllocations.Count} additional)");

                        // Apply peach category
                        try
                        {
                            // ✅ CRITICAL FIX: Skip category for recurring meeting occurrences
                            if (appt.RecurrenceState == Outlook.OlRecurrenceState.olApptOccurrence)
                            {
                                System.Diagnostics.Debug.WriteLine($"⚠️ SKIPPING category for recurring meeting occurrence (would affect entire series)");
                                System.Diagnostics.Debug.WriteLine($"   Database status is saved correctly for this specific occurrence");
                            }
                            else
                            {
                                var categoryName = "Timesheet Submitted";
                                var categoryColor = Outlook.OlCategoryColor.olCategoryColorPeach;

                                System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: About to apply category. Current appt.Categories='{appt.Categories}'");
                                System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: appt.EntryID={appt.EntryID}");
                                System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: RecurrenceState={appt.RecurrenceState}");

                                // Check if category exists, if not create it
                                var categories = Globals.ThisAddIn.Application.Session.Categories;
                                var existingCategory = categories.Cast<Outlook.Category>()
                                    .FirstOrDefault(c => c.Name == categoryName);

                                if (existingCategory == null)
                                {
                                    System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Category '{categoryName}' does not exist, creating with Peach color");
                                    categories.Add(categoryName, categoryColor);
                                    System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Category created successfully");
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Category '{categoryName}' already exists with color {existingCategory.Color}");

                                    // ✅ FIX: If category exists with wrong color, delete and recreate it
                                    if (existingCategory.Color != categoryColor)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Category has wrong color ({existingCategory.Color}), deleting and recreating with Peach");
                                        categories.Remove(categoryName);
                                        categories.Add(categoryName, categoryColor);
                                        System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Category recreated with Peach color");
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
                                    System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Updated categories (had existing): '{appt.Categories}'");
                                }
                                else
                                {
                                    appt.Categories = categoryName;
                                    System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Set categories (was empty): '{appt.Categories}'");
                                }

                                appt.Save();
                                System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: appt.Save() completed. Final categories='{appt.Categories}'");
                            }
                        }
                        catch (Exception catEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: FAILED to apply category! Exception: {catEx.Message}");
                            System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Stack trace: {catEx.StackTrace}");
                        }

                        MessageBox.Show(
    $"Timesheet submitted for {allocations.Count} programs!\n\n" +
    $"Original Program:\n  • {allocations[0].ProgramCode}: {allocations[0].Hours:F2} hrs\n\n" +
    $"Additional Programs:\n" +
    string.Join("\n", allocations.Skip(1).Select(a => $"  • {a.ProgramCode}: {a.Hours:F2} hrs")),
    "Multi-Program Timesheet Submitted",
    MessageBoxButtons.OK,
    MessageBoxIcon.Information);
                    }
                    else
                    {
                        // SINGLE-PROGRAM MODE: Original behavior (also handles when checkbox is checked but only 1 program)
                        // Get the program data - either from the single-program controls or the first allocation
                        string programCode, activityCode, stageCode;

                        if (dlg.IsMultiProgram && dlg.ProgramAllocations.Count == 1)
                        {
                            // Checkbox was checked but only one program - use allocation data
                            var singleAllocation = dlg.ProgramAllocations[0];
                            programCode = singleAllocation.ProgramCode ?? "";
                            activityCode = singleAllocation.ActivityCode ?? "";
                            stageCode = singleAllocation.StageCode ?? "";
                        }
                        else
                        {
                            // Normal single-program mode
                            programCode = dlg.ProgramCode ?? "";
                            activityCode = dlg.ActivityCode ?? "";
                            stageCode = dlg.StageCode ?? "";
                        }

                        // Update UserProperties ONLY - DON'T modify body
                        var ups = appt.UserProperties;
                        AddOrSetTextProp(ups, "ProgramCode", programCode);
                        AddOrSetTextProp(ups, "ActivityCode", activityCode);
                        AddOrSetTextProp(ups, "StageCode", stageCode);

                        var apptForRecipients = appt;

                        System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Starting background save for {email}");

                        await System.Threading.Tasks.Task.Run(async () =>
                        {
                            string recipients = "";
                            try
                            {
                                recipients = GetAllRecipients(apptForRecipients);
                            }
                            catch (Exception recipEx)
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed to get recipients: {recipEx.Message}");
                            }

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
                                UserDisplayName = email,
                                LastModifiedUtc = lastModified,
                                IsRecurring = isRecurring,
                                Recipients = recipients
                            };

                            await DbWriter.UpsertAsync(rec);
                        });

                        System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: DbWriter.UpsertAsync completed successfully");

                        // ✅ CRITICAL: Apply category AFTER successful database save
                        try
                        {
                            // ✅ CRITICAL FIX: Skip category for recurring meeting occurrences
                            if (appt.RecurrenceState == Outlook.OlRecurrenceState.olApptOccurrence)
                            {
                                System.Diagnostics.Debug.WriteLine($"⚠️ SKIPPING category for recurring meeting occurrence (would affect entire series)");
                                System.Diagnostics.Debug.WriteLine($"   Database status is saved correctly for this specific occurrence");
                            }
                            else
                            {
                                var categoryName = "Timesheet Submitted";
                                var categoryColor = Outlook.OlCategoryColor.olCategoryColorPeach;

                                System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: About to apply category. Current appt.Categories='{appt.Categories}'");
                                System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: appt.EntryID={appt.EntryID}");

                                // Check if category exists, if not create it
                                var categories = Globals.ThisAddIn.Application.Session.Categories;
                                var existingCategory = categories.Cast<Outlook.Category>()
                                    .FirstOrDefault(c => c.Name == categoryName);


                                if (existingCategory == null)
                                {
                                    System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Category '{categoryName}' does not exist, creating with Peach color");
                                    categories.Add(categoryName, categoryColor);
                                    System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Category created successfully");
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Category '{categoryName}' already exists with color {existingCategory.Color}");

                                    // ✅ FIX: If category exists with wrong color, delete and recreate it
                                    if (existingCategory.Color != categoryColor)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Category has wrong color ({existingCategory.Color}), deleting and recreating with Peach");
                                        categories.Remove(categoryName);
                                        categories.Add(categoryName, categoryColor);
                                        System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Category recreated with Peach color");
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
                                    System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Updated categories (had existing): '{appt.Categories}'");
                                }
                                else
                                {
                                    appt.Categories = categoryName;
                                    System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Set categories (was empty): '{appt.Categories}'");
                                }

                                appt.Save();
                                System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: appt.Save() completed. Final categories='{appt.Categories}'");
                            }
                        }
                        catch (Exception catEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: FAILED to apply category! Exception: {catEx.Message}");
                            System.Diagnostics.Debug.WriteLine($"OnSubmitTimesheet: Stack trace: {catEx.StackTrace}");
                        }

                        var actionText = firstExistingTimesheet != null ? "updated" : "submitted";
                        MessageBox.Show(
                            $"Timesheet {actionText} successfully!\n\n" +
                            $"Program:  {programCode}\n" +
                            $"Activity: {activityCode}\n" +
                            $"Stage:    {stageCode}",
                            $"Timesheet {(firstExistingTimesheet != null ? "Updated" : "Submitted")}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Submit Timesheet failed: " + ex.Message + "\n\nStack: " + ex.StackTrace, "Error");
            }
            finally
            {
                // Release COM objects
                if (appt != null)
                {
                    Marshal.ReleaseComObject(appt);
                    appt = null;
                }
                if (insp != null)
                {
                    Marshal.ReleaseComObject(insp);
                    insp = null;
                }
                if (exp != null)
                {
                    Marshal.ReleaseComObject(exp);
                    exp = null;
                }
            }
        }

        // === CANCEL TIMESHEET SUBMISSION ===
        public async void OnCancelTimesheet(Office.IRibbonControl control)
        {
            Outlook.Explorer exp = null;
            Outlook.Inspector insp = null;
            Outlook.AppointmentItem appt = null;

            try
            {
                var app = Globals.ThisAddIn.Application;

                exp = app.ActiveExplorer();

                if (exp != null && exp.Selection != null && exp.Selection.Count > 0)
                {
                    appt = exp.Selection[1] as Outlook.AppointmentItem;
                }

                if (appt == null)
                {
                    insp = app.ActiveInspector();
                    if (insp != null)
                        appt = insp.CurrentItem as Outlook.AppointmentItem;
                }

                if (appt == null)
                {
                    MessageBox.Show("No calendar item selected.", "Cancel Timesheet");
                    return;
                }

                var email = GetCurrentUserEmailAddress();

                var tempRec = new MeetingRecord
                {
                    GlobalId = GetGlobalId(appt),
                    EntryId = appt.EntryID ?? "",
                    StartUtc = appt.StartUTC,
                    UserDisplayName = email
                };

                // ✅ CRITICAL FIX: Get ALL timesheet records for this meeting (handles multi-program submissions)
                var existingTimesheets = await DbWriter.GetAllTimesheetsForMeetingAsync(tempRec);

                if (existingTimesheets == null || existingTimesheets.Count == 0)
                {
                    MessageBox.Show("No timesheet submission found for this meeting.", "Cancel Timesheet");
                    return;
                }

                // ✅ NEW: Show different messages for single vs multi-program submissions
                string confirmMessage;
                if (existingTimesheets.Count > 1)
                {
                    // Multi-program submission
                    var programList = string.Join("\n", existingTimesheets.Select(t =>
                        $"  • {t.ProgramCode}: {(t.HoursAllocated ?? (t.EndUtc - t.StartUtc).TotalHours):F2} hrs"));

                    confirmMessage =
                        $"Are you sure you want to cancel this MULTI-PROGRAM timesheet submission?\n\n" +
                        $"This will remove ALL {existingTimesheets.Count} program records:\n{programList}\n\n" +
                        $"This will remove all records from the database and clear the category from the meeting.";
                }
                else
                {
                    // Single-program submission
                    var existingTimesheet = existingTimesheets[0];
                    var torontoTime = existingTimesheet.LastModifiedTorontoTime;

                    confirmMessage =
                        $"Are you sure you want to cancel this timesheet submission?\n\n" +
                        $"Current submission:\n" +
                        $"Program:  {existingTimesheet.ProgramCode}\n" +
                        $"Activity: {existingTimesheet.ActivityCode}\n" +
                        $"Stage:    {existingTimesheet.StageCode}\n" +
                        $"Submitted: {torontoTime:yyyy-MM-dd HH:mm} (Toronto Time)\n\n" +
                        $"This will remove the record from the database and the category from the meeting.";
                }

                var dialogResult = MessageBox.Show(
                    confirmMessage,
                    "Confirm Cancellation",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (dialogResult != DialogResult.Yes)
                {
                    return;
                }

                // ✅ CRITICAL FIX: Delete ALL timesheet records for this meeting
                int deletedCount = 0;
                foreach (var record in existingTimesheets)
                {
                    var deleted = await DbWriter.DeleteTimesheetAsync(record);
                    if (deleted) deletedCount++;
                }

                if (deletedCount == 0)
                {
                    MessageBox.Show("Failed to delete timesheet record(s) from database.", "Error");
                    return;
                }

                System.Diagnostics.Debug.WriteLine($"OnCancelTimesheet: Successfully deleted {deletedCount} of {existingTimesheets.Count} timesheet record(s)");

                try
                {
                    var categoryName = "Timesheet Submitted";

                    System.Diagnostics.Debug.WriteLine($"OnCancelTimesheet: Attempting to remove category from appt.Categories='{appt.Categories}'");

                    // ✅ CRITICAL FIX: Explicitly clear the category by setting to a different value first
                    // This forces Outlook to refresh the category cache
                    if (!string.IsNullOrEmpty(appt.Categories))
                    {
                        var categories = appt.Categories.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(c => c.Trim())
                            .Where(c => !c.Equals(categoryName, StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        // ✅ WORKAROUND: Set to temp category first to force Outlook to clear cache
                        if (categories.Count == 0)
                        {
                            appt.Categories = "Temporary-Clear-Category";
                            appt.Save();
                            System.Diagnostics.Debug.WriteLine("OnCancelTimesheet: Applied temporary category to force cache clear");

                            // Now set to empty
                            appt.Categories = string.Empty;
                            appt.Save();
                            System.Diagnostics.Debug.WriteLine("OnCancelTimesheet: Cleared all categories");
                        }
                        else
                        {
                            appt.Categories = string.Join(", ", categories);
                            appt.Save();
                            System.Diagnostics.Debug.WriteLine($"OnCancelTimesheet: Updated categories to '{appt.Categories}'");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("OnCancelTimesheet: No categories found on appointment");
                    }

                    // ✅ CRITICAL FIX: Remove metadata to prevent ItemChange from re-inserting
                    var ups = appt.UserProperties;
                    var programCodeProp = ups.Find("ProgramCode");
                    var activityCodeProp = ups.Find("ActivityCode");
                    var stageCodeProp = ups.Find("StageCode");

                    if (programCodeProp != null) programCodeProp.Delete();
                    if (activityCodeProp != null) activityCodeProp.Delete();
                    if (stageCodeProp != null) stageCodeProp.Delete();

                    System.Diagnostics.Debug.WriteLine("OnCancelTimesheet: Removed metadata properties to prevent re-insertion");

                    appt.Save();
                    System.Diagnostics.Debug.WriteLine("OnCancelTimesheet: Final save completed");
                }
                catch (Exception catEx)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to remove category/metadata: {catEx.Message}\nStack: {catEx.StackTrace}");
                    // Continue even if category removal fails
                }

                MessageBox.Show(
                    $"Timesheet submission cancelled successfully!\n\n" +
                    $"Deleted {deletedCount} timesheet record(s) from the database and cleared the category.",
                    "Timesheet Cancelled");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cancel Timesheet failed: " + ex.Message + "\n\nStack: " + ex.StackTrace, "Error");
            }
            finally
            {
                // Release COM objects
                if (appt != null)
                {
                    Marshal.ReleaseComObject(appt);
                    appt = null;
                }
                if (insp != null)
                {
                    Marshal.ReleaseComObject(insp);
                    insp = null;
                }
                if (exp != null)
                {
                    Marshal.ReleaseComObject(exp);
                    exp = null;
                }
            }
        }

        // === Helpers ===

        private static string GetUP(Outlook.AppointmentItem appt, string name)
        {
            try
            {
                var ups = appt.UserProperties;
                var up = (ups != null) ? ups.Find(name) : null;
                return (up != null && up.Value != null) ? up.Value.ToString() : string.Empty;
            }
            catch { return string.Empty; }
        }

        private static void AddOrSetTextProp(Outlook.UserProperties ups, string name, string value)
        {
            var up = ups.Find(name) ?? ups.Add(name, Outlook.OlUserPropertyType.olText, false, Type.Missing);
            up.Value = value ?? string.Empty;
        }

        // Helper method to safely get GlobalAppointmentID
        private static string GetGlobalId(Outlook.AppointmentItem appt)
        {
            try { return appt?.GlobalAppointmentID ?? string.Empty; }
            catch { return string.Empty; }
        }

        // Helper method to get all recipients as semicolon-separated string
        private static string GetAllRecipients(Outlook.AppointmentItem appt)
        {
            if (appt == null) return string.Empty;

            // Check if this is a meeting (has recipients) or just a calendar appointment
            // For calendar appointments (no recipients), return empty string
            if (appt.MeetingStatus == Outlook.OlMeetingStatus.olNonMeeting)
            {
                return string.Empty;  // Calendar appointment, no recipients
            }

            var recipientEmails = new System.Collections.Generic.List<string>();
            Outlook.Recipients recipients = null;

            try
            {
                // Get all recipients from the meeting
                recipients = appt.Recipients;
                if (recipients != null && recipients.Count > 0)
                {
                    foreach (Outlook.Recipient recipient in recipients)
                    {
                        try
                        {
                            var email = GetRecipientEmail(recipient);
                            if (!string.IsNullOrWhiteSpace(email))
                            {
                                recipientEmails.Add(email);
                            }
                        }
                        catch { }
                        finally
                        {
                            // Release each recipient
                            if (recipient != null)
                            {
                                Marshal.ReleaseComObject(recipient);
                            }
                        }
                    }
                }

                // Add organizer if not already in list
                var organizerEmail = GetRecipientEmailFromAppointment(appt);
                if (!string.IsNullOrWhiteSpace(organizerEmail) && !recipientEmails.Contains(organizerEmail))
                {
                    recipientEmails.Insert(0, organizerEmail); // Organizer first
                }
            }
            catch { }
            finally
            {
                // Release recipients collection
                if (recipients != null)
                {
                    Marshal.ReleaseComObject(recipients);
                    recipients = null;
                }
            }

            return string.Join("; ", recipientEmails);
        }

        private string GetCurrentUserEmailAddress()
        {
            Outlook.NameSpace session = null;
            Outlook.Recipient currentUser = null;
            Outlook.AddressEntry addrEntry = null;

            try
            {
                session = Globals.ThisAddIn.Application.Session;
                currentUser = session.CurrentUser;
                addrEntry = currentUser?.AddressEntry;

                if (addrEntry != null)
                {
                    if ("EX".Equals(addrEntry.Type, StringComparison.OrdinalIgnoreCase))
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        var smtp = addrEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                        if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                    }
                    if (!string.IsNullOrWhiteSpace(addrEntry.Address)) return addrEntry.Address;
                }
                return currentUser?.Name ?? string.Empty;
            }
            catch { return string.Empty; }
            finally
            {
                if (addrEntry != null)
                {
                    Marshal.ReleaseComObject(addrEntry);
                    addrEntry = null;
                }
                if (currentUser != null)
                {
                    Marshal.ReleaseComObject(currentUser);
                    currentUser = null;
                }
                if (session != null)
                {
                    Marshal.ReleaseComObject(session);
                    session = null;
                }
            }
        }

        private static string GetRecipientEmail(Outlook.Recipient recipient)
        {
            if (recipient == null) return string.Empty;

            Outlook.AddressEntry addrEntry = null;

            try
            {
                addrEntry = recipient.AddressEntry;
                if (addrEntry != null)
                {
                    if ("EX".Equals(addrEntry.Type, StringComparison.OrdinalIgnoreCase))
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        var smtp = addrEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                        if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                    }

                    if (!string.IsNullOrWhiteSpace(addrEntry.Address))
                        return addrEntry.Address;
                }

                return recipient.Address ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                if (addrEntry != null)
                {
                    Marshal.ReleaseComObject(addrEntry);
                    addrEntry = null;
                }
            }
        }

        private static string GetRecipientEmailFromAppointment(Outlook.AppointmentItem appt)
        {
            if (appt == null) return string.Empty;

            Outlook.AddressEntry organizerEntry = null;
            Outlook.NameSpace session = null;
            Outlook.AddressEntry addrEntry = null;
            Outlook.Recipient currentUser = null;
            Outlook.AddressEntry currentUserAddrEntry = null;

            try
            {
                organizerEntry = appt.GetOrganizer();
                if (organizerEntry != null)
                {
                    session = Globals.ThisAddIn?.Application?.Session;
                    if (session != null)
                    {
                        addrEntry = session.GetAddressEntryFromID(organizerEntry.ID);
                        if (addrEntry != null)
                        {
                            if ("EX".Equals(addrEntry.Type, StringComparison.OrdinalIgnoreCase))
                            {
                                const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                                var smtp = addrEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                                if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                            }
                            if (!string.IsNullOrWhiteSpace(addrEntry.Address))
                                return addrEntry.Address;
                        }
                    }
                }
            }
            catch { /* Ignore errors getting organizer */ }
            finally
            {
                if (addrEntry != null)
                {
                    Marshal.ReleaseComObject(addrEntry);
                    addrEntry = null;
                }
                if (organizerEntry != null)
                {
                    Marshal.ReleaseComObject(organizerEntry);
                    organizerEntry = null;
                }
                if (session != null)
                {
                    Marshal.ReleaseComObject(session);
                    session = null;
                }
            }

            try
            {
                session = Globals.ThisAddIn?.Application?.Session;
                currentUser = session?.CurrentUser;
                if (currentUser != null)
                {
                    currentUserAddrEntry = currentUser.AddressEntry;
                    if (currentUserAddrEntry != null)
                    {
                        if ("EX".Equals(currentUserAddrEntry.Type, StringComparison.OrdinalIgnoreCase))
                        {
                            const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                            var smtp = currentUserAddrEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                            if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                        }
                        if (!string.IsNullOrWhiteSpace(currentUserAddrEntry.Address))
                            return currentUserAddrEntry.Address;
                    }
                }
            }
            catch { /* Ignore errors getting current user */ }
            finally
            {
                if (currentUserAddrEntry != null)
                {
                    Marshal.ReleaseComObject(currentUserAddrEntry);
                    currentUserAddrEntry = null;
                }
                if (currentUser != null)
                {
                    Marshal.ReleaseComObject(currentUser);
                    currentUser = null;
                }
                if (session != null)
                {
                    Marshal.ReleaseComObject(session);
                    session = null;
                }
            }

            return string.Empty;
        }

        // === Minimal inline dialog with Multi-Program Support ===
        private class ProgramPickerForm : Form
        {
            private ComboBox cboProgram;
            private ComboBox cboActivity;
            private ComboBox cboStage;
            private CheckBox chkMultiplePrograms;
            private Panel pnlMultiProgram;
            private System.Windows.Forms.FlowLayoutPanel flowPrograms;
            private Button btnAddProgram;
            private Label lblTotalTime;
            private Label lblAllocatedTime;
            private Button btnOk, btnCancel;

            private double _meetingDurationHours;
            private System.Collections.Generic.List<ProgramAllocation> _programAllocations = new System.Collections.Generic.List<ProgramAllocation>();

            // ✅ NEW: Cache for stage and activity data
            private System.Collections.Generic.List<StageCodeData> _stageCodes = new System.Collections.Generic.List<StageCodeData>();
            private System.Collections.Generic.List<ActivityCodeData> _activityCodes = new System.Collections.Generic.List<ActivityCodeData>();
            private System.Collections.Generic.List<string> _programCodes = new System.Collections.Generic.List<string>();

            public bool IsMultiProgram => chkMultiplePrograms?.Checked ?? false;
            public System.Collections.Generic.List<ProgramAllocation> ProgramAllocations => _programAllocations;

            // ✅ FIX: Always return values (needed for original program in multi-program mode)
            public string ProgramCode => cboProgram.SelectedItem?.ToString() ?? string.Empty;
            public string ActivityCode => GetActivityCode(cboActivity.SelectedItem);
            public string StageCode => GetStageCode(cboStage.SelectedItem);

            // ✅ NEW: Helper to extract actual code from display item
            private string GetActivityCode(object selectedItem)
            {
                if (selectedItem is ActivityCodeData actData)
                    return actData.ActivityCode;
                return selectedItem?.ToString() ?? string.Empty;
            }

            // ✅ NEW: Helper to extract actual code from display item
            private string GetStageCode(object selectedItem)
            {
                if (selectedItem is StageCodeData stageData)
                    return stageData.StageCode;
                return selectedItem?.ToString() ?? string.Empty;
            }

            // ✅ NEW: Load activity codes from database
            private async Task LoadActivitiesFromSqlAsync(string initActivity)
            {
                try
                {
                    var cs = ConfigurationManager.ConnectionStrings["OemsDatabase"].ConnectionString;
                    var list = new System.Collections.Generic.List<ActivityCodeData>();

                    using (var cn = new SqlConnection(cs))
                    using (var cmd = new SqlCommand("dbo.Timesheet_GetActivityCodes", cn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        await cn.OpenAsync();
                        using (var rdr = await cmd.ExecuteReaderAsync())
                        {
                            while (await rdr.ReadAsync())
                            {
                                list.Add(new ActivityCodeData
                                {
                                    ActivityCode = rdr["ActivityCode"] as string ?? "",
                                    ActivityDescription = rdr["ActivityDescription"] as string ?? "",
                                    SortOrder = rdr["SortOrder"] != DBNull.Value ? Convert.ToInt32(rdr["SortOrder"]) : 0
                                });
                            }
                        }
                    }

                    _activityCodes = list;

                    cboActivity.BeginUpdate();
                    try
                    {
                        cboActivity.Items.Clear();
                        if (list.Count == 0)
                        {
                            // Fallback to hardcoded values
                            cboActivity.Items.AddRange(new object[] {
                                "Air", "Accommodation", "Food and Beverage", "Side Excursions", "OTHER"
                            });
                        }
                        else
                        {
                            cboActivity.Items.AddRange(list.ToArray());


                            // Select initial value if provided
                            if (!string.IsNullOrWhiteSpace(initActivity))
                            {
                                var match = list.FirstOrDefault(a => a.ActivityCode.Equals(initActivity, StringComparison.OrdinalIgnoreCase));
                                if (match != null)
                                    cboActivity.SelectedItem = match;
                            }
                        }

                        if (cboActivity.SelectedIndex < 0 && cboActivity.Items.Count > 0)
                            cboActivity.SelectedIndex = 0;

                        cboActivity.Enabled = true;
                    }
                    finally
                    {
                        cboActivity.EndUpdate();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"LoadActivitiesFromSqlAsync failed: {ex.Message}");
                    // Fallback to hardcoded values
                    cboActivity.BeginUpdate();
                    try
                    {
                        cboActivity.Items.Clear();
                        cboActivity.Items.AddRange(new object[] {
                            "Air", "Accommodation", "Food and Beverage", "Side Excursions", "OTHER"
                        });
                        if (cboActivity.SelectedIndex < 0) cboActivity.SelectedIndex = 0;
                        cboActivity.Enabled = true;
                    }
                    finally { cboActivity.EndUpdate(); }
                }
            }

            // ✅ NEW: Load stage codes from database
            private async Task LoadStagesFromSqlAsync(string initStage)
            {
                try
                {
                    var cs = ConfigurationManager.ConnectionStrings["OemsDatabase"].ConnectionString;
                    var list = new System.Collections.Generic.List<StageCodeData>();

                    using (var cn = new SqlConnection(cs))
                    using (var cmd = new SqlCommand("dbo.Timesheet_GetStageCodes", cn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        // Pass user email as parameter (stored proc expects it)
                        var userEmail = GetCurrentUserEmail();
                        cmd.Parameters.Add(new SqlParameter("@UserEmail", SqlDbType.NVarChar, 320) { Value = userEmail });

                        await cn.OpenAsync();
                        using (var rdr = await cmd.ExecuteReaderAsync())
                        {
                            while (await rdr.ReadAsync())
                            {
                                list.Add(new StageCodeData
                                {
                                    StageCode = rdr["StageCode"] as string ?? "",
                                    StageDescription = rdr["StageDescription"] as string ?? "",
                                    SortOrder = rdr["SortOrder"] != DBNull.Value ? Convert.ToInt32(rdr["SortOrder"]) : 0
                                });
                            }
                        }
                    }

                    _stageCodes = list;

                    cboStage.BeginUpdate();
                    try
                    {
                        cboStage.Items.Clear();
                        if (list.Count == 0)
                        {
                            // Fallback to hardcoded values
                            cboStage.Items.AddRange(new object[] {
                    "Client Communication", "Internal Communication", "Vendor Communication", "Work Time"
                });
                        }
                        else
                        {
                            cboStage.Items.AddRange(list.ToArray());
                            // Select initial value if provided
                            if (!string.IsNullOrWhiteSpace(initStage))
                            {
                                var match = list.FirstOrDefault(s => s.StageCode.Equals(initStage, StringComparison.OrdinalIgnoreCase));
                                if (match != null)
                                    cboStage.SelectedItem = match;
                            }
                        }
                        if (cboStage.SelectedIndex < 0 && cboStage.Items.Count > 0)
                            cboStage.SelectedIndex = 0;
                        cboStage.Enabled = true;
                    }
                    finally
                    {
                        cboStage.EndUpdate();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"LoadStagesFromSqlAsync failed: {ex.Message}");
                    // Fallback to hardcoded values
                    cboStage.BeginUpdate();
                    try
                    {
                        cboStage.Items.Clear();
                        cboStage.Items.AddRange(new object[] {
                "Client Communication", "Internal Communication", "Vendor Communication", "Work Time"
            });
                        if (cboStage.SelectedIndex < 0) cboStage.SelectedIndex = 0;
                        cboStage.Enabled = true;
                    }
                    finally { cboStage.EndUpdate(); }
                }
            }

            private async Task LoadProgramsFromSqlAsync(string email, string initProgram)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(email)) throw new ArgumentException("Email required");

                    // ✅ Use OemsDatabase (DEV) connection for Test_Program and Test_Proposal tables
                    var cs = ConfigurationManager.ConnectionStrings["OemsDatabase"].ConnectionString;
                    var list = new System.Collections.Generic.List<string>();

                    using (var cn = new SqlConnection(cs))
                    {
                        await cn.OpenAsync();

                        // ✅ Use existing stored procedure (updated for DEV to read from test tables)
                        using (var cmd = new SqlCommand("dbo.TimeSheet_GetActivePrograms", cn))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.CommandTimeout = 30;

                            // Pass user email parameter (stored proc still expects it)
                            cmd.Parameters.Add(new SqlParameter("@UserEmail", SqlDbType.NVarChar, 320) { Value = email });

                            using (var rdr = await cmd.ExecuteReaderAsync())
                            {
                                while (await rdr.ReadAsync())
                                {
                                    var code = rdr["ProgramCode"] as string;
                                    if (!string.IsNullOrWhiteSpace(code))
                                    {
                                        list.Add(code.Trim());
                                    }
                                }
                            }
                        }

                        System.Diagnostics.Debug.WriteLine($"Loaded {list.Count} program/proposal codes from Test_Program and Test_Proposal tables (DEV)");
                    }

                    // Stored proc already does DISTINCT and ORDER BY
                    _programCodes = list;

                    cboProgram.BeginUpdate();
                    try
                    {
                        cboProgram.Items.Clear();

                        // ✅ NEW: Add fixed job codes FIRST (always at the top)
                        cboProgram.Items.Add("Project-OEMS A12004");
                        cboProgram.Items.Add("Project Monday A12005");
                        cboProgram.Items.Add("Buying/Proposal A13000");
                        cboProgram.Items.Add("People-Vacation A14001");
                        cboProgram.Items.Add("People-Personal Time A14002");
                        cboProgram.Items.Add("People-Sick A14003");
                        cboProgram.Items.Add("People-Stat Holiday A14004");
                        cboProgram.Items.Add("Finance-Invoicing/AR A15001");

                        // ✅ Add separator if we have real programs from database
                        if (list.Count > 0)
                        {
                            cboProgram.Items.Add("──────────────────────");
                        }

                        // ✅ Add real programs from database
                        if (list.Count == 0)
                        {
                            cboProgram.Items.Add("(No active programs found)");
                        }
                        else
                        {
                            cboProgram.Items.AddRange(list.ToArray());
                        }

                        // ✅ Select initial value if provided (works for both fixed and database codes)
                        SelectIfPresent(cboProgram, initProgram);
                        if (cboProgram.SelectedIndex < 0) cboProgram.SelectedIndex = 0;

                        cboProgram.Enabled = true;
                    }
                    finally
                    {
                        cboProgram.EndUpdate();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"LoadProgramsFromSqlAsync failed: {ex.Message}");
                    cboProgram.BeginUpdate();
                    try
                    {
                        cboProgram.Items.Clear();

                        // ✅ Even on error, add fixed job codes
                        cboProgram.Items.Add("Project-OEMS A12004");
                        cboProgram.Items.Add("Project Monday A12005");
                        cboProgram.Items.Add("Buying/Proposal A13000");
                        cboProgram.Items.Add("People-Vacation A14001");
                        cboProgram.Items.Add("People-Personal Time A14002");
                        cboProgram.Items.Add("People-Sick A14003");
                        cboProgram.Items.Add("People-Stat Holiday A14004");
                        cboProgram.Items.Add("Finance-Invoicing/AR A15001");
                        cboProgram.Items.Add("──────────────────────");
                        cboProgram.Items.AddRange(new object[] { "test-0001", "test-0002" });

                        if (cboProgram.SelectedIndex < 0) cboProgram.SelectedIndex = 0;
                        cboProgram.Enabled = true;
                    }
                    finally { cboProgram.EndUpdate(); }
                    MessageBox.Show("Failed to load Program/Proposal codes: " + ex.Message);
                }
            }

            public ProgramPickerForm() : this(null, null, null, 0) { }

            public ProgramPickerForm(string initProgram, string initActivity, string initStage) : this(initProgram, initActivity, initStage, 0) { }

            public ProgramPickerForm(string initProgram, string initActivity, string initStage, double meetingDurationHours)
            {
                _meetingDurationHours = meetingDurationHours;

                Text = "Meeting Information";
                FormBorderStyle = FormBorderStyle.FixedDialog;
                StartPosition = FormStartPosition.CenterScreen;
                MaximizeBox = false;
                MinimizeBox = false;
                Width = 520;

                // ✅ FIX: ALWAYS use compact layout - checkbox and buttons on same row
                // The "tall" layout was causing UI inconsistency
                Height = 220;

                // ✅ CRITICAL: Set to always appear in front of Outlook window
                TopMost = true;

                // Single program controls - EACH ON SEPARATE LINES
                var lblProgram = new Label { Left = 15, Top = 20, Width = 120, Text = "Program Code:" };
                cboProgram = new ComboBox { Left = 140, Top = 16, Width = 340, DropDownStyle = ComboBoxStyle.DropDownList, TabIndex = 1 };

                var lblActivity = new Label { Left = 15, Top = 55, Width = 120, Text = "Activity Code:" };
                cboActivity = new ComboBox { Left = 140, Top = 51, Width = 340, DropDownStyle = ComboBoxStyle.DropDownList, TabIndex = 2 };

                var lblStage = new Label { Left = 15, Top = 90, Width = 120, Text = "Stage Code:" };
                cboStage = new ComboBox { Left = 140, Top = 86, Width = 340, DropDownStyle = ComboBoxStyle.DropDownList, TabIndex = 3 };

                // ✅ FIX: Always create checkbox, positioned for compact mode
                chkMultiplePrograms = new CheckBox
                {
                    Left = 15,
                    Top = 125,
                    Width = 290,  // ✅ Increased width for new text
                    Text = "Add Additional Programs",  // ✅ UPDATED: Clearer label
                    TabIndex = 4,
                    Visible = meetingDurationHours > 0  // ✅ Only show for existing meetings
                };
                chkMultiplePrograms.CheckedChanged += ChkMultiplePrograms_CheckedChanged;

                // ✅ Multi-program panel for ADDITIONAL programs
                pnlMultiProgram = new Panel
                {
                    Left = 15,
                    Top = 155,
                    Width = 475,
                    Height = 270,  // ✅ Taller for heading
                    BorderStyle = BorderStyle.FixedSingle,
                    Visible = false,
                    BackColor = System.Drawing.Color.FromArgb(250, 250, 250)
                };

                flowPrograms = new FlowLayoutPanel
                {
                    Left = 10,
                    Top = 30,  // ✅ Moved down for heading
                    Width = 450,
                    Height = 180,
                    AutoScroll = true,
                    FlowDirection = FlowDirection.TopDown,
                    WrapContents = false,
                    BorderStyle = BorderStyle.None
                };

                btnAddProgram = new Button
                {
                    Left = 10,
                    Top = 215,  // ✅ Adjusted for heading
                    Width = 150,
                    Height = 25,
                    Text = "+ Add Program",
                    BackColor = Color.FromArgb(0, 120, 212),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };

                btnAddProgram.FlatAppearance.BorderSize = 0;
                btnAddProgram.Click += BtnAddProgram_Click;

                lblAllocatedTime = new Label
                {
                    Left = 170,
                    Top = 220,  // ✅ Adjusted for heading
                    Width = 295,  // ✅ INCREASED from 280 to 295 to prevent text cutoff
                    Height = 25,  // ✅ Increased from 20 to 25 to prevent descender cutoff
                    Text = $"Allocated: {_meetingDurationHours:F1} / {_meetingDurationHours:F1} hrs",  // ✅ Changed to F1 for consistency
                    Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                    ForeColor = System.Drawing.Color.Green
                };

                pnlMultiProgram.Controls.AddRange(new Control[] {
                    flowPrograms, btnAddProgram, lblAllocatedTime
                });

                // ✅ FIX: ALWAYS use compact layout - buttons on same row as checkbox
                btnOk = new Button { Left = 310, Top = 123, Width = 80, Height = 28, Text = "OK", DialogResult = DialogResult.OK, TabIndex = 5 };
                btnCancel = new Button { Left = 400, Top = 123, Width = 80, Height = 28, Text = "Cancel", DialogResult = DialogResult.Cancel, TabIndex = 6 };

                btnOk.Click += BtnOk_Click;

                Controls.AddRange(new Control[] {
                    lblProgram, cboProgram,
                    lblActivity, cboActivity,
                    lblStage, cboStage,
                    chkMultiplePrograms,
                    pnlMultiProgram,
                    btnOk, btnCancel
                });
                AcceptButton = btnOk; CancelButton = btnCancel;

                // ✅ NEW: Set to "Loading..." state for all dropdowns
                cboProgram.Items.Add("Loading...");
                cboProgram.SelectedIndex = 0;
                cboProgram.Enabled = false;

                cboActivity.Items.Add("Loading...");
                cboActivity.SelectedIndex = 0;
                cboActivity.Enabled = false;

                cboStage.Items.Add("Loading...");
                cboStage.SelectedIndex = 0;
                cboStage.Enabled = false;

                var programToSelect = initProgram;
                var activityToSelect = initActivity;
                var stageToSelect = initStage;

                this.Shown += async (s, e) =>
                {
                    var user = GetCurrentUserEmail() ?? string.Empty;

                    // ✅ Load all data in parallel for faster performance
                    await Task.WhenAll(
                        LoadProgramsFromSqlAsync(user, programToSelect),
                        LoadActivitiesFromSqlAsync(activityToSelect),
                        LoadStagesFromSqlAsync(stageToSelect)
                    );
                };
            }

            private void ChkMultiplePrograms_CheckedChanged(object sender, EventArgs e)
            {
                bool isMulti = chkMultiplePrograms.Checked;

                if (isMulti)
                {
                    // ✅ NEW: EXPAND form to show ADDITIONAL programs panel (original program stays at top)
                    this.Height = 530;  // ✅ Match ManageTimesheetPane
                    btnOk.Top = 435;     // ✅ Match ManageTimesheetPane
                    btnCancel.Top = 435; // ✅ Match ManageTimesheetPane
                    btnOk.Left = 320;
                    btnCancel.Left = 410;

                    // ✅ CRITICAL CHANGE: KEEP original program controls ENABLED (user can edit them)
                    cboProgram.Enabled = true;
                    cboActivity.Enabled = true;
                    cboStage.Enabled = true;

                    pnlMultiProgram.Visible = true;

                    // ✅ CRITICAL CHANGE: Don't auto-populate - user will add additional programs manually
                    // if (_programAllocations.Count == 0)
                    // {
                    //     AddProgramAllocationControl(
                    //         cboProgram.SelectedItem?.ToString(),
                    //         cboActivity.SelectedItem?.ToString(),
                    //         cboStage.SelectedItem?.ToString()
                    //     );
                    // }
                }
                else
                {
                    // ✅ SHRINK back to compact layout (always use compact)
                    this.Height = 220;
                    btnOk.Top = 123;
                    btnCancel.Top = 123;
                    btnOk.Left = 310;
                    btnCancel.Left = 400;

                    // Re-enable single-program controls
                    cboProgram.Enabled = true;
                    cboActivity.Enabled = true;
                    cboStage.Enabled = true;

                    // Hide multi-program panel
                    pnlMultiProgram.Visible = false;

                    // ✅ CLEAR all ADDITIONAL program allocations
                    foreach (Control ctrl in flowPrograms.Controls.OfType<ProgramAllocationControl>().ToList())
                    {
                        flowPrograms.Controls.Remove(ctrl);
                        ctrl.Dispose();
                    }
                    _programAllocations.Clear();
                }
            }

            private void BtnAddProgram_Click(object sender, EventArgs e)
            {
                AddProgramAllocationControl();
            }

            private void AddProgramAllocationControl(string initProgram = null, string initActivity = null, string initStage = null)
            {
                var allocationControl = new ProgramAllocationControl(
                    _meetingDurationHours,
                    cboProgram.Items.Cast<string>().ToList(),
                    initProgram,
                    initActivity,
                    initStage
                );
                allocationControl.OnRemove += (s, ev) =>
                {
                    var ctrl = s as ProgramAllocationControl;
                    _programAllocations.Remove(ctrl.Allocation);
                    flowPrograms.Controls.Remove(ctrl);
                    ctrl.Dispose();
                    UpdateTotalAllocated();
                };
                allocationControl.OnHoursChanged += (s, ev) => UpdateTotalAllocated();

                flowPrograms.Controls.Add(allocationControl);
                _programAllocations.Add(allocationControl.Allocation);
                UpdateTotalAllocated();
            }

            private void UpdateTotalAllocated()
            {
                // ✅ Calculate ADDITIONAL programs only (original is separate)
                double additionalHours = _programAllocations.Sum(p => p.Hours);
                double originalProgramHours = _meetingDurationHours - additionalHours;

                // ✅ Show original program code with its calculated hours
                string originalProgram = cboProgram.SelectedItem?.ToString() ?? "Original";
                lblAllocatedTime.Text = $"Total: {_meetingDurationHours:F1}h ({originalProgram}: {originalProgramHours:F1}h)";  // ✅ Changed to F1 for consistency

                // ✅ Validation: original program hours must be > 0
                bool isValid = originalProgramHours > 0.01;
                lblAllocatedTime.ForeColor = isValid ? System.Drawing.Color.Green : System.Drawing.Color.Red;
            }

            private void BtnOk_Click(object sender, EventArgs e)
            {
                // ✅ ALWAYS validate original program is selected
                if (string.IsNullOrWhiteSpace(cboProgram.SelectedItem?.ToString()))
                {
                    MessageBox.Show("Please select a program code.",
                        "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }

                if (chkMultiplePrograms != null && chkMultiplePrograms.Checked)
                {
                    // ✅ NEW VALIDATION: Require at least 1 ADDITIONAL program (original + at least 1 more)
                    if (_programAllocations.Count < 1)
                    {
                        MessageBox.Show(
                            "Please add at least 1 additional program.\n\n" +
                            "If this meeting only involves one program, please uncheck the 'Add Additional Programs' checkbox.",
                            "Additional Programs Required",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        DialogResult = DialogResult.None;
                        return;
                    }

                    // ✅ NEW VALIDATION: Check for duplicate program codes
                    string originalProgram = cboProgram.SelectedItem?.ToString() ?? "";

                    // ✅ CRITICAL FIX: Check if ANY additional program has the same code as original
                    var duplicatePrograms = _programAllocations
                        .Where(p => !string.IsNullOrWhiteSpace(p.ProgramCode) &&
                                   p.ProgramCode.Equals(originalProgram, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (duplicatePrograms.Count > 0)
                    {
                        MessageBox.Show(
                            $"Cannot add '{originalProgram}' as an additional program because it's already the original program.\n\n" +
                            $"Please select a different program for the additional allocation, or adjust the original program instead.",
                            "Duplicate Program Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        DialogResult = DialogResult.None;
                        return;
                    }

                    // ✅ NEW: Calculate total INCLUDING original program
                    // Additional programs must leave room for the original program
                    double additionalHours = _programAllocations.Sum(p => p.Hours);

                    // ✅ Validate that additional programs don't exceed meeting duration
                    if (additionalHours >= _meetingDurationHours)
                    {
                        MessageBox.Show(
                            $"Additional programs ({additionalHours:F2} hrs) cannot equal or exceed total meeting duration ({_meetingDurationHours:F2} hrs).\n\n" +
                            $"Please leave time for the original program.",
                            "Invalid Allocation",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        DialogResult = DialogResult.None;
                        return;
                    }

                    if (_programAllocations.Any(p => string.IsNullOrWhiteSpace(p.ProgramCode)))
                    {
                        MessageBox.Show("Please select a program code for all additional entries.",
                            "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DialogResult = DialogResult.None;
                        return;
                    }

                    System.Diagnostics.Debug.WriteLine($"[BtnOk_Click] Multi-program validation passed: 1 original + {_programAllocations.Count} additional programs");
                }
                else
                {
                    // ✅ Single-program mode validation remains the same
                    System.Diagnostics.Debug.WriteLine($"[BtnOk_Click] Single-program validation passed: {cboProgram.SelectedItem}");
                }
            }

            private static void SelectIfPresent(ComboBox combo, string value)
            {
                if (string.IsNullOrWhiteSpace(value)) return;
                var idx = combo.FindStringExact(value);
                if (idx >= 0) combo.SelectedIndex = idx;
            }

            private static string GetCurrentUserEmail()
            {
                Outlook.NameSpace session = null;
                Outlook.Recipient currentUser = null;
                Outlook.AddressEntry addrEntry = null;

                try
                {
                    session = Globals.ThisAddIn.Application.Session;
                    currentUser = session.CurrentUser;
                    addrEntry = currentUser?.AddressEntry;

                    if (addrEntry != null)
                    {
                        if ("EX".Equals(addrEntry.Type, StringComparison.OrdinalIgnoreCase))
                        {
                            const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                            var smtp = addrEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                            if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                        }
                        if (!string.IsNullOrWhiteSpace(addrEntry.Address)) return addrEntry.Address;
                    }
                    return currentUser?.Name ?? string.Empty;
                }
                catch { return string.Empty; }
                finally
                {
                    if (addrEntry != null)
                    {
                        Marshal.ReleaseComObject(addrEntry);
                        addrEntry = null;
                    }
                    if (currentUser != null)
                    {
                        Marshal.ReleaseComObject(currentUser);
                        currentUser = null;
                    }
                    if (session != null)
                    {
                        Marshal.ReleaseComObject(session);
                        session = null;
                    }
                }
            }
        }

        // Helper classes for program allocation
        private class ProgramAllocation
        {
            public string ProgramCode { get; set; }
            public string ActivityCode { get; set; }
            public string StageCode { get; set; }
            public double Hours { get; set; }
        }

        // ✅ NEW: Helper class for stage code data
        private class StageCodeData
        {
            public string StageCode { get; set; }
            public string StageDescription { get; set; }
            public int SortOrder { get; set; }

            public override string ToString() => StageDescription; // Display description in ComboBox
        }

        // ✅ NEW: Helper class for activity code data
        private class ActivityCodeData
        {
            public string ActivityCode { get; set; }
            public string ActivityDescription { get; set; }
            public int SortOrder { get; set; }

            public override string ToString() => ActivityDescription; // Display description in ComboBox
        }

        private class ProgramAllocationControl : Panel
        {
            public ProgramAllocation Allocation { get; private set; }
            public event EventHandler OnRemove;
            public event EventHandler OnHoursChanged;

            private ComboBox cboProgram;
            private ComboBox cboActivity;
            private ComboBox cboStage;
            private TrackBar trackHours;
            private Label lblHours;
            private Button btnRemove;
            private double _maxHours;

            public ProgramAllocationControl(double maxHours, System.Collections.Generic.List<string> programs, string initProgram = null, string initActivity = null, string initStage = null)
            {
                _maxHours = maxHours;
                Allocation = new ProgramAllocation { Hours = maxHours };

                Width = 450;
                Height = 135;
                BorderStyle = BorderStyle.FixedSingle;
                Margin = new Padding(0, 0, 0, 8);
                BackColor = System.Drawing.Color.FromArgb(250, 250, 250);

                cboProgram = new ComboBox { Left = 10, Top = 10, Width = 350, DropDownStyle = ComboBoxStyle.DropDownList };
                cboProgram.Items.AddRange(programs.ToArray());
                if (!string.IsNullOrWhiteSpace(initProgram))
                {
                    var idx = cboProgram.FindStringExact(initProgram);
                    if (idx >= 0) cboProgram.SelectedIndex = idx;
                }
                if (cboProgram.SelectedIndex < 0 && cboProgram.Items.Count > 0) cboProgram.SelectedIndex = 0;
                cboProgram.SelectedIndexChanged += (s, e) => Allocation.ProgramCode = cboProgram.SelectedItem?.ToString() ?? "";

                cboActivity = new ComboBox { Left = 10, Top = 40, Width = 350, DropDownStyle = ComboBoxStyle.DropDownList };
                // ✅ Use database-friendly descriptions
                cboActivity.Items.AddRange(new object[] { "Work Time", "Client Communication", "Vendor Communication", "Internal Communication" });
                if (!string.IsNullOrWhiteSpace(initActivity))
                {
                    var idx = cboActivity.FindStringExact(initActivity);
                    if (idx >= 0) cboActivity.SelectedIndex = idx;
                }
                if (cboActivity.SelectedIndex < 0 && cboActivity.Items.Count > 0) cboActivity.SelectedIndex = 0;
                cboActivity.SelectedIndexChanged += (s, e) => Allocation.ActivityCode = cboActivity.SelectedItem?.ToString() ?? "";

                cboStage = new ComboBox { Left = 10, Top = 70, Width = 350, DropDownStyle = ComboBoxStyle.DropDownList };
                // ✅ Use database-friendly descriptions
                cboStage.Items.AddRange(new object[] { "Client Meeting", "Internal Meeting", "Email", "Vendor Research", "Meeting", "Timesheet", "Design", "Registration" });
                if (!string.IsNullOrWhiteSpace(initStage))
                {
                    var idx = cboStage.FindStringExact(initStage);
                    if (idx >= 0) cboStage.SelectedIndex = idx;
                }
                if (cboStage.SelectedIndex < 0 && cboStage.Items.Count > 0) cboStage.SelectedIndex = 0;
                cboStage.SelectedIndexChanged += (s, e) => Allocation.StageCode = cboStage.SelectedItem?.ToString() ?? "";

                // ✅ UPDATED: Trackbar in minutes (0 to maxHours*60), with tick every 15 minutes
                int maxMinutes = (int)(maxHours * 60);
                trackHours = new TrackBar
                {
                    Left = 10,
                    Top = 100,
                    Width = 300,
                    Minimum = 0,
                    Maximum = maxMinutes,
                    Value = maxMinutes,
                    TickFrequency = 15,  // Tick every 15 minutes
                    SmallChange = 5,     // Arrow keys move by 5 minutes
                    LargeChange = 15     // Page up/down moves by 15 minutes
                };
                trackHours.ValueChanged += TrackHours_ValueChanged;

                // ✅ UPDATED: Show both hours and minutes in label
                int initialMinutes = (int)(Allocation.Hours * 60);
                lblHours = new Label
                {
                    Left = 320,
                    Top = 105,
                    Width = 70,
                    Text = FormatTimeLabel(initialMinutes),
                    Font = new System.Drawing.Font("Segoe UI", 8, System.Drawing.FontStyle.Bold)
                };

                btnRemove = new Button { Left = 370, Top = 10, Width = 70, Height = 25, Text = "Delete", BackColor = System.Drawing.Color.Red, ForeColor = System.Drawing.Color.White, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 8, System.Drawing.FontStyle.Bold) };
                btnRemove.FlatAppearance.BorderSize = 0;
                btnRemove.Click += (s, e) => OnRemove?.Invoke(this, EventArgs.Empty);

                Controls.AddRange(new Control[] { cboProgram, cboActivity, cboStage, trackHours, lblHours, btnRemove });

                Allocation.ProgramCode = cboProgram.SelectedItem?.ToString() ?? "";
                Allocation.ActivityCode = cboActivity.SelectedItem?.ToString() ?? "";
                Allocation.StageCode = cboStage.SelectedItem?.ToString() ?? "";
            }

            // ✅ NEW: Helper to format time as "Xh Ymins" or "Ymins" if less than 1 hour
            private string FormatTimeLabel(int totalMinutes)
            {
                if (totalMinutes < 60)
                {
                    return $"{totalMinutes}mins";
                }
                else
                {
                    int hours = totalMinutes / 60;
                    int minutes = totalMinutes % 60;
                    if (minutes == 0)
                        return $"{hours}h";
                    else
                        return $"{hours}h {minutes}mins";
                }
            }

            private void TrackHours_ValueChanged(object sender, EventArgs e)
            {
                // ✅ UPDATED: Convert minutes to hours (decimal)
                int minutes = trackHours.Value;
                Allocation.Hours = minutes / 60.0;
                lblHours.Text = FormatTimeLabel(minutes);
                OnHoursChanged?.Invoke(this, EventArgs.Empty);
            }
        }
    }
}