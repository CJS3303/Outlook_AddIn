# Submitted Tab Implementation

## Overview
The Submitted tab displays submitted timesheet records from the database with options to:
- Cancel submissions (delete records)
- Un-ignore meetings (remove ignored status)

## Key Implementation Notes

1. **LoadSubmittedMeetingsAsync** loads data directly from database (unlike unsubmitted which loads from Outlook)
2. **Reuses existing DbWriter methods:**
   - `DbWriter.DeleteTimesheetAsync()` - for canceling submissions
   - `DbWriter.CancelIgnoreTimesheetAsync()` - for un-ignoring
3. **Visual distinction:**
   - Green headers for submitted items (vs blue for unsubmitted)
   - Light green background for submitted cards (vs white)

## Required Changes Summary

The LoadSubmittedMeetingsAsync method should:
1. Query database for records with status='submitted'
2. Filter for last 30 days
3. Group by Today/This Week/Older
4. Create panels showing Program, Activity, Stage
5. Include "Cancel Submit" and "Un-Ignore" buttons
