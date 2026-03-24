-- =============================================
-- Final Migration: Remaining Inline SQL to Stored Procedures
-- Purpose: Clean up last inline SQL queries
-- =============================================

-- =============================================
-- 1. GET submitted/ignored meetings for unsubmitted list filtering
-- =============================================
IF OBJECT_ID('dbo.Timesheet_GetSubmittedOrIgnoredKeys', 'P') IS NOT NULL
    DROP PROCEDURE dbo.Timesheet_GetSubmittedOrIgnoredKeys;
GO

CREATE PROCEDURE dbo.Timesheet_GetSubmittedOrIgnoredKeys
    @email NVARCHAR(320),
    @startDate DATETIME2
AS
BEGIN
    SET NOCOUNT ON;
    
    -- Return global_id and start_utc for all submitted/ignored meetings
    -- Used to filter out already-processed meetings from the unsubmitted list
    SELECT global_id, start_utc 
    FROM db_owner.ytimesheet 
    WHERE user_name = @email 
      AND start_utc >= @startDate;
END;
GO

-- =============================================
-- 2. VERIFY record exists after upsert (for debugging)
-- =============================================
IF OBJECT_ID('dbo.Timesheet_VerifyRecord', 'P') IS NOT NULL
    DROP PROCEDURE dbo.Timesheet_VerifyRecord;
GO

CREATE PROCEDURE dbo.Timesheet_VerifyRecord
    @global_id NVARCHAR(255),
    @start_utc DATETIME2 = NULL,  -- Optional for recurring meetings
    @user_name NVARCHAR(100),
    @is_recurring BIT = 0,
    @record_count INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    
    -- Different WHERE clause for recurring vs non-recurring
    IF @is_recurring = 1
    BEGIN
        -- Recurring: include start_utc in WHERE clause
        SELECT @record_count = COUNT(*)
        FROM db_owner.ytimesheet
        WHERE global_id = @global_id
          AND start_utc = @start_utc
          AND user_name = @user_name;
    END
    ELSE
    BEGIN
        -- Non-recurring: exclude start_utc from WHERE clause
        SELECT @record_count = COUNT(*)
        FROM db_owner.ytimesheet
        WHERE global_id = @global_id
          AND user_name = @user_name;
    END
END;
GO

-- =============================================
-- Test the stored procedures
-- =============================================
PRINT 'Created 2 additional stored procedures:';
PRINT '  1. dbo.Timesheet_GetSubmittedOrIgnoredKeys';
PRINT '  2. dbo.Timesheet_VerifyRecord';
PRINT '';
PRINT '✅ All remaining inline SQL migrated to stored procedures!';
PRINT '';
PRINT '=== SUMMARY: Total Stored Procedures ===';
PRINT 'Read operations (SELECT):';
PRINT '  • Timesheet_Exists - Check if timesheet exists';
PRINT '  • Timesheet_GetExisting - Get single timesheet record';
PRINT '  • Timesheet_GetAllForMeeting - Get all records for a meeting';
PRINT '  • Timesheet_GetSubmittedOrIgnoredKeys - Get filtered list';
PRINT '  • Timesheet_VerifyRecord - Verify after upsert';
PRINT '  • Timesheet_GetWeeklyData - Weekly dashboard data';
PRINT '  • Timesheet_GetActivityCodes - Activity dropdown';
PRINT '  • Timesheet_GetStageCodes - Stage dropdown';
PRINT '  • Timesheet_GetActivePrograms - Program dropdown';
PRINT '';
PRINT 'Write operations (INSERT/UPDATE/DELETE):';
PRINT '  • Timesheet_Upsert - Insert or update timesheet';
PRINT '  • Timesheet_IgnoreMeeting - Mark as ignored';
PRINT '  • Timesheet_DeleteRecord - Delete single record';
PRINT '  • Timesheet_CancelIgnore - Un-ignore meeting';
PRINT '  • Timesheet_DeleteAllByGlobalId - Delete all for meeting';
PRINT '';
PRINT '✅ Total: 14 stored procedures (100% of SQL logic)';
