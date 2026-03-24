using System;
using System.Data;
using System.Configuration;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;

namespace OutlookAddIn1
{
    public static class DbWriter
    {
        private static readonly string ConnString =
            System.Configuration.ConfigurationManager
                   .ConnectionStrings["OemsDatabase"]?.ConnectionString;

        private static object DbOrNull(string s)
        {
            return string.IsNullOrWhiteSpace(s) ? (object)DBNull.Value : s;
        }

        private static object DbOrNull<T>(T? v) where T : struct
        {
            return v.HasValue ? (object)v.Value : DBNull.Value;
        }

        /// <summary>
        /// Checks if a timesheet already exists for the given meeting and user
        /// </summary>
        public static async Task<bool> TimesheetExistsAsync(MeetingRecord r)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ConnString))
                {
                    throw new InvalidOperationException("ConnectionStrings['OemsDatabase'] is null/empty. Check App.config key.");
                }

                var globalId = !string.IsNullOrWhiteSpace(r.GlobalId) ? r.GlobalId : r.EntryId;
                if (string.IsNullOrWhiteSpace(globalId))
                {
                    throw new ArgumentException("global_id is required (use GlobalAppointmentID or EntryID).",
                                                nameof(r.GlobalId));
                }

                using (var cn = new SqlConnection(ConnString))
                {
                    await cn.OpenAsync().ConfigureAwait(false);

                    // ✅ NEW: Use stored procedure instead of inline SQL
                    using (var cmd = new SqlCommand("dbo.Timesheet_Exists", cn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = globalId });
                        cmd.Parameters.Add(new SqlParameter("@start_utc", SqlDbType.DateTime2) { Value = r.StartTorontoTime });
                        cmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar, 100) { Value = DbOrNull(r.UserDisplayName) });

                        // Output parameter
                        var existsParam = new SqlParameter("@exists", SqlDbType.Bit) { Direction = ParameterDirection.Output };
                        cmd.Parameters.Add(existsParam);

                        await cmd.ExecuteNonQueryAsync().ConfigureAwait(false);

                        return (bool)existsParam.Value;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("TimesheetExistsAsync failed: " + ex.Message);
                // If we can't check, assume it doesn't exist to allow the operation
                return false;
            }
        }

        /// <summary>
        /// Gets the existing timesheet record for the given meeting and user
        /// Returns null if not found
        /// </summary>
        public static async Task<MeetingRecord> GetExistingTimesheetAsync(MeetingRecord r)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ConnString))
                {
                    throw new InvalidOperationException("ConnectionStrings['OemsDatabase'] is null/empty. Check App.config key.");
                }

                var globalId = !string.IsNullOrWhiteSpace(r.GlobalId) ? r.GlobalId : r.EntryId;
                if (string.IsNullOrWhiteSpace(globalId))
                {
                    throw new ArgumentException("global_id is required (use GlobalAppointmentID or EntryID).",
                                                nameof(r.GlobalId));
                }

                using (var cn = new SqlConnection(ConnString))
                {
                    await cn.OpenAsync().ConfigureAwait(false);

                    // ✅ FIXED: Use inline SQL instead of missing stored procedure
                    using (var cmd = new SqlCommand(@"
                        SELECT TOP 1 
                            job_code, 
                            activity_code, 
                            stage_code, 
                            source, 
                            start_utc, 
                            end_utc, 
                            hours_allocated, 
                            last_modified_utc, 
                            status
                        FROM db_owner.ytimesheet
                        WHERE global_id = @global_id 
                          AND start_utc = @start_utc 
                          AND user_name = @user_name
                        ORDER BY last_modified_utc DESC", cn))
                    {
                        cmd.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = globalId });
                        cmd.Parameters.Add(new SqlParameter("@start_utc", SqlDbType.DateTime2) { Value = r.StartTorontoTime });
                        cmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar, 100) { Value = DbOrNull(r.UserDisplayName) });

                        using (var reader = await cmd.ExecuteReaderAsync().ConfigureAwait(false))
                        {
                            if (await reader.ReadAsync().ConfigureAwait(false))
                            {
                                // Database stores Toronto time, read it directly
                                var startTorontoTime = reader["start_utc"] is DateTime st ? st : DateTime.MinValue;
                                var endTorontoTime = reader["end_utc"] is DateTime et ? et : DateTime.MinValue;
                                var lastModifiedTorontoTime = reader["last_modified_utc"] is DateTime lm ? lm : DateTime.MinValue;

                                // Convert Toronto time back to UTC for internal use
                                var torontoTz = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                                var startUtc = TimeZoneInfo.ConvertTimeToUtc(startTorontoTime, torontoTz);
                                var endUtc = TimeZoneInfo.ConvertTimeToUtc(endTorontoTime, torontoTz);
                                var lastModifiedUtc = TimeZoneInfo.ConvertTimeToUtc(lastModifiedTorontoTime, torontoTz);

                                // Read hours_allocated if available
                                double? hoursAllocated = null;
                                if (reader["hours_allocated"] != DBNull.Value)
                                {
                                    hoursAllocated = Convert.ToDouble(reader["hours_allocated"]);
                                }

                                string status = reader["status"] as string ?? "submitted";

                                return new MeetingRecord
                                {
                                    ProgramCode = reader["job_code"] as string ?? "",
                                    ActivityCode = reader["activity_code"] as string ?? "",
                                    StageCode = reader["stage_code"] as string ?? "",
                                    Source = reader["source"] as string ?? "",
                                    StartUtc = startUtc,
                                    EndUtc = endUtc,
                                    HoursAllocated = hoursAllocated,
                                    LastModifiedUtc = lastModifiedUtc,
                                    Status = status,
                                    UserDisplayName = r.UserDisplayName,
                                    GlobalId = r.GlobalId,
                                    EntryId = r.EntryId
                                };
                            }
                        }
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GetExistingTimesheetAsync failed: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Gets ALL existing timesheet records for a given meeting and user (for multi-program submissions)
        /// Returns empty list if none found
        /// </summary>
        public static async Task<System.Collections.Generic.List<MeetingRecord>> GetAllTimesheetsForMeetingAsync(MeetingRecord r)
        {
            var results = new System.Collections.Generic.List<MeetingRecord>();

            try
            {
                if (string.IsNullOrWhiteSpace(ConnString))
                {
                    throw new InvalidOperationException("ConnectionStrings['OemsDatabase'] is null/empty. Check App.config key.");
                }

                var globalId = !string.IsNullOrWhiteSpace(r.GlobalId) ? r.GlobalId : r.EntryId;
                if (string.IsNullOrWhiteSpace(globalId))
                {
                    throw new ArgumentException("global_id is required (use GlobalAppointmentID or EntryID).",
                                                nameof(r.GlobalId));
                }

                using (var cn = new SqlConnection(ConnString))
                {
                    await cn.OpenAsync().ConfigureAwait(false);

                    // ✅ FIXED: Use inline SQL instead of missing stored procedure
                    using (var cmd = new SqlCommand(@"
                        SELECT 
                            job_code, 
                            activity_code, 
                            stage_code, 
                            source, 
                            start_utc, 
                            end_utc, 
                            hours_allocated, 
                            last_modified_utc, 
                            status
                        FROM db_owner.ytimesheet
                        WHERE global_id = @global_id 
                          AND start_utc = @start_utc 
                          AND user_name = @user_name
                        ORDER BY last_modified_utc DESC", cn))
                    {
                        cmd.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = globalId });
                        cmd.Parameters.Add(new SqlParameter("@start_utc", SqlDbType.DateTime2) { Value = r.StartTorontoTime });
                        cmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar, 100) { Value = DbOrNull(r.UserDisplayName) });

                        using (var reader = await cmd.ExecuteReaderAsync().ConfigureAwait(false))
                        {
                            while (await reader.ReadAsync().ConfigureAwait(false))
                            {
                                // ✅ CRITICAL FIX: Database stores Toronto time directly
                                // READ times as-is from database (they are already in Toronto time)
                                var startTorontoTime = reader["start_utc"] is DateTime st ? st : DateTime.MinValue;
                                var endTorontoTime = reader["end_utc"] is DateTime et ? et : DateTime.MinValue;
                                var lastModifiedTorontoTime = reader["last_modified_utc"] is DateTime lm ? lm : DateTime.MinValue;

                                // ✅ NO CONVERSION: Times in database ARE Toronto time
                                // Keep them as StartTorontoTime, EndTorontoTime, LastModifiedTorontoTime
                                // For internal use (StartUtc, EndUtc, LastModifiedUtc), we need to convert FROM Toronto TO UTC

                                var torontoTz = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                                var startUtc = TimeZoneInfo.ConvertTimeToUtc(startTorontoTime, torontoTz);
                                var endUtc = TimeZoneInfo.ConvertTimeToUtc(endTorontoTime, torontoTz);
                                var lastModifiedUtc = TimeZoneInfo.ConvertTimeToUtc(lastModifiedTorontoTime, torontoTz);

                                // Read hours_allocated if available
                                double? hoursAllocated = null;
                                if (reader["hours_allocated"] != DBNull.Value)
                                {
                                    hoursAllocated = Convert.ToDouble(reader["hours_allocated"]);
                                }

                                string status = reader["status"] as string ?? "submitted";

                                results.Add(new MeetingRecord
                                {
                                    ProgramCode = reader["job_code"] as string ?? "",
                                    ActivityCode = reader["activity_code"] as string ?? "",
                                    StageCode = reader["stage_code"] as string ?? "",
                                    Source = reader["source"] as string ?? "",
                                    StartUtc = startUtc,
                                    EndUtc = endUtc,
                                    HoursAllocated = hoursAllocated,
                                    LastModifiedUtc = lastModifiedUtc,
                                    Status = status,
                                    UserDisplayName = r.UserDisplayName,
                                    GlobalId = r.GlobalId,
                                    EntryId = r.EntryId
                                });
                            }
                        }
                    }
                }

                return results;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GetAllTimesheetsForMeetingAsync failed: " + ex.Message);
                return results;
            }
        }

        public static async Task UpsertAsync(MeetingRecord r)
        {
            try
            {
                // Step 1: Check connection string
                if (string.IsNullOrWhiteSpace(ConnString))
                {
                    throw new InvalidOperationException("ConnectionStrings['OemsDatabase'] is null/empty. Check App.config key.");
                }

                // Step 2: Validate input
                var globalId = !string.IsNullOrWhiteSpace(r.GlobalId) ? r.GlobalId : r.EntryId;
                if (string.IsNullOrWhiteSpace(globalId))
                {
                    throw new ArgumentException("global_id is required (use GlobalAppointmentID or EntryID).", nameof(r.GlobalId));
                }

                // ✅ NEW: Log what we're about to upsert
                System.Diagnostics.Debug.WriteLine($"========== UPSERT START ==========");
                System.Diagnostics.Debug.WriteLine($"User: {r.UserDisplayName}");
                System.Diagnostics.Debug.WriteLine($"Program: {r.ProgramCode}");
                System.Diagnostics.Debug.WriteLine($"Activity: {r.ActivityCode}");
                System.Diagnostics.Debug.WriteLine($"Stage: {r.StageCode}");
                System.Diagnostics.Debug.WriteLine($"GlobalId: {globalId}");
                System.Diagnostics.Debug.WriteLine($"StartTorontoTime: {r.StartTorontoTime:yyyy-MM-dd HH:mm:ss}");
                System.Diagnostics.Debug.WriteLine($"EndTorontoTime: {r.EndTorontoTime:yyyy-MM-dd HH:mm:ss}");
                System.Diagnostics.Debug.WriteLine($"HoursAllocated: {r.HoursAllocated}");
                System.Diagnostics.Debug.WriteLine($"Status: {r.Status ?? "submitted"}");

                using (var cn = new SqlConnection(ConnString))
                {
                    await cn.OpenAsync().ConfigureAwait(false);

                    // ✅ NEW: Log connection info
                    using (var who = new SqlCommand("SELECT DB_NAME(), SUSER_SNAME()", cn))
                    using (var rdr = await who.ExecuteReaderAsync().ConfigureAwait(false))
                    {
                        if (await rdr.ReadAsync().ConfigureAwait(false))
                        {
                            var dbName = rdr.GetString(0);
                            var loginName = rdr.GetString(1);
                            System.Diagnostics.Debug.WriteLine($"✅ Connected to DB: {dbName}, Login: {loginName}");
                        }
                    }

                    // Execute stored procedure
                    using (var cmd = new SqlCommand("dbo.Timesheet_Upsert", cn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        // Core identity/time - SEND TORONTO TIME TO DATABASE
                        cmd.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = globalId });
                        cmd.Parameters.Add(new SqlParameter("@start_utc", SqlDbType.DateTime2) { Value = r.StartTorontoTime });
                        cmd.Parameters.Add(new SqlParameter("@end_utc", SqlDbType.DateTime2) { Value = r.EndTorontoTime });

                        // Outlook refs
                        cmd.Parameters.Add(new SqlParameter("@entry_id", SqlDbType.NVarChar, 255) { Value = DbOrNull(r.EntryId) });
                        cmd.Parameters.Add(new SqlParameter("@subject", SqlDbType.NVarChar, 500) { Value = DbOrNull(r.Subject) });

                        // Job
                        cmd.Parameters.Add(new SqlParameter("@job_code", SqlDbType.NVarChar, 50) { Value = DbOrNull(r.ProgramCode) });
                        cmd.Parameters.Add(new SqlParameter("@job_name", SqlDbType.NVarChar, 255) { Value = DBNull.Value });
                        cmd.Parameters.Add(new SqlParameter("@proposalId", SqlDbType.BigInt) { Value = DBNull.Value });

                        // User
                        cmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar, 100) { Value = DbOrNull(r.UserDisplayName) });
                        cmd.Parameters.Add(new SqlParameter("@userId", SqlDbType.BigInt) { Value = DBNull.Value });

                        // Activity/Stage
                        cmd.Parameters.Add(new SqlParameter("@activity_code", SqlDbType.NVarChar, 50) { Value = DbOrNull(r.ActivityCode) });
                        cmd.Parameters.Add(new SqlParameter("@activity_description", SqlDbType.NVarChar, 255) { Value = DBNull.Value });
                        cmd.Parameters.Add(new SqlParameter("@stage_code", SqlDbType.NVarChar, 50) { Value = DbOrNull(r.StageCode) });
                        cmd.Parameters.Add(new SqlParameter("@stage_description", SqlDbType.NVarChar, 255) { Value = DBNull.Value });

                        // Source + last modified - SEND TORONTO TIME TO DATABASE
                        cmd.Parameters.Add(new SqlParameter("@source", SqlDbType.NVarChar, 50) { Value = DbOrNull(r.Source) });
                        cmd.Parameters.Add(new SqlParameter("@last_modified_utc", SqlDbType.DateTime2) { Value = r.LastModifiedTorontoTime });

                        // Recurring flag
                        cmd.Parameters.Add(new SqlParameter("@is_recurring", SqlDbType.Bit) { Value = r.IsRecurring });

                        // Recipients (all meeting attendees for reference)
                        cmd.Parameters.Add(new SqlParameter("@recipients", SqlDbType.NVarChar, -1) { Value = DbOrNull(r.Recipients) });

                        // Status - default to 'submitted' if not specified
                        cmd.Parameters.Add(new SqlParameter("@status", SqlDbType.NVarChar, 20) { Value = DbOrNull(r.Status ?? "submitted") });

                        // ✅ Hours allocated (for multi-program time splits)
                        cmd.Parameters.Add(new SqlParameter("@hours_allocated", SqlDbType.Decimal)
                        {
                            Precision = 10,
                            Scale = 2,
                            Value = DbOrNull(r.HoursAllocated)
                        });

                        System.Diagnostics.Debug.WriteLine($"Executing stored procedure: dbo.Timesheet_Upsert");

                        var rc = await cmd.ExecuteNonQueryAsync().ConfigureAwait(false);

                        System.Diagnostics.Debug.WriteLine($"✅ Upsert completed - return code: {rc}");
                    }

                    // ✅ NEW: Verify the record was actually inserted/updated
                    using (var verify = new SqlCommand(@"
                        SELECT COUNT(*) 
                        FROM db_owner.ytimesheet 
                        WHERE global_id = @global_id 
                          AND start_utc = @start_utc 
                          AND user_name = @user_name", cn))
                    {
                        verify.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = globalId });
                        verify.Parameters.Add(new SqlParameter("@start_utc", SqlDbType.DateTime2) { Value = r.StartTorontoTime });
                        verify.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar, 100) { Value = DbOrNull(r.UserDisplayName) });

                        var count = (int)await verify.ExecuteScalarAsync().ConfigureAwait(false);

                        if (count > 0)
                        {
                            System.Diagnostics.Debug.WriteLine($"✅ VERIFIED: Found {count} record(s) in database");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine("⚠️ WARNING: Upsert executed but NO records found in database! This usually means a trigger or constraint prevented the insert.");
                        }
                    }
                }

                System.Diagnostics.Debug.WriteLine($"========== UPSERT SUCCESS ==========");
            }
            catch (SqlException ex)
            {
                var msg = $"SQL Error {ex.Number} (State {ex.State}): {ex.Message}\n\nProcedure: {ex.Procedure}\nLine: {ex.LineNumber}";
                System.Diagnostics.Debug.WriteLine($"❌ SQL EXCEPTION: {msg}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                throw;
            }
            catch (Exception ex)
            {
                var msg = $"Upsert failed for {r.UserDisplayName}: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"❌ EXCEPTION: {msg}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        /// <summary>
        /// Deletes a timesheet record for the given meeting and user
        /// ✅ UPDATED: Now includes job_code to support multi-program deletions
        /// Returns true if a record was deleted, false otherwise
        /// </summary>
        public static async Task<bool> DeleteTimesheetAsync(MeetingRecord r)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ConnString))
                {
                    throw new InvalidOperationException("ConnectionStrings['OemsDatabase'] is null/empty. Check App.config key.");
                }

                var globalId = !string.IsNullOrWhiteSpace(r.GlobalId) ? r.GlobalId : r.EntryId;
                if (string.IsNullOrWhiteSpace(globalId))
                {
                    throw new ArgumentException("global_id is required (use GlobalAppointmentID or EntryID).",
                                                nameof(r.GlobalId));
                }

                using (var cn = new SqlConnection(ConnString))
                {
                    await cn.OpenAsync().ConfigureAwait(false);

                    // ✅ CRITICAL FIX: Convert StartUtc (which is actual UTC from Outlook) to Toronto time for database matching
                    // The database stores times as if they were Toronto local time (despite the column name saying "utc")
                    var torontoTz = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                    var startTimeForQuery = TimeZoneInfo.ConvertTimeFromUtc(r.StartUtc, torontoTz);

                    System.Diagnostics.Debug.WriteLine($"DeleteTimesheetAsync: Parameters for DELETE:");
                    System.Diagnostics.Debug.WriteLine($"  GlobalId: '{globalId}'");
                    System.Diagnostics.Debug.WriteLine($"  StartUtc (from Outlook): {r.StartUtc:yyyy-MM-dd HH:mm:ss.fff}");
                    System.Diagnostics.Debug.WriteLine($"  Converted to Toronto time for DB: {startTimeForQuery:yyyy-MM-dd HH:mm:ss.fff}");
                    System.Diagnostics.Debug.WriteLine($"  UserDisplayName: {r.UserDisplayName}");
                    System.Diagnostics.Debug.WriteLine($"  ProgramCode: {r.ProgramCode}");

                    // ✅ DEBUG: Check what records EXIST in database with these parameters
                    using (var checkCmd = new SqlCommand(@"
                        SELECT COUNT(*) as cnt, 
                               MIN(start_utc) as earliest_start,
                               MAX(start_utc) as latest_start
                        FROM db_owner.ytimesheet
                        WHERE global_id = @global_id 
                          AND user_name = @user_name", cn))
                    {
                        checkCmd.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = globalId });
                        checkCmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar, 100) { Value = DbOrNull(r.UserDisplayName) });

                        using (var reader = await checkCmd.ExecuteReaderAsync().ConfigureAwait(false))
                        {
                            if (await reader.ReadAsync().ConfigureAwait(false))
                            {
                                var cnt = (int)reader["cnt"];
                                var earliestStart = reader["earliest_start"] != DBNull.Value ? reader["earliest_start"] : (object)null;
                                var latestStart = reader["latest_start"] != DBNull.Value ? reader["latest_start"] : (object)null;

                                System.Diagnostics.Debug.WriteLine($"DeleteTimesheetAsync: Found {cnt} records with this GlobalId+User");
                                if (cnt > 0)
                                {
                                    System.Diagnostics.Debug.WriteLine($"  Earliest start_utc: {earliestStart}");
                                    System.Diagnostics.Debug.WriteLine($"  Latest start_utc: {latestStart}");
                                }
                            }
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(r.ProgramCode))
                    {
                        // ✅ Use EXACT time match (±30 seconds) with Toronto time converted from UTC
                        var timeRangeStart = startTimeForQuery.AddSeconds(-30);
                        var timeRangeEnd = startTimeForQuery.AddSeconds(30);

                        System.Diagnostics.Debug.WriteLine($"DeleteTimesheetAsync: Time range: {timeRangeStart:yyyy-MM-dd HH:mm:ss} to {timeRangeEnd:yyyy-MM-dd HH:mm:ss}");

                        using (var cmd = new SqlCommand(@"
                            DELETE FROM db_owner.ytimesheet
                            WHERE global_id = @global_id 
                              AND start_utc >= @start_utc_min
                              AND start_utc <= @start_utc_max
                              AND user_name = @user_name 
                              AND job_code = @job_code", cn))
                        {
                            cmd.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = globalId });
                            cmd.Parameters.Add(new SqlParameter("@start_utc_min", SqlDbType.DateTime2) { Value = timeRangeStart });
                            cmd.Parameters.Add(new SqlParameter("@start_utc_max", SqlDbType.DateTime2) { Value = timeRangeEnd });
                            cmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar, 100) { Value = DbOrNull(r.UserDisplayName) });
                            cmd.Parameters.Add(new SqlParameter("@job_code", SqlDbType.NVarChar, 50) { Value = DbOrNull(r.ProgramCode) });

                            var rowsDeleted = await cmd.ExecuteNonQueryAsync().ConfigureAwait(false);

                            System.Diagnostics.Debug.WriteLine($"DeleteTimesheetAsync: Deleted {rowsDeleted} record(s) for program={r.ProgramCode}, global_id={globalId}");

                            return rowsDeleted > 0;
                        }
                    }
                    else
                    {
                        // Delete all records for this meeting (no specific program code)
                        using (var cmd = new SqlCommand("dbo.Timesheet_DeleteRecord", cn))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = globalId });
                            cmd.Parameters.Add(new SqlParameter("@start_utc", SqlDbType.DateTime2) { Value = startTimeForQuery });
                            cmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar, 100) { Value = DbOrNull(r.UserDisplayName) });

                            var rowsDeleted = await cmd.ExecuteNonQueryAsync().ConfigureAwait(false);

                            System.Diagnostics.Debug.WriteLine($"DeleteTimesheetAsync: Deleted {rowsDeleted} record(s) (all programs) for global_id={globalId}");

                            return rowsDeleted > 0;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DeleteTimesheetAsync failed: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Deletes ALL timesheet records for a given appointment (when meeting is cancelled)
        /// Returns the number of records deleted
        /// </summary>
        public static async Task<int> DeleteAllTimesheetsByGlobalIdAsync(string globalId, string entryId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ConnString))
                {
                    throw new InvalidOperationException("ConnectionStrings['OemsDatabase'] is null/empty. Check App.config key.");
                }

                var id = !string.IsNullOrWhiteSpace(globalId) ? globalId : entryId;
                if (string.IsNullOrWhiteSpace(id))
                {
                    // System.Diagnostics.Debug.WriteLine("DeleteAllTimesheetsByGlobalIdAsync: No valid ID provided");
                    return 0;
                }

                using (var cn = new SqlConnection(ConnString))
                {
                    await cn.OpenAsync().ConfigureAwait(false);

                    // ✅ NEW: Use stored procedure instead of inline DELETE
                    using (var cmd = new SqlCommand("dbo.Timesheet_DeleteAllByGlobalId", cn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = id });

                        var rowsDeleted = await cmd.ExecuteNonQueryAsync().ConfigureAwait(false);

                        // System.Diagnostics.Debug.WriteLine($"Deleted {rowsDeleted} timesheet record(s) for global_id={id} (all users)");

                        return rowsDeleted;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DeleteAllTimesheetsByGlobalIdAsync failed: {ex.Message}");
                return 0;  // Return 0 instead of throwing to avoid breaking the cancellation flow
            }
        }

        /// <summary>
        /// Marks a meeting as permanently ignored by setting status = 'ignored'
        /// This prevents it from showing up in the unsubmitted list
        /// Uses stored procedure dbo.Timesheet_IgnoreMeeting
        /// ✅ FIXED: Works with meetings that don't have GlobalAppointmentID yet (unsaved/unsent meetings)
        /// </summary>
        public static async Task<bool> IgnoreTimesheetAsync(MeetingRecord r)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ConnString))
                {
                    // System.Diagnostics.Debug.WriteLine("IgnoreTimesheetAsync: Connection string is null/empty");
                    throw new InvalidOperationException("ConnectionStrings['OemsDatabase'] is null/empty. Check App.config key.");
                }

                // ✅ CRITICAL FIX: Use EntryId as fallback if GlobalId is empty
                // New meetings (created via "New Online Meeting") don't have GlobalAppointmentID until sent
                var globalId = !string.IsNullOrWhiteSpace(r.GlobalId) ? r.GlobalId : r.EntryId;
                if (string.IsNullOrWhiteSpace(globalId))
                {
                    // System.Diagnostics.Debug.WriteLine("IgnoreTimesheetAsync: Both GlobalId AND EntryId are null/empty");
                    throw new ArgumentException("Either global_id or entry_id is required");
                }

                // ✅ ENHANCED DEBUG: Log all parameters BEFORE database call
                // System.Diagnostics.Debug.WriteLine($"IgnoreTimesheetAsync: Attempting to ignore meeting:");
                // System.Diagnostics.Debug.WriteLine($"  GlobalId: '{r.GlobalId}' (original)");
                // System.Diagnostics.Debug.WriteLine($"  EntryId: '{r.EntryId}' (fallback)");
                // System.Diagnostics.Debug.WriteLine($"  Using ID: '{globalId}' (final)");
                // System.Diagnostics.Debug.WriteLine($"  StartTorontoTime: {r.StartTorontoTime:yyyy-MM-dd HH:mm:ss}");
                // System.Diagnostics.Debug.WriteLine($"  EndTorontoTime: {r.EndTorontoTime:yyyy-MM-dd HH:mm:ss}");
                // System.Diagnostics.Debug.WriteLine($"  UserDisplayName: {r.UserDisplayName}");
                // System.Diagnostics.Debug.WriteLine($"  Subject: {r.Subject}");

                using (var cn = new SqlConnection(ConnString))
                {
                    await cn.OpenAsync().ConfigureAwait(false);
                    // System.Diagnostics.Debug.WriteLine($"IgnoreTimesheetAsync: Database connection opened");

                    // ✅ NEW: Use stored procedure instead of inline MERGE
                    using (var cmd = new SqlCommand("dbo.Timesheet_IgnoreMeeting", cn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = globalId });
                        cmd.Parameters.Add(new SqlParameter("@start_utc", SqlDbType.DateTime2) { Value = r.StartTorontoTime });
                        cmd.Parameters.Add(new SqlParameter("@end_utc", SqlDbType.DateTime2) { Value = r.EndTorontoTime });
                        cmd.Parameters.Add(new SqlParameter("@entry_id", SqlDbType.NVarChar, 255) { Value = DbOrNull(r.EntryId) });
                        cmd.Parameters.Add(new SqlParameter("@subject", SqlDbType.NVarChar, 500) { Value = DbOrNull(r.Subject) });
                        cmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar, 100) { Value = DbOrNull(r.UserDisplayName) });
                        cmd.Parameters.Add(new SqlParameter("@last_modified_utc", SqlDbType.DateTime2) { Value = r.LastModifiedTorontoTime });

                        // System.Diagnostics.Debug.WriteLine($"IgnoreTimesheetAsync: Executing stored procedure...");
                        var rowsAffected = await cmd.ExecuteNonQueryAsync().ConfigureAwait(false);

                        // System.Diagnostics.Debug.WriteLine($"IgnoreTimesheetAsync: Procedure completed - rows affected: {rowsAffected}");

                        // ✅ VERIFY: Check if record actually exists in database
                        using (var verifyCmd = new SqlCommand(@"
                            SELECT status FROM db_owner.ytimesheet
                            WHERE global_id = @global_id 
                              AND start_utc = @start_utc 
                              AND user_name = @user_name", cn))
                        {
                            verifyCmd.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = globalId });
                            verifyCmd.Parameters.Add(new SqlParameter("@start_utc", SqlDbType.DateTime2) { Value = r.StartTorontoTime });
                            verifyCmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar, 100) { Value = DbOrNull(r.UserDisplayName) });

                            var status = await verifyCmd.ExecuteScalarAsync().ConfigureAwait(false);

                            if (status != null)
                            {
                                // System.Diagnostics.Debug.WriteLine($"IgnoreTimesheetAsync: ✅ VERIFIED - Record exists with status='{status}'");
                                return true;
                            }
                            else
                            {
                                // System.Diagnostics.Debug.WriteLine($"IgnoreTimesheetAsync: ❌ WARNING - Procedure reported {rowsAffected} rows but record not found in verification!");
                                // System.Diagnostics.Debug.WriteLine($"  This usually means a constraint or trigger prevented the INSERT");
                                return false;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"IgnoreTimesheetAsync EXCEPTION: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"IgnoreTimesheetAsync Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Cancels a previously ignored timesheet (removes the 'ignored' status record)
        /// This allows the meeting to appear in the unsubmitted list again
        /// ✅ FIXED: Verify deletion instead of relying on stored procedure return value
        /// </summary>
        public static async Task<bool> CancelIgnoreTimesheetAsync(MeetingRecord r)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ConnString))
                {
                    throw new InvalidOperationException("ConnectionStrings['OemsDatabase'] is null/empty. Check App.config key.");
                }

                var globalId = !string.IsNullOrWhiteSpace(r.GlobalId) ? r.GlobalId : r.EntryId;
                if (string.IsNullOrWhiteSpace(globalId))
                {
                    throw new ArgumentException("global_id is required (use GlobalAppointmentID or EntryID).");
                }

                System.Diagnostics.Debug.WriteLine($"CancelIgnoreTimesheetAsync: Parameters:");
                System.Diagnostics.Debug.WriteLine($"  GlobalId (from record): '{r.GlobalId}'");
                System.Diagnostics.Debug.WriteLine($"  EntryId (from record): '{r.EntryId}'");
                System.Diagnostics.Debug.WriteLine($"  Using ID: '{globalId}'");
                System.Diagnostics.Debug.WriteLine($"  StartTorontoTime: {r.StartTorontoTime:yyyy-MM-dd HH:mm:ss}");
                System.Diagnostics.Debug.WriteLine($"  UserDisplayName: {r.UserDisplayName}");

                using (var cn = new SqlConnection(ConnString))
                {
                    await cn.OpenAsync().ConfigureAwait(false);

                    // ✅ FIX: Check if record exists BEFORE attempting delete
                    bool recordExistsBeforeDelete = false;
                    using (var checkCmd = new SqlCommand(@"
                        SELECT COUNT(*) 
                        FROM db_owner.ytimesheet
                        WHERE global_id = @global_id 
                          AND start_utc = @start_utc 
                          AND user_name = @user_name 
                          AND status = 'ignored'", cn))
                    {
                        checkCmd.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = globalId });
                        checkCmd.Parameters.Add(new SqlParameter("@start_utc", SqlDbType.DateTime2) { Value = r.StartTorontoTime });
                        checkCmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar, 100) { Value = DbOrNull(r.UserDisplayName) });

                        var count = (int)await checkCmd.ExecuteScalarAsync().ConfigureAwait(false);
                        recordExistsBeforeDelete = count > 0;

                        System.Diagnostics.Debug.WriteLine($"CancelIgnoreTimesheetAsync: Found {count} ignored record(s) BEFORE delete");
                    }

                    if (!recordExistsBeforeDelete)
                    {
                        System.Diagnostics.Debug.WriteLine($"CancelIgnoreTimesheetAsync: No ignored record found to delete");
                        return false;
                    }

                    // ✅ Execute the delete stored procedure
                    using (var cmd = new SqlCommand("dbo.Timesheet_CancelIgnore", cn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = globalId });
                        cmd.Parameters.Add(new SqlParameter("@start_utc", SqlDbType.DateTime2) { Value = r.StartTorontoTime });
                        cmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar, 100) { Value = DbOrNull(r.UserDisplayName) });

                        // ✅ NOTE: ExecuteNonQueryAsync may return -1 due to SQL Server bug with RETURN @@ROWCOUNT
                        // We'll verify deletion below instead of relying on this value
                        var rowsDeleted = await cmd.ExecuteNonQueryAsync().ConfigureAwait(false);

                        System.Diagnostics.Debug.WriteLine($"CancelIgnoreTimesheetAsync: Procedure executed - return value: {rowsDeleted} (may be -1 due to SQL bug)");
                    }

                    // ✅ VERIFY: Check if record was actually deleted by checking if it still exists
                    bool recordExistsAfterDelete = false;
                    using (var verifyCmd = new SqlCommand(@"
                        SELECT COUNT(*) 
                        FROM db_owner.ytimesheet
                        WHERE global_id = @global_id 
                          AND start_utc = @start_utc 
                          AND user_name = @user_name", cn))
                    {
                        verifyCmd.Parameters.Add(new SqlParameter("@global_id", SqlDbType.NVarChar, 255) { Value = globalId });
                        verifyCmd.Parameters.Add(new SqlParameter("@start_utc", SqlDbType.DateTime2) { Value = r.StartTorontoTime });
                        verifyCmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar, 100) { Value = DbOrNull(r.UserDisplayName) });

                        var count = (int)await verifyCmd.ExecuteScalarAsync().ConfigureAwait(false);
                        recordExistsAfterDelete = count > 0;

                        System.Diagnostics.Debug.WriteLine($"CancelIgnoreTimesheetAsync: Found {count} record(s) AFTER delete");
                    }

                    // ✅ Success if record no longer exists
                    bool deletionSuccessful = !recordExistsAfterDelete;

                    if (deletionSuccessful)
                    {
                        System.Diagnostics.Debug.WriteLine($"CancelIgnoreTimesheetAsync: ✅ SUCCESS - Record was deleted");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"CancelIgnoreTimesheetAsync: ❌ FAILED - Record still exists after delete attempt");
                    }

                    return deletionSuccessful;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CancelIgnoreTimesheetAsync EXCEPTION: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                return false;
            }
        }
    }
}