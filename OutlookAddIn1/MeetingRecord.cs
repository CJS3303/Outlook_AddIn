using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddIn1
{
    public sealed class MeetingRecord
    {
        private static TimeZoneInfo _torontoTimeZone;
        
        private static TimeZoneInfo TorontoTimeZone
        {
            get
            {
                if (_torontoTimeZone == null)
                {
                    try
                    {
                        _torontoTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                    }
                    catch
                    {
                        _torontoTimeZone = TimeZoneInfo.Local;
                    }
                }
                return _torontoTimeZone;
            }
        }
        
        public string Source { get; set; }
        public string EntryId { get; set; }
        public string GlobalId { get; set; }
        public string Subject { get; set; }
        public DateTime StartUtc { get; set; }
        public DateTime EndUtc { get; set; }
        public double? HoursAllocated { get; set; }  // ✅ NEW: Explicit time allocation (for multi-program splits)
        public string ProgramCode { get; set; }
        public string ActivityCode { get; set; }
        public string StageCode { get; set; }
        public string UserDisplayName { get; set; }
        public DateTime LastModifiedUtc { get; set; }
        public bool IsRecurring { get; set; }
        public bool IsRecurringOccurrence { get; set; }  // ✅ NEW: Track if this is a recurring occurrence (vs master or non-recurring)
        public string Recipients { get; set; }  // Semicolon-separated list of ALL meeting attendees (for reference only)
        public string Status { get; set; }  // 'submitted' or 'ignored'
        
        // Helper properties for Toronto time (Eastern Time) - with error handling
        public DateTime StartTorontoTime
        {
            get
            {
                try
                {
                    return TimeZoneInfo.ConvertTimeFromUtc(StartUtc, TorontoTimeZone);
                }
                catch
                {
                    return StartUtc;  // Fallback to UTC if conversion fails
                }
            }
        }
        
        public DateTime EndTorontoTime
        {
            get
            {
                try
                {
                    return TimeZoneInfo.ConvertTimeFromUtc(EndUtc, TorontoTimeZone);
                }
                catch
                {
                    return EndUtc;  // Fallback to UTC if conversion fails
                }
            }
        }
        
        public DateTime LastModifiedTorontoTime
        {
            get
            {
                try
                {
                    return TimeZoneInfo.ConvertTimeFromUtc(LastModifiedUtc, TorontoTimeZone);
                }
                catch
                {
                    return LastModifiedUtc;  // Fallback to UTC if conversion fails
                }
            }
        }
    }
}
