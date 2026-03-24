using System;

namespace OutlookAddIn1
{
    // Helper class for program allocation (shared across forms)
    public class ProgramAllocation
    {
        public string ProgramCode { get; set; }
        public string ActivityCode { get; set; }
        public string StageCode { get; set; }
        public double Hours { get; set; }
    }
}
