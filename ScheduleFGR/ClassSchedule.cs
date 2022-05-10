using System;
using System.Collections.Generic;
using System.Text;

namespace ScheduleFGR
{
    public class ClassSchedule
    {
        public string Term { get; set; }
        public List<SlotSchedule> SlotSchedules { get; set; }
        public string Room { get; set; }
        public string SubjectCode { get; set; }
        public string StudentGroup { get; set; }
        public int Type { get; set; }
        public int TotalSlot { get; set; }
        public DateTime StartDate { get; set; }
        public List<DateTime> ListDayOff { get; set; }
        public DateTime StartDatePart2 { get; set; }
        //sap xep tu 1-2-3, mon so 1 hoc truoc
        public int OrdinalSubject { get; set; }
        public List<ScheduleDetail> ScheduleDetails { get; set; }

    }
}
