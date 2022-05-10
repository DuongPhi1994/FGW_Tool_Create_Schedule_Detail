using System;
using System.Collections.Generic;
using System.Text;

namespace ScheduleFGR
{
    public class CheckResult
    {
        public string DuplicateCheck { get; set; }
        public string Result { get; set; }
        public CheckResult() { }
        public CheckResult(string duplicateCheck, string result)
        {
            this.DuplicateCheck = duplicateCheck;
            this.Result = result;
        }
    }
}
