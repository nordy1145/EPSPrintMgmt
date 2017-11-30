using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPSPrintMgmt.Models
{
    public class JobsFailedVIew
    {
        public Pager Pager { get; set; }
        public Hangfire.Storage.Monitoring.JobList<Hangfire.Storage.Monitoring.FailedJobDto> FailedJobs { get; set; }
        //public PrinterCreation SucceededJobs {get;set;}
        public IEnumerable<string> Items { get; set; }
    }
    public class JobsFailed
    {
        //public string 
    }
}