using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPSPrintMgmt.Models
{
    public class JobsProcessingView
    {
        public Pager Pager { get; set; }
        public Hangfire.Storage.Monitoring.JobList<Hangfire.Storage.Monitoring.ProcessingJobDto> ProcessingJobs { get; set; }
        //public PrinterCreation SucceededJobs {get;set;}
    }
}