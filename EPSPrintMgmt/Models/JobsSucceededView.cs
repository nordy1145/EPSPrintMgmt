using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPSPrintMgmt.Models
{
    public class JobsSucceededView
    {
        public Pager Pager { get; set; }
        public IEnumerable<PrinterCreation> SucceededJobs { get; set; }
        //public PrinterCreation SucceededJobs {get;set;}
        public IEnumerable<string> Items { get; set; }
    }
    public class JobsSucceed
    {
        //public string 
    }
}