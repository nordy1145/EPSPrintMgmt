using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Printing;
using System.ComponentModel;

namespace EPSPrintMgmt.Models
{
    public class PrintJob
    {
        public string Server { get; set; }
        public string Printer { get; set; }
        [DisplayName ("Job ID")]
        public int PrintJobID { get; set; }
        public string Status { get; set; }
        [DisplayName ("Job Name")]
        public string PrintJobName { get; set; }
        [DisplayName("Print Driver")]
        public string PrintDriver { get; set; }
        [DisplayName("Time Job Submitted")]
        public DateTime? TimeSubmitted { get; set; }
        [DisplayName("Pages Printed")]
        public int PagesPrinted { get; set; }
        [DisplayName("Total Pages")]
        public int TotalPages { get; set; }
        [DisplayName("Source of Print Job")]
        public string HostPrintQueue { get; set; }
        [DisplayName("User")]
        public string Owner { get; set; }

        public void CancelPrintJob()
        {
            PrintServer pS = new PrintServer(this.Server);
            var myPrintQueue = pS.GetPrintQueues(); //.Where(t=>t.FullName.Contains("PRLPFC115HP"));
            var pq = myPrintQueue.Where(p => p.Name.Equals(this.Printer)).First();
            if (pq == null)
                return;
            pq.Refresh();
            var jobs = pq.GetPrintJobInfoCollection();
            var theJob = jobs.Where(j => j.JobIdentifier.Equals(this.PrintJobID)).First();
            if (theJob == null)
                return;
            //Console.WriteLine("Print Job ID " + theJob.JobIdentifier);
            //Console.WriteLine("Print Job ID " + theJob.Name);
            //Console.WriteLine("Print Job ID " + theJob.JobName);
            theJob.Cancel();
        }
    }

}