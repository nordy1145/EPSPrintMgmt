using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Printing;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.Collections;
using EPSPrintMgmt.Models;
using Microsoft.Management.Infrastructure;
using System.Configuration;
using System.Net.Mail;

namespace EPSPrintMgmt.Controllers
{

    public class PrintJobController : Controller
    {
        // GET: PrintJob
        public ActionResult Index()
        {
            return View(GetPrintJobs(Support.GetAllPrintServers()));
        }
        public ActionResult IndexWPurge()
        {
            return View(GetPrintJobs(Support.GetAllPrintServers()));
        }
        public ActionResult Delete(int id, string printServer, string printer)
        {
            return View(GetPrintJob(printServer, printer, id));
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(PrintJob printJob)
        {
            if (Support.AdditionalSecurity() == true)
            {
                var theADGroup = Support.ADGroupCanPurgePrintQueues();
                bool isInRole = User.IsInRole(theADGroup);
                if (isInRole == false)
                {
                    Support.SendEmail("Failed cancel print job", "User :" + User.Identity.Name.ToString() + " attempted to purge print jobs and failed because user does not have access.");
                    return RedirectToAction("Index");
                }
            }

            if (CancelPrintJob(printJob.Server, printJob.Printer, printJob.PrintJobID))
            {
                Support.SendEmail("Canceled Print Job", "The following print job has been canceled: " + printJob.PrintJobName + Environment.NewLine + "Printer: " + printJob.Printer + Environment.NewLine + "Print server: " + printJob.Server + Environment.NewLine + "User: " + User.Identity.Name);
            }
            else
            {
                Support.SendEmail("Canceled Print Job", "The following print job failed to cancel: " + printJob.PrintJobName + Environment.NewLine + "Printer: " + printJob.Printer + Environment.NewLine + "Print server: " + printJob.Server + Environment.NewLine + "User: " + User.Identity.Name);
            }
            return RedirectToAction("Index");
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteMultipleJobs(List<PrintJob> printjobs)
        {
            if (Support.AdditionalSecurity() == true)
            {
                var theADGroup = Support.ADGroupCanPurgePrintQueues();
                bool isInRole = User.IsInRole(theADGroup);
                if (isInRole == false)
                {
                    Support.SendEmail("Failed cancel print job","User :"+User.Identity.Name.ToString()+" attempted to purge print jobs and failed because user does not have access.");
                    return RedirectToAction("Index");
                }
            }


            List<string> outcome = new List<string>();
            foreach (var pj in printjobs)
            {
                if (pj.ToDelete)
                {
                if(CancelPrintJob(pj.Server, pj.Printer, pj.PrintJobID))
                    {
                        outcome.Add("The following print job has been canceled: " + pj.PrintJobName + Environment.NewLine + "Printer: " + pj.Printer + Environment.NewLine + "Print server: " + pj.Server + Environment.NewLine + "User: " + User.Identity.Name + Environment.NewLine + Environment.NewLine);
                    }
                    else
                    {
                        outcome.Add("The following print job failed to cancel: " + pj.PrintJobName + Environment.NewLine + "Printer: " + pj.Printer + Environment.NewLine + "Print server: " + pj.Server + Environment.NewLine + "User: " + User.Identity.Name + Environment.NewLine + Environment.NewLine);
                    }
                }

            }
            Support.SendEmail("Canceled Print Jobs", string.Concat(outcome));
            return RedirectToAction("Index");
        }
        static public List<PrintJob> GetPrintJobs(List<string> printserver)
        {
            List<PrintJob> printJobs = new List<PrintJob>();
            string Namespace = @"root\cimv2";
            string OSQuery = "SELECT * FROM Win32_PrintJob";
            foreach (string ps in printserver)
            {
                CimSession mySession = CimSession.Create(ps);
                IEnumerable<CimInstance> queryInstance = mySession.QueryInstances(Namespace, "WQL", OSQuery);
                try
                {
                    foreach (var props in queryInstance)
                    {
                        var Printer = props.CimInstanceProperties["Name"].Value.ToString();
                        Printer = Printer.Substring(0, Printer.LastIndexOf(','));
                        var PrintJobID = Int32.Parse(props.CimInstanceProperties["JobID"].Value.ToString());
                        var PrintJobName = props.CimInstanceProperties["Document"].Value.ToString();
                        var HostPrintQueue = props.CimInstanceProperties["HostPrintQueue"].Value.ToString();
                        var Status = props.CimInstanceProperties["Status"].Value.ToString();
                        var PrintDriver = props.CimInstanceProperties["DriverName"].Value.ToString();
                        //var TimeSubmitted = ManagementDateTimeConverter.ToDateTime(props.CimInstanceProperties["StartTime"].ToString());
                        var TimeSubmitted = Convert.ToDateTime(props.CimInstanceProperties["TimeSubmitted"].Value);
                        var PagesPrinted = Int32.Parse(props.CimInstanceProperties["PagesPrinted"].Value.ToString());
                        var TotalPages = Int32.Parse(props.CimInstanceProperties["TotalPages"].Value.ToString());
                        var Owner = props.CimInstanceProperties["Owner"].Value.ToString();

                        printJobs.Add(new PrintJob
                        {
                            Printer = Printer,
                            PrintJobID = PrintJobID,
                            PrintJobName = PrintJobName,
                            Server = ps,
                            Status = Status,
                            PrintDriver = PrintDriver,
                            TimeSubmitted = TimeSubmitted,
                            PagesPrinted = PagesPrinted,
                            TotalPages = TotalPages,
                            HostPrintQueue = HostPrintQueue,
                            Owner = Owner

                        });
                    }

                }
                catch
                {

                }
            }

            return (printJobs);
        }
        static public PrintJob GetPrintJob(string printserver, string printer, int printJobID)
        {
            PrintJob printJob = new PrintJob();
            string Namespace = @"root\cimv2";
            string OSQuery = "SELECT * FROM Win32_PrintJob";
            string FullPrinterName = printer + ", " + printJobID;
            CimSession mySession = CimSession.Create(printserver);
            IEnumerable<CimInstance> queryInstance = mySession.QueryInstances(Namespace, "WQL", OSQuery).Where(p => p.CimInstanceProperties["Name"].Value.Equals(FullPrinterName));
            //queryInstance.Where(p => p.CimInstanceProperties["JobID"].Equals(printJobID));
            var testing = queryInstance.Where(p => p.CimInstanceProperties["Name"].Value.Equals(FullPrinterName));
            //queryInstance.Where(p => p.CimInstanceProperties["JobID"].Equals(printJobID)).FirstOrDefault();

            //var testing = from k in queryInstance
            //              where k.CimInstanceProperties["JobID"] == printJobID
            //              select k;

            if (queryInstance.FirstOrDefault() == null)
            {
                return (printJob);
            }

            var props = queryInstance.FirstOrDefault();
            var Printer = props.CimInstanceProperties["Name"].Value.ToString();
            Printer = Printer.Substring(0, Printer.LastIndexOf(','));
            var PrintJobID = Int32.Parse(props.CimInstanceProperties["JobID"].Value.ToString());
            var PrintJobName = props.CimInstanceProperties["Document"].Value.ToString();
            var HostPrintQueue = props.CimInstanceProperties["HostPrintQueue"].Value.ToString();
            var Status = props.CimInstanceProperties["Status"].Value.ToString();
            var PrintDriver = props.CimInstanceProperties["DriverName"].Value.ToString();
            //var TimeSubmitted = ManagementDateTimeConverter.ToDateTime(props.CimInstanceProperties["StartTime"].ToString());
            var TimeSubmitted = Convert.ToDateTime(props.CimInstanceProperties["TimeSubmitted"].Value);
            var PagesPrinted = Int32.Parse(props.CimInstanceProperties["PagesPrinted"].Value.ToString());
            var TotalPages = Int32.Parse(props.CimInstanceProperties["TotalPages"].Value.ToString());
            var Owner = props.CimInstanceProperties["Owner"].Value.ToString();

            printJob.HostPrintQueue = HostPrintQueue;
            printJob.Server = printserver;
            printJob.Printer = Printer;
            printJob.PrintJobID = PrintJobID;
            printJob.PrintJobName = PrintJobName;
            printJob.Status = Status;
            printJob.PrintDriver = PrintDriver;
            printJob.TimeSubmitted = TimeSubmitted;
            printJob.PagesPrinted = PagesPrinted;
            printJob.TotalPages = TotalPages;
            printJob.Owner = Owner;
            //foreach (var props in queryInstance)
            //{
            //    var jobID = Int32.Parse(props.CimInstanceProperties["JobID"].Value.ToString());
            //    var printName = props.CimInstanceProperties["Name"].ToString();
            //    printJob.Add(new PrintJob { PrintJobID = jobID, PrintJobName = printName });
            //}
            return (printJob);
        }
        static public bool CancelPrintJob(string printServer, string printQueue, int printJob)
        {
            try
            {
            PrintServer pS = new PrintServer(@"\\" + printServer);
            var myPrintQueue = pS.GetPrintQueues(); //.Where(t=>t.FullName.Contains("PRLPFC115HP"));
            var pq = myPrintQueue.Where(p => p.Name.Equals(printQueue)).First();
            pq.Refresh();

            var jobs = pq.GetPrintJobInfoCollection();
            var theJob = jobs.Where(j => j.JobIdentifier.Equals(printJob)).First();
            theJob.Cancel();
            pq.Refresh();
                return true;
            }
            catch
            {
                return false;
            }

        }

    }
}