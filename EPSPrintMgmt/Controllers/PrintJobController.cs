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
            return View(GetPrintJobs(GetAllPrintServers()));
        }

        public ActionResult Delete(int id, string printServer,string printer)
        {
            return View(GetPrintJob(printServer,printer,id));
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(PrintJob printJob)
        {
            CancelPrintJob(printJob.Server, printJob.Printer, printJob.PrintJobID);
            SendEmail("Canceled Print Job", "The following print job has been canceled: "+printJob.PrintJobName + Environment.NewLine + "Printer: "+printJob.Printer + Environment.NewLine+"Print server: " + printJob.Server + Environment.NewLine+"User: " + User.Identity.Name);
            return RedirectToAction("Index");
        }


        static public List<PrintJob> PrintJobs()
        {
            List<PrintJob> printJobs = new List<PrintJob>();
            string Namespace = @"root\cimv2";
            string OSQuery = "SELECT * FROM Win32_PrintJob";
            CimSession mySession = CimSession.Create("ryan-pc");
            IEnumerable<CimInstance> queryInstance = mySession.QueryInstances(Namespace, "WQL", OSQuery);
            foreach (var props in queryInstance)
            {
                var jobID = Int32.Parse(props.CimInstanceProperties["JobID"].Value.ToString());
                var printName = props.CimInstanceProperties["Name"].ToString();
                printJobs.Add(new PrintJob { PrintJobID = jobID, PrintJobName = printName });
            }

            return (printJobs);
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
                        Owner=Owner

                    });
                }
            }

            return (printJobs);
        }
        static public List<string> GetEPSServers()
        {
            List<string> epsServers = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EPS")).Select(k => ConfigurationManager.AppSettings[k]).ToList();
            return (epsServers);
        }
        static public List<string> GetAllPrintServers()
        {
            List<string> epsServers = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("Server")).Select(k => ConfigurationManager.AppSettings[k]).ToList();
            return (epsServers);
        }


        static public PrintJob GetPrintJob(string printserver,string printer, int printJobID)
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

            if (queryInstance.FirstOrDefault() ==null)
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

        static public void CancelPrintJob(string printServer, string printQueue, int printJob)
        {
            PrintServer pS = new PrintServer(@"\\"+printServer);
            var myPrintQueue = pS.GetPrintQueues(); //.Where(t=>t.FullName.Contains("PRLPFC115HP"));
            var pq = myPrintQueue.Where(p => p.Name.Equals(printQueue)).First();
            if (pq == null)
                return;
            pq.Refresh();
            var jobs = pq.GetPrintJobInfoCollection();
            var theJob = jobs.Where(j => j.JobIdentifier.Equals(printJob)).First();
            if (theJob == null)
                return;
            theJob.Cancel();
            pq.Refresh();
            //var pJ = myPrintQueue.GetJob(printJob);
            //Console.WriteLine(pJ.Name);
            //Console.WriteLine("Number of Pages before clearing: " + pJ.NumberOfPages);
            //pJ.Cancel();
        }

        static public string GetRelayServer()
        {
            string relayServer = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("MailRelay")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return (relayServer);
        }

        static public string GetEmailTo()
        {
            string relayServer = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EmailTo")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return (relayServer);
        }
        static public string GetEmailFrom()
        {
            string relayServer = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EmailFrom")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return (relayServer);
        }
        private static void SendEmail(string subject, string body)
        {
            MailMessage message = new MailMessage(GetEmailFrom(), GetEmailTo(), subject, body);

            SmtpClient mailClient = new SmtpClient(GetRelayServer());
            mailClient.Send(message);
        }
    }
}