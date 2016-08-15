using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Management;
using System.Web.Mvc;
using System.Printing;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.Collections;
using EPSPrintMgmt.Models;
using Microsoft.Management.Infrastructure;
using System.Configuration;

namespace EPSPrintMgmt.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Printer()
        {
            ViewBag.Message = "Your Printer page.";
           

            return View(GetPrinters(@"\\mcd-eps-3"));
        }

        public ActionResult PrintJobs()
        {
            return View(GetPrintJobs(GetEPSServers()));
        }

        public ActionResult ThePrinters()
        {
            return View(GetPrintersWithPrintQueueClass(@"\\mcd-eps-3"));
        }

        static public List<Printer> GetPrinters(string server)
        {
            PrintServer printServer = new PrintServer(server);
            var myPrintQueues = printServer.GetPrintQueues().OrderBy(t => t.Name);
            List<Printer> printerList = new List<Printer>();
            foreach (PrintQueue pq in myPrintQueues)
            {
                pq.Refresh();
                printerList.Add(new Printer { Name=pq.Name,Driver=pq.QueueDriver.Name,PrintServer=pq.HostingPrintServer.Name});
            }
            return (printerList);
        }

        static public List<PrintQueue> GetPrintersWithPrintQueueClass(string server)
        {
            PrintServer printServer = new PrintServer(server);
            var myPrintQueues = printServer.GetPrintQueues().OrderBy(t => t.Name);
            List<PrintQueue> printerList = new List<PrintQueue>();
            List<PrintTest> printerListExt = new List<PrintTest>();
            foreach (PrintQueue pq in myPrintQueues)
            {
                pq.Refresh();
                printerList.Add(pq);
                //printerListExt.Add(pq);
            }
            return (printerList);
        }
        static public List<PrintJob> GetPrintJobs()
        {
            List<PrintJob> printJobs = new List<PrintJob>();
            string Namespace = @"root\cimv2";
            string OSQuery = "SELECT * FROM Win32_PrintJob";
            CimSession mySession = CimSession.Create("mcd-eps-3");
            IEnumerable<CimInstance> queryInstance = mySession.QueryInstances(Namespace, "WQL", OSQuery);
            foreach(var props in queryInstance)
            {
                var jobID = Int32.Parse(props.CimInstanceProperties["JobID"].Value.ToString());
                var printName = props.CimInstanceProperties["Name"].ToString();
                printJobs.Add(new PrintJob {PrintJobID=jobID,PrintJobName=printName });
            }
            return (printJobs);

        }

        static public List<PrintJob> GetPrintJobs(string printserver)
        {
            List<PrintJob> printJobs = new List<PrintJob>();
            string Namespace = @"root\cimv2";
            string OSQuery = "SELECT * FROM Win32_PrintJob";
            CimSession mySession = CimSession.Create(printserver);
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
                    Printer = Printer.Substring(0,Printer.LastIndexOf(','));
                    var PrintJobID = Int32.Parse(props.CimInstanceProperties["JobID"].Value.ToString());
                    var PrintJobName = props.CimInstanceProperties["Document"].Value.ToString();
                    var Server = props.CimInstanceProperties["HostPrintQueue"].Value.ToString();
                    var Status = props.CimInstanceProperties["Status"].Value.ToString();
                    var PrintDriver = props.CimInstanceProperties["DriverName"].Value.ToString();
                    //var TimeSubmitted = ManagementDateTimeConverter.ToDateTime(props.CimInstanceProperties["StartTime"].ToString());
                    var TimeSubmitted = Convert.ToDateTime(props.CimInstanceProperties["TimeSubmitted"].Value);
                    var PagesPrinted = Int32.Parse(props.CimInstanceProperties["PagesPrinted"].Value.ToString());
                    var TotalPages = Int32.Parse(props.CimInstanceProperties["TotalPages"].Value.ToString());

                    printJobs.Add(new PrintJob {
                        Printer=Printer,
                        PrintJobID = PrintJobID,
                        PrintJobName = PrintJobName,
                        Server=Server,
                        Status=Status,
                        PrintDriver=PrintDriver,
                        TimeSubmitted=TimeSubmitted,
                        PagesPrinted=PagesPrinted,
                        TotalPages=TotalPages

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
    }
}