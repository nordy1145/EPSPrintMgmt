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
using System.Printing.IndexedProperties;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Web.Caching;

namespace EPSPrintMgmt.Controllers
{
    public class PrinterController : Controller
    {
        // GET: Printer
        [OutputCache(Duration = 300, VaryByParam = "PrintServer;sortOrder")]
        public ActionResult Index(string printServer, string sortOrder)
        {
            ViewBag.NameSortParm = String.IsNullOrEmpty(sortOrder) ? "Name_desc" : "";
            ViewBag.JobsSortParm = sortOrder == "Jobs" ? "Jobs" : "Jobs_desc";
            ViewBag.DriverSortParm = sortOrder == "Driver" ? "Driver_desc" : "Driver";
            List<Printer> printers;
            string theFirstEPSSErver = GetEPSServers().First();
            string checkPrintServer;
            string currentEPSServer ="";

            if (Session["currentPrintServerLookup"] != null)
            {
                checkPrintServer = Session["currentPrintServerLookup"].ToString();
            }
            else
            {
                checkPrintServer = null;
            }

            if (printServer == null)
            {
                if (checkPrintServer == null)
                {
                    Session["currentPrintServerLookup"] = theFirstEPSSErver;
                    currentEPSServer = theFirstEPSSErver;
                    printers = GetPrinters(theFirstEPSSErver);
                }
                else if (checkPrintServer != theFirstEPSSErver)
                {
                    printers = GetPrinters(checkPrintServer);
                }
                else
                {
                    printers = GetPrinters(theFirstEPSSErver);
                }
            }
            else
            {
                Session["currentPrintServerLookup"] = printServer;
                currentEPSServer = printServer;
                printers = GetPrinters(printServer);
            }
            ViewData["printServers"] = GetAllPrintServers();

            switch (sortOrder)
            {
                case "Name_desc":
                    printers = printers.OrderByDescending(x => x.Name).ToList();
                    break;
                case "Jobs":
                    printers = printers.OrderBy(x => x.NumberJobs).ToList();
                    break;
                case "Jobs_desc":
                    printers = printers.OrderByDescending(x => x.NumberJobs).ToList();
                    break;
                case "Driver_desc":
                    printers = printers.OrderByDescending(x => x.Driver).ToList();
                    break;
                case "Driver":
                    printers = printers.OrderBy(x => x.Driver).ToList();
                    break;
                default:
                    printers = printers.OrderBy(x => x.Name).ToList();
                    break;
            }

            Session["IsEPSServer"] = GetEPSServers().Exists(s => s == currentEPSServer).ToString();
            var testingstuff = GetEPSServers().Exists(s => s == currentEPSServer).ToString();

            return View(printers);
        }
        public ActionResult Error()
        {
            ViewBag.RedirectError = TempData["RedirectToError"];
            return View();
        }
        public ActionResult Success()
        {
            return View();
        }
        public ActionResult Create()
        {
            ViewData["printDrivers"] = GetAllPrintDrivers();
            return View();
        }

        public ActionResult CreateJsonRequest()
        {
            ViewData["printDrivers"] = GetAllPrintDrivers();
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "DeviceID,DriverName")]AddPrinterClass theNewPrinter)
        {
            if (ModelState.IsValid)
            {
                if (ValidHostname(theNewPrinter.DeviceID) == true)
                {
                    foreach (var server in GetEPSServers())
                    {
                        if (ExistingPrinterPort(theNewPrinter.DeviceID, server) == false)
                        {
                            AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.DeviceID.ToUpper(), HostAddress = theNewPrinter.DeviceID, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                            AddNewPrinterPort(AddThePort, server);
                        }
                        else
                        {

                        }
                        AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.DeviceID.ToUpper(), DriverName = theNewPrinter.DriverName, EnableBIDI = false, PortName = theNewPrinter.DeviceID, Published = false, Shared = false };
                        AddNewPrinter(newPrinter, server);
                    }
                    SendEmail("New Printer Added to EPS", "Printer: " + theNewPrinter.DeviceID + " has been added successfully.  Created by user: " + User.Identity.Name);
                    TempData["SuccessMessage"] = "Congrats, the printer installed correctly!  Enjoy your day.";
                    return RedirectToAction("Success");
                }
                SendEmail("Failed EPS Install", "Printer: " + theNewPrinter.DeviceID + " failed the DNS lookup.  By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist.  Please try again.";
                return RedirectToAction("Error");
            }
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            return RedirectToAction("Error");

        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public JsonResult CreatePrinterJSON([Bind(Include = "DeviceID,DriverName")]AddPrinterClass theNewPrinter)
        {
            List<string> outcome = new List<string>();
            if (ModelState.IsValid)
            {
                if (ValidHostname(theNewPrinter.DeviceID) == true)
                {

                    Parallel.ForEach(GetEPSServers(), (server) =>
                    {
                        CimSession mySession = CimSession.Create(server);
                        var testtheConnection = mySession.TestConnection();
                        if (testtheConnection == true)
                        {
                            if (ExistingPrinterPort(theNewPrinter.DeviceID, server) == false)
                            {
                                AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.DeviceID, HostAddress = theNewPrinter.DeviceID, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                                AddNewPrinterPort(AddThePort, server);
                            }
                            else
                            {

                            }
                            AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.DeviceID, DriverName = theNewPrinter.DriverName, EnableBIDI = false, PortName = theNewPrinter.DeviceID, Published = false, Shared = false };
                            outcome.Add(AddNewPrinterStringReturn(newPrinter, server));

                        }
                        else
                        {
                            outcome.Add(server + "Is not a valid server.  Please contact the creator of this thing and have them check the web.config for valid EPS servers.");
                        }

                    });
                    //foreach (var server in GetEPSServers())
                    //{
                    //}
                    SendEmail("New Printer Added to EPS", string.Join(Environment.NewLine, outcome) + Environment.NewLine + "Created by user: " + User.Identity.Name);
                    TempData["SuccessMessage"] = "Congrats, the printer installed correctly!  Enjoy your day.";
                    return Json(outcome, JsonRequestBehavior.AllowGet);

                }
                SendEmail("Failed EPS Install", "Printer: " + theNewPrinter.DeviceID + " failed the DNS lookup." + Environment.NewLine + "By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist.  Please try again.";
                outcome.Add("Hostname of the Printer does not exist.  Please try again.");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            outcome.Add("Something went wrong with the Model.  Please try again.");
            return Json(outcome, JsonRequestBehavior.AllowGet);

        }

        public ActionResult Options(string printer, string printServer)
        {
            Session["IsEPSServer"] = GetEPSServers().Exists(s => s == printServer).ToString();
            return View(GetPrinter(printServer, printer));
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Options([Bind(Include = "Name,PrintServer")]Printer theNewPrinter)
        {
            if (ModelState.IsValid)
            {
                if (ClearPrintQueue(theNewPrinter.Name, theNewPrinter.PrintServer) == true)
                {
                    SendEmail("Print Queue Cleared", "Printer: " + theNewPrinter.Name + " on server: " + theNewPrinter.PrintServer + " has been cleared successfully. by user: " + User.Identity.Name);
                    TempData["SuccessMessage"] = "Congrats, the print queue has been cleared!  Enjoy your day.";
                    return RedirectToAction("Success");
                }
                else
                {
                    SendEmail("Print Queue failed to clear", "Printer: " + theNewPrinter.Name + " on server: " + theNewPrinter.PrintServer + " has failed to clear. by user: " + User.Identity.Name);
                    TempData["RedirectToError"] = "Could not clear print queue.  Please try again or logon to the server directly to clear it.";
                    return RedirectToAction("Error");
                }
            }
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            return RedirectToAction("Error");
        }
        public ActionResult Edit(string printer)
        {
            var myPrinter = GetPrinter(GetEPSServers().First(), printer);
            ViewData["printDrivers"] = GetAllPrintDrivers();
            return View(myPrinter);
        }
        public ActionResult EditEnterprisePrinter(string printer,string printServer)
        {
            var myPrinter = GetPrinter(printServer, printer);
            ViewData["printDrivers"] = GetAllPrintDrivers();
            return View(myPrinter);
        }
        public JsonResult EditPrinterJSON([Bind(Include = "Name,Driver")]Printer theNewPrinter)
        {
            List<string> outcome = new List<string>();
            if (ModelState.IsValid)
            {
                if (ValidHostname(theNewPrinter.Name) == true)
                {

                    Parallel.ForEach(GetEPSServers(), (server) =>
                    {
                        CimSession mySession = CimSession.Create(server);
                        var testtheConnection = mySession.TestConnection();
                        if (testtheConnection == true)
                        {
                            if (ExistingPrinterPort(theNewPrinter.Name, server) == false)
                            {
                                AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.Name, HostAddress = theNewPrinter.Name, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                                AddNewPrinterPort(AddThePort, server);
                            }
                            else
                            {

                            }
                            outcome.Add(DeletePrinter(theNewPrinter.Name, server));
                            AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.Name, DriverName = theNewPrinter.Driver, EnableBIDI = false, PortName = theNewPrinter.Name, Published = false, Shared = false };
                            outcome.Add(AddNewPrinterStringReturn(newPrinter, server));

                        }
                        else
                        {
                            outcome.Add(server + "Is not a valid server.  Please contact the creator of this thing and have them check the web.config for valid EPS servers.");
                        }

                    });
                    //foreach (var server in GetEPSServers())
                    //{
                    //}
                    SendEmail("EPS Printer Edited", string.Join(Environment.NewLine, outcome) + Environment.NewLine + "Created by user: " + User.Identity.Name);
                    TempData["SuccessMessage"] = "Congrats, the printer updated correctly!  Enjoy your day.";
                    return Json(outcome, JsonRequestBehavior.AllowGet);

                }
                SendEmail("Failed EPS Edit", "Printer: " + theNewPrinter.Name + " failed the DNS lookup." + Environment.NewLine + "By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist.  Please try again.";
                outcome.Add("Hostname of the Printer does not exist.  Please try again.");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            outcome.Add("Something went wrong with the Model.  Please try again.");
            return Json(outcome, JsonRequestBehavior.AllowGet);

        }
        public JsonResult EditEnterprisePrinterJSON([Bind(Include = "Name,Driver,PrintServer")]Printer theNewPrinter)
        {
            List<string> outcome = new List<string>();
            if (ModelState.IsValid)
            {
                if (ValidHostname(theNewPrinter.Name) == true)
                {

                        CimSession mySession = CimSession.Create(theNewPrinter.PrintServer);
                        var testtheConnection = mySession.TestConnection();
                        if (testtheConnection == true)
                        {
                            if (ExistingPrinterPort(theNewPrinter.Name, theNewPrinter.PrintServer) == false)
                            {
                                AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.Name, HostAddress = theNewPrinter.Name, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                                AddNewPrinterPort(AddThePort, theNewPrinter.PrintServer);
                            }
                            else
                            {

                            }
                            outcome.Add(DeletePrinter(theNewPrinter.Name, theNewPrinter.PrintServer));
                            AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.Name, DriverName = theNewPrinter.Driver, EnableBIDI = false, PortName = theNewPrinter.Name, Published = false, Shared = false };
                            outcome.Add(AddNewPrinterStringReturn(newPrinter, theNewPrinter.PrintServer));

                        }
                        else
                        {
                            outcome.Add(theNewPrinter.PrintServer + "Is not a valid server.  Please contact the creator of this thing and have them check the web.config for valid EPS servers.");
                        }

                    SendEmail("EPS Printer Edited", string.Join(Environment.NewLine, outcome) + Environment.NewLine + "Created by user: " + User.Identity.Name);
                    TempData["SuccessMessage"] = "Congrats, the printer updated correctly!  Enjoy your day.";
                    return Json(outcome, JsonRequestBehavior.AllowGet);

                }
                SendEmail("Failed EPS Edit", "Printer: " + theNewPrinter.Name + " failed the DNS lookup." + Environment.NewLine + "By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist.  Please try again.";
                outcome.Add("Hostname of the Printer does not exist.  Please try again.");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            outcome.Add("Something went wrong with the Model.  Please try again.");
            return Json(outcome, JsonRequestBehavior.AllowGet);

        }
        static public List<Printer> GetPrinters(string server)
        {
            PrintServer printServer = new PrintServer(@"\\" + server);
            var myPrintQueues = printServer.GetPrintQueues().OrderBy(t => t.Name);
            List<Printer> printerList = new List<Printer>();
            foreach (PrintQueue pq in myPrintQueues)
            {
                //pq.Refresh();
                printerList.Add(new Printer { Name = pq.Name, Driver = pq.QueueDriver.Name, PrintServer = pq.HostingPrintServer.Name.TrimStart('\\'), NumberJobs = pq.NumberOfJobs });
            }
            return (printerList);
        }
        static public Printer GetPrinter(string server, string printer)
        {
            PrintServer printServer = new PrintServer(@"\\" + server);
            var myPrintQueues = printServer.GetPrintQueue(printer);
            myPrintQueues.Refresh();
            Printer printerList = new Printer { Name = myPrintQueues.Name, Driver = myPrintQueues.QueueDriver.Name, PrintServer = myPrintQueues.HostingPrintServer.Name, NumberJobs = myPrintQueues.NumberOfJobs };
            return (printerList);
        }


        static private bool ExistingPrinterPort(string portName, string serverName)
        {
            string Namespace = @"root\cimv2";
            string OSQuery = "SELECT * FROM Win32_TCPIPPrinterPort";
            CimSession mySession = CimSession.Create(serverName);
            var testtheConnection = mySession.TestConnection();
            if (testtheConnection == true)
            {
                IEnumerable<CimInstance> queryInstance = mySession.QueryInstances(Namespace, "WQL", OSQuery);
                var exist = queryInstance.Where(p => p.CimInstanceProperties["Name"].Value.Equals(portName));
                if (exist.FirstOrDefault() != null)
                {
                    return true;
                }
            }
            else
            {
                return false;
            }
            return false;
        }

        static private void AddNewPrinterPort(AddPrinterPortClass theNewPrinterPort, string thePrintServer)
        {
            string Namespace = @"root\cimv2";
            string className = "Win32_TCPIPPrinterPort";
            CimInstance newPrinter = new CimInstance(className, Namespace);
            newPrinter.CimInstanceProperties.Add(CimProperty.Create("Name", theNewPrinterPort.Name, CimFlags.Any));
            newPrinter.CimInstanceProperties.Add(CimProperty.Create("SNMPEnabled", theNewPrinterPort.SNMPEnabled, CimFlags.Any));
            newPrinter.CimInstanceProperties.Add(CimProperty.Create("Protocol", theNewPrinterPort.Protocol, CimFlags.Any));
            newPrinter.CimInstanceProperties.Add(CimProperty.Create("PortNumber", theNewPrinterPort.PortNumber, CimFlags.Any));
            newPrinter.CimInstanceProperties.Add(CimProperty.Create("HostAddress", theNewPrinterPort.HostAddress, CimFlags.Any));

            CimSession Session = CimSession.Create(thePrintServer);
            CimInstance myPrinter = Session.CreateInstance(Namespace, newPrinter);
            myPrinter.Dispose();
            Session.Dispose();
        }

        static private void AddNewPrinter(AddPrinterClass theNewPrinter, string thePrintServer)
        {


            PrintServer printServer = new PrintServer(@"\\" + thePrintServer);
            PrintPropertyDictionary printProps = new PrintPropertyDictionary { };
            // Share the new printer using Remove/Add methods
            PrintBooleanProperty shared = new PrintBooleanProperty("IsShared", false);
            PrintBooleanProperty BIDI = new PrintBooleanProperty("EnableBIDI", false);
            PrintBooleanProperty published = new PrintBooleanProperty("Published", false);
            printProps.Add("IsShared", shared);
            printProps.Add("EnableBIDI", BIDI);
            printProps.Add("Published", published);
            String[] port = new String[] { theNewPrinter.DeviceID };


            try
            {
                PrintQueue AddingPrinterHere = printServer.InstallPrintQueue(theNewPrinter.DeviceID, theNewPrinter.DriverName, port, "WinPrint", printProps);
                printServer.Commit();
            }
            catch (System.Printing.PrintSystemException e)
            {

            }
            //printServer.Commit();
            printServer.Dispose();


        }
        static private string AddNewPrinterStringReturn(AddPrinterClass theNewPrinter, string thePrintServer)
        {


            PrintServer printServer = new PrintServer(@"\\" + thePrintServer);
            PrintPropertyDictionary printProps = new PrintPropertyDictionary { };
            // Share the new printer using Remove/Add methods
            PrintBooleanProperty shared = new PrintBooleanProperty("IsShared", false);
            PrintBooleanProperty BIDI = new PrintBooleanProperty("EnableBIDI", false);
            PrintBooleanProperty published = new PrintBooleanProperty("Published", false);
            printProps.Add("IsShared", shared);
            printProps.Add("EnableBIDI", BIDI);
            printProps.Add("Published", published);
            String[] port = new String[] { theNewPrinter.DeviceID };


            try
            {
                PrintQueue AddingPrinterHere = printServer.InstallPrintQueue(theNewPrinter.DeviceID, theNewPrinter.DriverName, port, "WinPrint", printProps);
                printServer.Commit();
            }
            catch (System.Printing.PrintSystemException e)
            {
                printServer.Dispose();
                return (theNewPrinter.DeviceID + " failed to install on " + thePrintServer + " with error message " + e.Message);
            }
            //printServer.Commit();
            printServer.Dispose();
            return (theNewPrinter.DeviceID + " was added successfully on " + thePrintServer);

        }
        static private string DeletePrinter(string printer, string server)
        {
            string Namespace = @"root\cimv2";
            var WMIQuery = new ObjectQuery("SELECT * FROM Win32_Printer WHERE DeviceID = '" + printer + "'");
            var conneOption = new ConnectionOptions
            {
                EnablePrivileges = true,
                Authentication = AuthenticationLevel.Packet,
                Impersonation = ImpersonationLevel.Impersonate
            };
            var scope = new ManagementScope(@"\\" + server + @"\" + Namespace, conneOption);
            ManagementObjectSearcher s = new ManagementObjectSearcher(scope, WMIQuery);
            ManagementObjectCollection col = s.Get();
            try
            {
                foreach (ManagementObject i in col)
                {
                    i.Delete();
                    return "Successfully deleted: "+i["Name"].ToString()+" on server: "+server;
                }

            }
            catch
            {
                return "Delete failed on server: "+server+".  Try catch failed";

            }
            return "Delete failed on server: "+server+". Returned failed at the end of the method.";
        }
        static private bool ClearPrintQueue(string printer, string server)
        {
            PrintServer printServer = new PrintServer(@"\\" + server);
            PrintQueue pq = new PrintQueue(printServer, printer, PrintSystemDesiredAccess.AdministratePrinter);
            //var myPrintQueues = printServer.GetPrintQueue(printer);
            try
            {
                pq.Purge();
                //myPrintQueues.Purge();
                return true;
            }
            catch
            {
                return false;
            }
        }

        static public List<string> GetEPSServers()
        {
            List<string> epsServers = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EPS")).Select(k => ConfigurationManager.AppSettings[k]).ToList();
            return (epsServers);
        }
        static public List<string> GetAllPrintServers()
        {
            List<string> allServers = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("Server")).Select(k => ConfigurationManager.AppSettings[k]).ToList();
            return (allServers);
        }

        static public List<string> GetAllPrintDrivers()
        {
            List<string> printDrivers = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("PrintDriver")).Select(k => ConfigurationManager.AppSettings[k]).ToList();
            return (printDrivers);
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
        private static bool ValidHostname(string hostname)
        {
            System.Net.IPHostEntry host;

            try
            {
                host = System.Net.Dns.GetHostEntry(hostname);
            }
            catch (System.Net.Sockets.SocketException e)
            {
                //Console.WriteLine(e.Message);
                return false;
            }
            return true;
        }
        //static public void AddPrinter(Printer printer)
        //{
        //    List<PrintJob> printJobs = new List<PrintJob>();
        //    string Namespace = @"root\cimv2";
        //    string OSQuery = "SELECT * FROM Win32_Printer";
        //    string className = "Win32_Printer";
        //    CimInstance newPrinter = new CimInstance(className, Namespace);
        //    newPrinter.CimInstanceProperties.Add(CimProperty.Create("Drivername", "HP Universal Printing PCL 5 (v5.5.0)",CimFlags.Any));
        //    newPrinter.CimInstanceProperties.Add(CimProperty.Create("PortName", "PBLMCISXEROX",CimFlags.Any));
        //    newPrinter.CimInstanceProperties.Add(CimProperty.Create("Shared", false, CimFlags.Any));
        //    newPrinter.CimInstanceProperties.Add(CimProperty.Create("Published", false, CimFlags.Any));
        //    newPrinter.CimInstanceProperties.Add(CimProperty.Create("EnableBIDI", false, CimFlags.Any));
        //    newPrinter.CimInstanceProperties.Add(CimProperty.Create("DeviceID", "TestingStuff", CimFlags.Key));

        //    CimSession mySession = CimSession.Create("ryan-pc");



        //    //CimInstance printInstance = new CimInstance(className, Namespace);
        //    //CimInstance updateInstance = mySession.CreateInstance(Namespace,printInstance);
        //    //updateInstance.CimInstanceProperties.Add();
        //    //updateInstance.

        //    //IEnumerable<CimInstance> queryInstance = mySession.QueryInstances(Namespace, "WQL", OSQuery);
        //    //foreach (var props in queryInstance)
        //    //{
        //    //    var jobID = Int32.Parse(props.CimInstanceProperties["JobID"].Value.ToString());
        //    //    var printName = props.CimInstanceProperties["Name"].ToString();
        //    //    printJobs.Add(new PrintJob { PrintJobID = jobID, PrintJobName = printName });
        //    //}

        //    //return (printJobs);
        //}
    }
}