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
            //Following params are used to determine the sort order of the list of printers.
            ViewBag.NameSortParm = String.IsNullOrEmpty(sortOrder) ? "Name_desc" : "";
            ViewBag.JobsSortParm = sortOrder == "Jobs" ? "Jobs" : "Jobs_desc";
            ViewBag.DriverSortParm = sortOrder == "Driver" ? "Driver_desc" : "Driver";

            //initialize the list of printers
            List<Printer> printers;

            //initlize some strings used throughout
            string theFirstEPSSErver = GetEPSServers().First();
            string checkPrintServer;
            string currentEPSServer = "";

            //Define checkPrintServer to determine if a print server param has been passed in or not.
            if (Session["currentPrintServerLookup"] != null)
            {
                checkPrintServer = Session["currentPrintServerLookup"].ToString();
            }
            else
            {
                checkPrintServer = null;
            }

            //This will determine what print server to actually use
            if (printServer == null)
            {
                //If no printer is passed in, return the first from the list.
                //The list pulls directly from the web.config file EPSServers section.
                if (checkPrintServer == null)
                {
                    Session["currentPrintServerLookup"] = theFirstEPSSErver;
                    currentEPSServer = theFirstEPSSErver;
                    printers = GetPrinters(theFirstEPSSErver);
                }
                //Checks to see if it's not the first EPS server in list, then returns the actual one.
                else if (checkPrintServer != theFirstEPSSErver)
                {
                    printers = GetPrinters(checkPrintServer);
                }
                //returns the first one if it's actually a passed in param.
                else
                {
                    printers = GetPrinters(theFirstEPSSErver);
                }
            }
            else
            //sets the session print server to be passed back when going back and forth between views.
            //Returns the print server passed into the view.
            {
                Session["currentPrintServerLookup"] = printServer;
                currentEPSServer = printServer;
                printers = GetPrinters(printServer);
            }
            // Gets all the print servers for the drop down option in the Index View of this controller.
            ViewData["printServers"] = GetAllPrintServers();

            //Switch statement to determine the sort order of the view.
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

            //Used to determine if the options tab will include specific edit options for EPS servers only.
            Session["IsEPSServer"] = GetEPSServers().Exists(s => s == currentEPSServer).ToString();

            //return the list of printers one way or another.
            return View(printers);
        }
        public ActionResult Error()
        {
            //Pass an error message from other views if there is some sort of error on the PrinterController.
            ViewBag.RedirectError = TempData["RedirectToError"];
            return View();
        }
        public ActionResult Success()
        {
            return View();
        }
        public ActionResult Create()
        {
            //Pass the print drivers from the Web.config file to the view
            ViewData["printDrivers"] = GetAllPrintDrivers();
            return View();
        }

        //This is used for AJAX version of creating a new EPS printer.
        public ActionResult CreateJsonRequest()
        {
            ViewData["printDrivers"] = GetAllPrintDrivers();
            return View();
        }

        //old non AJAX and non Parallel way of adding printers.
        //This should not be referenced by anything anymore.
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

        //Used for parallel and AJAX processing of new EPS printer creation.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public JsonResult CreatePrinterJSON([Bind(Include = "DeviceID,DriverName")]AddPrinterClass theNewPrinter)
        {
            //Initialize list to display to end users if it completes or not.
            List<string> outcome = new List<string>();

            //Make sure incoming data is good to go.
            if (ModelState.IsValid)
            {
                //Validate the printer exists in DNS.
                if (ValidHostname(theNewPrinter.DeviceID) == true)
                {
                    //kick off multiple threads to install printers quickly
                    Parallel.ForEach(GetEPSServers(), (server) =>
                    {
                        //Used to test if it can connect via Cim first.  If it cannot it skips that server.
                        CimSession mySession = CimSession.Create(server);
                        var testtheConnection = mySession.TestConnection();
                        if (testtheConnection == true)
                        {
                            //Start the process of installing a printer.
                            //Checks to see if the Printer port already exists on the server.
                            if (ExistingPrinterPort(theNewPrinter.DeviceID, server) == false)
                            {
                                //Adds printer port, currently just using the name of the printer for the port name.  So DNS must be setup first in order to add printer port.
                                AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.DeviceID, HostAddress = theNewPrinter.DeviceID, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                                AddNewPrinterPort(AddThePort, server);
                            }
                            else
                            {
                                //If printer port already exists, don't do anything!
                            }
                            //Try to add the printer now the port is defined on the server.
                            //first setup the props of the printer
                            AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.DeviceID, DriverName = theNewPrinter.DriverName, EnableBIDI = false, PortName = theNewPrinter.DeviceID, Published = false, Shared = false };
                            //Use the string return function to determine if the printer was successfully added or not.
                            outcome.Add(AddNewPrinterStringReturn(newPrinter, server));
                        }
                        //Cannot connect to the server, so send a message back to the user about it.
                        else
                        {
                            outcome.Add(server + "Is not a valid server.  Please contact the creator of this thing and have them check the web.config for valid EPS servers.");
                        }
                        //End the parallel processing.
                    });
                    //Email users from Web.Config to confirm everything went well!
                    SendEmail("New Printer Added to EPS", string.Join(Environment.NewLine, outcome) + Environment.NewLine + "Created by user: " + User.Identity.Name);
                    //Send a success message the Success View.
                    TempData["SuccessMessage"] = "Congrats, the printer installed correctly!  Enjoy your day.";
                    //return JSON results to the AJAX request of the view.
                    return Json(outcome, JsonRequestBehavior.AllowGet);
                }
                //Send email and return results if DNS does not exist for the printer.
                SendEmail("Failed EPS Install", "Printer: " + theNewPrinter.DeviceID + " failed the DNS lookup." + Environment.NewLine + "By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist.  Please try again.";
                outcome.Add("Hostname of the Printer does not exist.  Please try again.");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            //Return error message that 
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            outcome.Add("Something went wrong with the Model.  Please try again.");
            return Json(outcome, JsonRequestBehavior.AllowGet);

        }
        //Used to get options for a specific printer on a server.
        //The view currently allows to edit a print driver and clear a print queue.
        public ActionResult Options(string printer, string printServer)
        {
            //used to determine if the print server in question is an EPS server.
            //Only limited functionality for non EPS servers is currently defined.
            Session["IsEPSServer"] = GetEPSServers().Exists(s => s == printServer).ToString();
            //Return the printer for the View.
            return View(GetPrinter(printServer, printer));
        }

        //Post from the Options page, specifically for the Clear Print Queue option.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Options([Bind(Include = "Name,PrintServer")]Printer theNewPrinter)
        {
            if (ModelState.IsValid)
            {
                //attempts to clear the print queue.  Function returns true/false.
                if (ClearPrintQueue(theNewPrinter.Name, theNewPrinter.PrintServer) == true)
                {
                    //Send out notification and return success if function returns true.
                    SendEmail("Print Queue Cleared", "Printer: " + theNewPrinter.Name + " on server: " + theNewPrinter.PrintServer + " has been cleared successfully. by user: " + User.Identity.Name);
                    TempData["SuccessMessage"] = "Congrats, the print queue has been cleared!  Enjoy your day.";
                    return RedirectToAction("Success");
                }
                else
                {
                    //something went wrong and it couldn't clear the queue.
                    SendEmail("Print Queue failed to clear", "Printer: " + theNewPrinter.Name + " on server: " + theNewPrinter.PrintServer + " has failed to clear. by user: " + User.Identity.Name);
                    TempData["RedirectToError"] = "Could not clear print queue.  Please try again or logon to the server directly to clear it.";
                    return RedirectToAction("Error");
                }
            }
            //Shouldn't get this far, but returns an error if it does.
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            return RedirectToAction("Error");
        }

        public ActionResult Edit(string printer)
        {
            var myPrinter = GetPrinter(GetEPSServers().First(), printer);
            ViewData["printDrivers"] = GetAllPrintDrivers();
            return View(myPrinter);
        }
        public ActionResult EditEnterprisePrinter(string printer, string printServer)
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

        //Future use to edit a non EPS printer
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
                    return "Successfully deleted: " + i["Name"].ToString() + " on server: " + server;
                }

            }
            catch
            {
                return "Delete failed on server: " + server + ".  Try catch failed";

            }
            return "Delete failed on server: " + server + ". Returned failed at the end of the method.";
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
    }
}