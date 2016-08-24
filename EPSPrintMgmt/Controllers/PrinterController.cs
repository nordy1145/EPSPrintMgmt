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
using System.Net;

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

            //initialize some strings used throughout
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
            ViewData["useIP"] = UsePrinterIPAddr();
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
        public ActionResult Create([Bind(Include = "DeviceID,DriverName,PortName")]AddPrinterClass theNewPrinter)
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
        public JsonResult CreatePrinterJSON([Bind(Include = "DeviceID,DriverName,PortName")]AddPrinterClass theNewPrinter)
        {
            //Initialize list to display to end users if it completes or not.
            List<string> outcome = new List<string>();

            //Make sure incoming data is good to go.
            if (ModelState.IsValid)
            {
                //Validate the printer exists in DNS or Valid IP Address
                //Web.Config determines if it is really used or not.
                if (ValidHostname(theNewPrinter.PortName) == true)
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
                            if (ExistingPrinterPort(theNewPrinter.PortName, server) == false)
                            {
                                //Adds printer port, currently just using the name of the printer for the port name.  So DNS must be setup first in order to add printer port.
                                AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.PortName, HostAddress = theNewPrinter.PortName, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                                AddNewPrinterPort(AddThePort, server);
                            }
                            else
                            {
                                //If printer port already exists, don't do anything!
                            }
                            if (CheckCurrentPrinter(theNewPrinter.DeviceID, server))
                            {
                                outcome.Add(theNewPrinter.DeviceID + " already exists on " + server);
                            }
                            else
                            {
                                //Try to add the printer now the port is defined on the server.
                                //first setup the props of the printer
                                AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.DeviceID, DriverName = theNewPrinter.DriverName, EnableBIDI = false, PortName = theNewPrinter.PortName, Published = false, Shared = false };
                                //Use the string return function to determine if the printer was successfully added or not.
                                outcome.Add(AddNewPrinterStringReturn(newPrinter, server));
                            }
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
                SendEmail("Failed EPS Install", "Printer: " + theNewPrinter.DeviceID + " failed the DNS lookup or IP Address validation." + Environment.NewLine + "By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist or it is an invalid IP Address.  Please try again.";
                outcome.Add("Hostname of the Printer does not exist or it is an invalid IP Address.  Please try again.");
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
        //Returns view to Edit a specific EPS Printer
        public ActionResult Edit(string printer)
        {
            //Just uses the first EPS print server to get the properties.
            var myPrinter = GetPrinter(GetEPSServers().First(), printer);
            //Return a list of available print drivers from web.config
            ViewData["printDrivers"] = GetAllPrintDrivers();
            //Determines if an PortName field should be returned to users.
            ViewData["useIP"] = UsePrinterIPAddr();
            //return view with printer info.
            return View(myPrinter);
        }
        //Will be used at some point to edit non EPS printers, potentially...
        public ActionResult EditEnterprisePrinter(string printer, string printServer)
        {
            var myPrinter = GetPrinter(printServer, printer);
            ViewData["printDrivers"] = GetAllPrintDrivers();
            return View(myPrinter);
        }
        //AJAX Json response for editing an EPS Printer.
        //Currently deletes 
        public JsonResult EditPrinterJSON([Bind(Include = "Name,Driver,PortName")]Printer theNewPrinter)
        {
            //Initialize string for the output.
            List<string> outcome = new List<string>();

            if (ModelState.IsValid)
            {
                //Verify the printer name is still correct.
                if (ValidHostname(theNewPrinter.PortName) == true)
                {
                    //Start parallel processing on each EPS server.
                    Parallel.ForEach(GetEPSServers(), (server) =>
                    {
                        //Verify a Cim Session can be created before calling methods that use it... Maybe a bit backwards.
                        CimSession mySession = CimSession.Create(server);
                        var testtheConnection = mySession.TestConnection();
                        if (testtheConnection == true)
                        {
                            //Continue on if a session can be created
                            if (ExistingPrinterPort(theNewPrinter.PortName, server) == false)
                            {
                                //Check to see if port doesn't exist and create it if it does not exist.
                                AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.PortName, HostAddress = theNewPrinter.PortName, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                                AddNewPrinterPort(AddThePort, server);
                            }
                            else
                            {
                                //don't do anything if port is actually there.
                            }
                            //Need to delete out the old printer first, since I haven't found a good way to change print drivers/properities.
                            outcome.Add(DeletePrinter(theNewPrinter.Name, server));
                            //create the new printer object with the correct props.
                            AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.Name, DriverName = theNewPrinter.Driver, EnableBIDI = false, PortName = theNewPrinter.PortName, Published = false, Shared = false };
                            //Try and add the printer back in after it's been deleted.
                            outcome.Add(AddNewPrinterStringReturn(newPrinter, server));

                        }
                        else
                        {
                            //Print server doesn't exist or is done.  Need to check web.config currently.
                            outcome.Add(server + "Is not a valid server.  Please contact the creator of this thing and have them check the web.config for valid EPS servers.");
                        }

                    });
                    //Finish the Parallel loop and return the results.
                    SendEmail("EPS Printer Edited", string.Join(Environment.NewLine, outcome) + Environment.NewLine + "Created by user: " + User.Identity.Name);
                    TempData["SuccessMessage"] = "Congrats, the printer updated correctly!  Enjoy your day.";
                    return Json(outcome, JsonRequestBehavior.AllowGet);

                }
                SendEmail("Failed EPS Edit", "Printer: " + theNewPrinter.Name + " failed the DNS lookup or IP address validation." + Environment.NewLine + "By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist or it is not a valid IP address.  Please try again.";
                outcome.Add("Hostname of the Printer does not exist or it is not a valid IP address.  Please try again.");
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
        //Method to return all printers from a specified print server
        static public List<Printer> GetPrinters(string server)
        {
            //Currently using the PrintServer class instead of a WMI query.
            //Was roughly 3-4 times faster to use PrintServer instead of WMI on 1400 printers.  Still takes about 10 seconds though.
            //PrintServer class requires the 2 wacks in the server name.
            PrintServer printServer = new PrintServer(@"\\" + server);
            //Make a pretty ordered list by name.
            var myPrintQueues = printServer.GetPrintQueues().OrderBy(t => t.Name);
            //Create an empty list to return.
            List<Printer> printerList = new List<Printer>();
            //Loop through and add all items to the custom class.
            foreach (PrintQueue pq in myPrintQueues)
            {
                //pq.Refresh();
                printerList.Add(new Printer { Name = pq.Name, Driver = pq.QueueDriver.Name, PrintServer = pq.HostingPrintServer.Name.TrimStart('\\'), NumberJobs = pq.NumberOfJobs, PortName = pq.QueuePort.Name });
            }
            //return the printers added to the custom class.
            return (printerList);
        }
        //Get details for a specific printer.
        static public Printer GetPrinter(string server, string printer)
        {
            //Currently using the PrintServer class instead of a WMI query.
            //Was roughly 3-4 times faster to use PrintServer instead of WMI on 1400 printers.  Still takes about 10 seconds though.
            //PrintServer class requires the 2 wacks in the server name.
            PrintServer printServer = new PrintServer(@"\\" + server);
            //Get the one print queue from the print server.
            var myPrintQueues = printServer.GetPrintQueue(printer);
            //refresh and add the print queue to the custom class.
            myPrintQueues.Refresh();
            Printer printerList = new Printer { Name = myPrintQueues.Name, Driver = myPrintQueues.QueueDriver.Name, PrintServer = myPrintQueues.HostingPrintServer.Name, NumberJobs = myPrintQueues.NumberOfJobs, PortName = myPrintQueues.QueuePort.Name };
            return (printerList);
        }
        //Return a true/false if printer port is active.  Need to know when adding a printer.
        static private bool ExistingPrinterPort(string portName, string serverName)
        {
            //This uses Powershell CimSession and WMI to query the information from the server.
            string Namespace = @"root\cimv2";
            string OSQuery = "SELECT * FROM Win32_TCPIPPrinterPort";
            CimSession mySession = CimSession.Create(serverName);
            //Verify the session created correctly, otherwise it will error out if it fails to connect.
            var testtheConnection = mySession.TestConnection();
            if (testtheConnection == true)
            {
                //Query WMI for Printer Ports on the server.
                IEnumerable<CimInstance> queryInstance = mySession.QueryInstances(Namespace, "WQL", OSQuery);
                //Check to see if it exists in the query response.
                var exist = queryInstance.Where(p => p.CimInstanceProperties["Name"].Value.Equals(portName));
                //If the port exists, then return true.
                if (exist.FirstOrDefault() != null)
                {
                    //return true and exit.
                    return true;
                }
            }
            else
            {
                //session didn't connect.  Return False
                return false;
            }
            //if you've gotten this far, something went really wrong.
            return false;
        }

        //Add a printer port to a server.
        //Assumes you have verified the print server and ip/dns entries are valid.
        static private void AddNewPrinterPort(AddPrinterPortClass theNewPrinterPort, string thePrintServer)
        {
            //Uses Powershell and WMI to create the new printer port.
            string Namespace = @"root\cimv2";
            string className = "Win32_TCPIPPrinterPort";

            //Create the CimInstance for the new printer port. Items for a WMI query really.
            CimInstance newPrinter = new CimInstance(className, Namespace);
            newPrinter.CimInstanceProperties.Add(CimProperty.Create("Name", theNewPrinterPort.Name, CimFlags.Any));
            newPrinter.CimInstanceProperties.Add(CimProperty.Create("SNMPEnabled", theNewPrinterPort.SNMPEnabled, CimFlags.Any));
            newPrinter.CimInstanceProperties.Add(CimProperty.Create("Protocol", theNewPrinterPort.Protocol, CimFlags.Any));
            newPrinter.CimInstanceProperties.Add(CimProperty.Create("PortNumber", theNewPrinterPort.PortNumber, CimFlags.Any));
            newPrinter.CimInstanceProperties.Add(CimProperty.Create("HostAddress", theNewPrinterPort.HostAddress, CimFlags.Any));

            //Create the Cimsession to the print server.
            CimSession Session = CimSession.Create(thePrintServer);
            //Actually create the printer port on the print server.
            CimInstance myPrinter = Session.CreateInstance(Namespace, newPrinter);
            //Cleanup
            myPrinter.Dispose();
            Session.Dispose();
        }

        //This shouldn't be used anymore.
        //old method for adding a printer.
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

        //Add a printer and return success or failure in a string.
        static private string AddNewPrinterStringReturn(AddPrinterClass theNewPrinter, string thePrintServer)
        {


            PrintServer printServer = new PrintServer(@"\\" + thePrintServer);
            PrintPropertyDictionary printProps = new PrintPropertyDictionary { };
            // Share the new printer using Remove/Add methods
            PrintBooleanProperty shared = new PrintBooleanProperty("IsShared", true);
            PrintBooleanProperty BIDI = new PrintBooleanProperty("EnableBIDI", false);
            PrintBooleanProperty isBIDI = new PrintBooleanProperty("IsBidiEnabled", false);
            PrintBooleanProperty published = new PrintBooleanProperty("Published", false);
            PrintBooleanProperty direct = new PrintBooleanProperty("IsDirect", false);
            PrintBooleanProperty spoolFirst = new PrintBooleanProperty("ScheduleCompletedJobsFirst", true);
            PrintBooleanProperty doComplete = new PrintBooleanProperty("DoCompleteFirst", true);
            printProps.Add("IsShared", shared);
            printProps.Add("EnableBIDI", BIDI);
            printProps.Add("IsBidiEnabled", isBIDI);
            printProps.Add("Published", published);
            printProps.Add("IsDirect", direct);
            printProps.Add("DoCompleteFirst", doComplete);
            printProps.Add("ScheduleCompletedJobsFirst", spoolFirst);
            String[] port = new String[] { theNewPrinter.PortName };


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
        static public bool CheckCurrentPrinter(string printer, string printserver)
        {
            //PrintServer class requires the 2 wacks in the server name.
            PrintServer checkprintServer = new PrintServer(@"\\" + printserver);
            //Get the one print queue from the print server.
            try
            {
                var myPrintQueues = checkprintServer.GetPrintQueue(printer);
                //myPrintQueues.Refresh();

            }
            catch
            {
                checkprintServer.Dispose();
                return false;
            }
            //refresh and add the print queue to the custom class.
            //Printer printerList = new Printer { Name = myPrintQueues.Name, Driver = myPrintQueues.QueueDriver.Name, PrintServer = myPrintQueues.HostingPrintServer.Name, NumberJobs = myPrintQueues.NumberOfJobs, PortName = myPrintQueues.QueuePort.Name };
            //return (printerList);
            return true;
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
        static public bool UsePrinterIPAddr()
        {
            string useIPAddr = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("UsePrinterIPAddress")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(useIPAddr, "true", true) == 0)
            {
                return true;
            }
            return false;
        }
        private static void SendEmail(string subject, string body)
        {
            MailMessage message = new MailMessage(GetEmailFrom(), GetEmailTo(), subject, body);

            SmtpClient mailClient = new SmtpClient(GetRelayServer());
            mailClient.Send(message);
        }

        //Currently checks to validate it's an actual IP address or a valid DNS entry.
        //Doesn't Ping the IP address, just makes sure it's a valid IPV4 or IPV6 address.
        private static bool ValidHostname(string hostname)
        {
            IPHostEntry host;
            IPAddress address;

            if (IPAddress.TryParse(hostname, out address))
            {
                switch (address.AddressFamily)
                {
                    case System.Net.Sockets.AddressFamily.InterNetwork:
                        return true;

                    case System.Net.Sockets.AddressFamily.InterNetworkV6:
                        return true;

                    default:
                        // umm... yeah... I'm going to need to take your red packet and...
                        return false;
                }
            }


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