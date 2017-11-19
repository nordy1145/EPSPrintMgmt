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
using System.IO;
using System.Xml;
using System.Diagnostics;
using System.Web.Hosting;
using System.Threading;
using Microsoft.Win32;

namespace EPSPrintMgmt.Controllers
{
    public class PrinterController : Controller
    {
        //Setup for logging.
        readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        // GET: Printer
        //[OutputCache(Duration = 300, VaryByParam = "PrintServer;sortOrder")]
        public ActionResult Index(string printServer, string sortOrder)
        {
            //Following params are used to determine the sort order of the list of printers.
            ViewBag.NameSortParm = String.IsNullOrEmpty(sortOrder) ? "Name_desc" : "";
            ViewBag.JobsSortParm = sortOrder == "Jobs" ? "Jobs" : "Jobs_desc";
            ViewBag.DriverSortParm = sortOrder == "Driver" ? "Driver_desc" : "Driver";

            //initialize the list of printers
            List<Printer> printers;

            //initialize some strings used throughout
            string theFirstEPSSErver = Support.GetEPSServers().First();
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
                    //Used to determine if the options tab will include specific edit options for EPS servers only.
                    Session["IsEPSServer"] = Support.GetEPSServers().Exists(s => s == currentEPSServer).ToString();
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
                //Used to determine if the options tab will include specific edit options for EPS servers only.
                Session["IsEPSServer"] = Support.GetEPSServers().Exists(s => s == currentEPSServer).ToString();
            }
            // Gets all the print servers for the drop down option in the Index View of this controller.
            ViewData["printServers"] = Support.GetAllPrintServers();

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
            Session["IsEPSServer"] = Support.GetEPSServers().Exists(s => s == currentEPSServer).ToString();

            //Used to determine if the view should have links for the Printer name aka DNS or Printer IP aka Printer Port
            ViewData["useIP"] = Support.UsePrinterIPAddr();

            //Used to determine if the view should show the number of print jobs
            ViewData["ShowPrintJobs"] = Support.ShowNumberPrintJobs();

            logger.Info("User: " + User.Identity.Name.ToString() + " has viewed all printers for server " + Session["currentPrintServerLookup"] + ".");

            //return the list of printers one way or another.
            return View(printers);
        }
        public ActionResult CanAddEPSAndEntPrinters()
        {
            if (Support.AddEPSAndEnterprisePrinters() == true)
            {
                return Content("true");
            }
            else
            {
                return Content("false");
            }
            //return Content();
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
        public ActionResult Search()
        {
            // Gets all the print servers for the drop down option in the Index View of this controller.
            ViewData["printServers"] = Support.GetAllPrintServers();
            return View();
        }
        public ActionResult Create()
        {
            //Pass the print drivers from the Web.config file to the view
            ViewData["printDrivers"] = Support.GetAllPrintDrivers();
            ViewData["useIP"] = Support.UsePrinterIPAddr();
            ViewData["useTrays"] = Support.UsePrintTrays();
            ViewData["getTrays"] = Support.GetTrays();
            ViewData["getGold"] = Support.GetEPSGoldPrinters();
            ViewData["useGold"] = Support.UseEPSGoldPrinter();
            ViewData["cloneDevSettings"] = Support.clonePrinterDeviceSettings();
            return View(new AddPrinterClass{cloneDevSettings=false });
        }
        public ActionResult CreateENT()
        {
            if (Support.EditEnterprisePrinters() == false)
            {
                return RedirectToAction("Index", "Home", "");
            }
            //Pass the print drivers from the Web.config file to the view
            ViewData["printDrivers"] = Support.GetAllEnterprisePrintDrivers();
            ViewData["useIP"] = Support.EnterpriseUsePrinterIPAddr();
            ViewData["useTrays"] = Support.UsePrintTrays();
            ViewData["getTrays"] = Support.GetTrays();
            ViewData["getEntServers"] = Support.GetEnterprisePrintServers();
            ViewData["getGold"] = Support.GetEntGoldPrinters();
            ViewData["useGold"] = Support.UseEntGoldPrinter();
            ViewData["cloneDevSettings"] = Support.cloneEntPrinterDeviceSettings();
            return View(new AddPrinterClass { cloneDevSettings = false });
        }
        public ActionResult CreateEPSAndEnt()
        {
            if (Support.AddEPSAndEnterprisePrinters() == false)
            {
                return RedirectToAction("Index", "Home", "");
            }
            //EPS Info
            //Pass the print drivers from the Web.config file to the view
            ViewData["EPSprintDrivers"] = Support.GetAllPrintDrivers();
            ViewData["EPSuseIP"] = Support.UsePrinterIPAddr();
            ViewData["EPSuseTrays"] = Support.UsePrintTrays();
            ViewData["EPSgetTrays"] = Support.GetTrays();
            ViewData["EPSgetGold"] = Support.GetEPSGoldPrinters();
            ViewData["EPSuseGold"] = Support.UseEPSGoldPrinter();
            ViewData["cloneDevSettings"] = Support.clonePrinterDeviceSettings();

            //Pass the print drivers from the Web.config file to the view
            ViewData["ENTprintDrivers"] = Support.GetAllEnterprisePrintDrivers();
            ViewData["ENTuseIP"] = Support.EnterpriseUsePrinterIPAddr();
            ViewData["ENTuseTrays"] = Support.UseEnterprisePrintTrays();
            ViewData["ENTgetTrays"] = Support.GetTrays();
            ViewData["ENTgetEntServers"] = Support.GetEnterprisePrintServers();
            ViewData["ENTgetGold"] = Support.GetEntGoldPrinters();
            ViewData["ENTuseGold"] = Support.UseEntGoldPrinter();
            ViewData["ENTcloneDevSettings"] = Support.cloneEntPrinterDeviceSettings();
            return View(new AddEPSandENTPrinterClass {cloneDevSettings=false,cloneENTDevSettings=false });
        }
        //Used for parallel and AJAX processing of new EPS printer creation.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public JsonResult CreatePrinterJSON([Bind(Include = "DeviceID,DriverName,PortName,Tray,Comments,Location,SourcePrinter,cloneDevSettings")]AddPrinterClass theNewPrinter)
        {
            //Do some timing on the whole process.
            Stopwatch newwatch = new Stopwatch();
            newwatch.Start();

            //Initialize list to display to end users if it completes or not.
            List<string> outcome = new List<string>();

            //Used to determine if user has correct access to create a printer.
            //If they don't then it returns an error message for them.
            if (!Support.IsUserAuthorized(Support.ADGroupCanAddEPSPrinter()))
            {
                outcome.Add("You are not authorized to add printers.");
                logger.Info("User " + User.Identity.Name + " attempted to create EPS printer: " + theNewPrinter.DeviceID);
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }

            //Make sure incoming data is good to go.
            if (ModelState.IsValid)
            {
                //Validate the printer exists in DNS or Valid IP Address
                //Web.Config determines if it is really used or not.
                if (!Support.ValidatePrinterDNS() || Support.ValidHostname(theNewPrinter.PortName) == true)
                {
                    //kick off multiple threads to install printers quickly
                    Parallel.ForEach(Support.GetEPSServers(), server =>
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
                                logger.Info("New Port added.  Port Name : " + theNewPrinter.PortName + " on server " + server + " by user " + User.Identity.Name);
                            }
                            else
                            {
                                logger.Info("New Port attempted to be added but already exists.  Port Name : " + theNewPrinter.PortName + " on server " + server + " by user " + User.Identity.Name);
                                //If printer port already exists, don't do anything!
                            }
                            if (CheckCurrentPrinter(theNewPrinter.DeviceID, server))
                            {
                                outcome.Add(theNewPrinter.DeviceID + " already exists on " + server);
                                logger.Info("New Printer attempted to be added but already exists.  Printer Name : " + theNewPrinter.DeviceID + " on server " + server + " by user " + User.Identity.Name);
                            }
                            else
                            {
                                //Try to add the printer now the port is defined on the server.
                                //first setup the props of the printer
                                AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.DeviceID, DriverName = theNewPrinter.DriverName, EnableBIDI = false, PortName = theNewPrinter.PortName, Published = false, Shared = false, SourcePrinter = theNewPrinter.SourcePrinter };
                                //Use the string return function to determine if the printer was successfully added or not.
                                outcome.Add(AddNewPrinterStringReturn(newPrinter, server));
                                logger.Info("New Printer added.  Printer Name : " + theNewPrinter.DeviceID + " on server " + server + " by user " + User.Identity.Name);
                                if (Support.UseEPSGoldPrinter() && theNewPrinter.SourcePrinter != null)
                                {
                                    var clonePrinterSettings = false;
                                    if(Support.clonePrinterDeviceSettings()&& theNewPrinter.cloneDevSettings)
                                    {
                                        clonePrinterSettings = true;
                                    }
                                    var cloneOutcome = clonePrintQueue(newPrinter, server, Support.GetEPSGoldPrintServer(), clonePrinterSettings);
                                    outcome.Add(cloneOutcome);
                                    logger.Info("Printer Name: " + theNewPrinter.DeviceID + " print settings outcome: " + cloneOutcome);
                                }

                                //Commented out the following items since changing print trays is really a bad idea this way.....

                                //if (Support.UsePrintTrays())
                                //{
                                //    outcome.Add(SetPrinterTray(theNewPrinter.DeviceID, server, theNewPrinter.Tray));
                                //    logger.Info("Print Tray Set for Printer Name : " + theNewPrinter.DeviceID + " Tray Info: " + theNewPrinter.Tray + " on server " + server + " by user " + User.Identity.Name + " time to install" + newwatch.ElapsedMilliseconds + "ms");
                                //}
                            }
                        }
                        //Cannot connect to the server, so send a message back to the user about it.
                        else
                        {
                            outcome.Add(server + " is not an active or a valid server.  Please verify the server is up or configured correctly in the web.config file.");
                            logger.Info("Current Print Server is not active or invalid.  Print Server: " + server + " by user " + User.Identity.Name);
                        }
                        mySession.Dispose();
                        //End the parallel processing.
                    });
                    logger.Info("Finished Adding and going to send email out now.  Milliseconds taken: " + newwatch.ElapsedMilliseconds);
                    newwatch.Stop();
                    //Email users from Web.Config to confirm everything went well!
                    Support.SendEmail("New Printer Added to EPS", string.Join(Environment.NewLine, outcome) + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to install." + Environment.NewLine + "Created by user: " + User.Identity.Name);
                    //Send a success message the Success View.
                    TempData["SuccessMessage"] = "Congrats, the printer installed correctly!  Enjoy your day.";
                    //return JSON results to the AJAX request of the view.
                    outcome.Add(@"<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + @" ms to install.</h5>");
                    return Json(outcome, JsonRequestBehavior.AllowGet);
                }
                newwatch.Stop();
                //Send email and return results if DNS does not exist for the printer.
                Support.SendEmail("Failed EPS Install", "Printer: " + theNewPrinter.DeviceID + " failed the DNS lookup or IP Address validation." + Environment.NewLine + "By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist or it is an invalid IP Address.  Please try again.";
                outcome.Add("Hostname of the Printer does not exist or it is an invalid IP Address.  Please try again.");
                outcome.Add("<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to attempt to install." + @"</h5>");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            newwatch.Stop();
            outcome.Add("`r`n" + newwatch.ElapsedMilliseconds + " ms to install.");
            //Return error message that 
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            outcome.Add("Something went wrong with the Model.  Please try again.");
            logger.Error("Failed to initialize the model");
            return Json(outcome, JsonRequestBehavior.AllowGet);

        }
        //Used for parallel and AJAX processing of new Enterprise printer creation.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public JsonResult CreateEnterprisePrinterJSON([Bind(Include = "DeviceID,DriverName,PortName,Tray,Comments,Location,PrintServer,SourcePrinter,cloneDevSettings")]AddPrinterClass theNewPrinter)
        {
            //Do some timing on the whole process.
            Stopwatch newwatch = new Stopwatch();
            newwatch.Start();

            //Initialize list to display to end users if it completes or not.
            List<string> outcome = new List<string>();

            //Used to determine if user has correct access to create a printer.
            //If they don't then it returns an error message for them.
            if (!Support.IsUserAuthorized(Support.ADGroupCanAddEnterprisePrinter()))
            {
                outcome.Add("You are not authorized to add enterprise printers.");
                logger.Info("User " + User.Identity.Name + " attempted to create Enterprise printer: " + theNewPrinter.DeviceID);
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }

            //Make sure incoming data is good to go.
            if (ModelState.IsValid)
            {
                //Validate the printer exists in DNS or Valid IP Address
                //Web.Config determines if it is really used or not.
                if (!Support.ValidatePrinterDNS() || Support.ValidHostname(theNewPrinter.PortName) == true)
                {
                    //Used to test if it can connect via Cim first.  If it cannot it skips that server.
                    CimSession mySession = CimSession.Create(theNewPrinter.PrintServer);
                    var testtheConnection = mySession.TestConnection();
                    if (testtheConnection == true)
                    {
                        //Start the process of installing a printer.
                        //Checks to see if the Printer port already exists on the server.
                        if (ExistingPrinterPort(theNewPrinter.PortName, theNewPrinter.PrintServer) == false)
                        {
                            //Adds printer port, currently just using the name of the printer for the port name.  So DNS must be setup first in order to add printer port.
                            AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.PortName, HostAddress = theNewPrinter.PortName, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                            AddNewPrinterPort(AddThePort, theNewPrinter.PrintServer);
                            logger.Info("New Port added.  Port Name : " + theNewPrinter.PortName + " on server " + theNewPrinter.PrintServer + " by user " + User.Identity.Name);
                        }
                        else
                        {
                            logger.Info("New Port attempted to be added but already exists.  Port Name : " + theNewPrinter.PortName + " on server " + theNewPrinter.PrintServer + " by user " + User.Identity.Name);
                            //If printer port already exists, don't do anything!
                        }
                        if (CheckCurrentPrinter(theNewPrinter.DeviceID, theNewPrinter.PrintServer))
                        {
                            outcome.Add(theNewPrinter.DeviceID + " already exists on " + theNewPrinter.PrintServer);
                            logger.Info("New Printer attempted to be added but already exists.  Printer Name : " + theNewPrinter.DeviceID + " on server " + theNewPrinter.PrintServer + " by user " + User.Identity.Name);
                        }
                        else
                        {
                            //Try to add the printer now the port is defined on the server.
                            //first setup the props of the printer
                            AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.DeviceID, DriverName = theNewPrinter.DriverName, EnableBIDI = false, PortName = theNewPrinter.PortName, Published = false, Shared = false,SourcePrinter=theNewPrinter.SourcePrinter };
                            PrintServer thePrintServer = new PrintServer(@"\\" + theNewPrinter.PrintServer, PrintSystemDesiredAccess.AdministrateServer);
                            //Use the string return function to determine if the printer was successfully added or not.
                            outcome.Add(AddNewEnterprisePrinterStringReturn(newPrinter, mySession, thePrintServer));
                            logger.Info("New Printer added.  Printer Name : " + theNewPrinter.DeviceID + " on server " + theNewPrinter.PrintServer + " by user " + User.Identity.Name);

                            if (Support.UseEntGoldPrinter() && theNewPrinter.SourcePrinter != null)
                            {
                                var clonePrinterSettings = false;
                                if (Support.cloneEntPrinterDeviceSettings() && theNewPrinter.cloneDevSettings)
                                {
                                    clonePrinterSettings = true;
                                }
                                var cloneOutcome = clonePrintQueue(newPrinter, theNewPrinter.PrintServer, Support.GetEntGoldPrintServer(), clonePrinterSettings);
                                outcome.Add(cloneOutcome);
                                logger.Info("Printer Name: " + theNewPrinter.DeviceID + " print settings outcome: " + cloneOutcome);
                            }

                            //if (Support.UsePrintTrays())
                            //{
                            //    outcome.Add(SetPrinterTray(theNewPrinter.DeviceID, theNewPrinter.PrintServer, theNewPrinter.Tray));
                            //    logger.Info("Print Tray Set for Printer Name : " + theNewPrinter.DeviceID + " Tray Info: " + theNewPrinter.Tray + " on server " + theNewPrinter.PrintServer + " by user " + User.Identity.Name + " time to install" + newwatch.ElapsedMilliseconds + "ms");
                            //}
                            thePrintServer.Dispose();
                        }
                    }
                    //Cannot connect to the server, so send a message back to the user about it.
                    else
                    {
                        outcome.Add(theNewPrinter.PrintServer + " is not an active or a valid server.  Please verify the server is up or configured correctly in the web.config file.");
                        logger.Info("Current Print Server is not active or invalid.  Print Server: " + theNewPrinter.PrintServer + " by user " + User.Identity.Name);
                    }
                    mySession.Dispose();
                    //End the parallel processing.

                    newwatch.Stop();
                    //Email users from Web.Config to confirm everything went well!
                    Support.SendEnterpriseEmail("New Printer Added to Enterprise Print Server.", string.Join(Environment.NewLine, outcome) + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to install." + Environment.NewLine + "Created by user: " + User.Identity.Name);
                    //Send a success message the Success View.
                    TempData["SuccessMessage"] = "Congrats, the printer installed correctly!  Enjoy your day.";
                    //return JSON results to the AJAX request of the view.
                    outcome.Add(@"<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + @" ms to install.</h5>");
                    return Json(outcome, JsonRequestBehavior.AllowGet);
                }
                newwatch.Stop();
                //Send email and return results if DNS does not exist for the printer.
                Support.SendEnterpriseEmail("Failed EPS Install", "Printer: " + theNewPrinter.DeviceID + " failed the DNS lookup or IP Address validation." + Environment.NewLine + "By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist or it is an invalid IP Address.  Please try again.";
                outcome.Add("Hostname of the Printer does not exist or it is an invalid IP Address.  Please try again.");
                outcome.Add("<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to attempt to install." + @"</h5>");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            newwatch.Stop();
            outcome.Add("`r`n" + newwatch.ElapsedMilliseconds + " ms to install.");
            //Return error message that 
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            outcome.Add("Something went wrong with the Model.  Please try again.");
            logger.Error("Failed to initialize the model");
            return Json(outcome, JsonRequestBehavior.AllowGet);

        }

        //Used for parallel and AJAX processing of new EPS and Enterprise printer creation.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public JsonResult CreateEPSAndEnterprisePrinterJSON([Bind(Include = "DeviceID,DriverName,PortName,Tray,Comments,Location,PrintServer,SourceServer,SourcePrinter,ENTDriverName,ENTPortName,ENTSourcePrinter,ENTSourceServer,ENTShared,ENTPublished,ENTEnableBIDI,ENTTray,cloneDevSettings,cloneEntDevSettings")]AddEPSandENTPrinterClass theNewPrinter)
        {
            //Do some timing on the whole process.
            Stopwatch newwatch = new Stopwatch();
            newwatch.Start();

            //Initialize list to display to end users if it completes or not.
            List<string> outcome = new List<string>();

            //Used to determine if user has correct access to create a printer.
            //If they don't then it returns an error message for them.
            if (!Support.IsUserAuthorized(Support.ADGroupCanAddEPSPrinter()) || !Support.IsUserAuthorized(Support.ADGroupCanAddEnterprisePrinter()))
            {
                outcome.Add("You are not authorized to add EPS and Enterprise printers.");
                logger.Info("User " + User.Identity.Name + " attempted to create EPS and Enterprise printer: " + theNewPrinter.DeviceID);
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }

            //Make sure incoming data is good to go.
            if (ModelState.IsValid)
            {
                //Validate the printer exists in DNS or Valid IP Address
                //Web.Config determines if it is really used or not.
                if (!Support.ValidatePrinterDNS() || Support.ValidHostname(theNewPrinter.PortName) == true)
                {
                    logger.Debug(newwatch.ElapsedMilliseconds + " ms has passed.  Before Parallel foreach");
                    var allPrintServers = Support.GetEPSServers();
                    allPrintServers.Add(theNewPrinter.PrintServer);
                    //kick off multiple threads to install printers quickly for EPS Print Servers.
                    Parallel.ForEach(allPrintServers, server =>
                    {
                        logger.Debug("Before 1st CimSession creation: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                        //Used to test if it can connect via Cim first.  If it cannot it skips that server.
                        CimSession mySession = CimSession.Create(server);
                        var testtheConnection = mySession.TestConnection();
                        if (testtheConnection == true)
                        {
                            logger.Debug("CimSession creation completed Successfully: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                            //Start the process of installing a printer.
                            //Checks to see if the Printer port already exists on the server.

                            //Start with Enterprise Print Servers.  Have different parameters passed for Enterprise Print Servers.
                            //Checks the Printer Port info
                            if (Support.GetEnterprisePrintServers().Contains(server))
                            {
                                logger.Debug("Before ENT Port Creation.  Testing port : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                if (ExistingPrinterPort(theNewPrinter.ENTPortName, server) == false)
                                {
                                    logger.Debug("After ENT Port Check and Before Port Creation: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    //Adds printer port, currently just using the name of the printer for the port name.  So DNS must be setup first in order to add printer port.
                                    AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.ENTPortName, HostAddress = theNewPrinter.ENTPortName, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                                    AddNewPrinterPort(AddThePort, server);
                                    logger.Debug("After ENT Port Creation: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    logger.Info("New Port added.  Port Name : " + theNewPrinter.ENTPortName + " on server " + server + " by user " + User.Identity.Name);
                                }
                                else
                                {
                                    logger.Debug("After ENT Port Check with No Ports to add : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    logger.Info("New Port attempted to be added but already exists.  Port Name : " + theNewPrinter.ENTPortName + " on server " + server + " by user " + User.Identity.Name);
                                    //If printer port already exists, don't do anything!
                                }
                            }
                            //Check the EPS Print Servers Port Info
                            else
                            {
                                logger.Debug("Before EPS Port Creation.  Testing port : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                if (ExistingPrinterPort(theNewPrinter.PortName, server) == false)
                                {
                                    logger.Debug("After EPS Port Check and Before Port Creation: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    //Adds printer port, currently just using the name of the printer for the port name.  So DNS must be setup first in order to add printer port.
                                    AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.PortName, HostAddress = theNewPrinter.PortName, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                                    AddNewPrinterPort(AddThePort, server);
                                    logger.Debug("After EPS Port Creation: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    logger.Info("New Port added.  Port Name : " + theNewPrinter.PortName + " on server " + server + " by user " + User.Identity.Name);
                                }
                                else
                                {
                                    logger.Debug("After EPS Port Check with No Ports to add : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    logger.Info("New Port attempted to be added but already exists.  Port Name : " + theNewPrinter.PortName + " on server " + server + " by user " + User.Identity.Name);
                                    //If printer port already exists, don't do anything!
                                }
                            }

                            logger.Debug("Before Testing if printer exists : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                            if (CheckCurrentPrinter(theNewPrinter.DeviceID, server))
                            {
                                logger.Debug("After Testing if printer exists and nothing to do: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                outcome.Add(theNewPrinter.DeviceID + " already exists on " + server);
                                logger.Info("New Printer attempted to be added but already exists.  Printer Name : " + theNewPrinter.DeviceID + " on server " + server + " by user " + User.Identity.Name);
                            }
                            else
                            {
                                //determine if this is an EPS or Enterprise Print Server
                                //Add Enterprise Print Queue First.
                                if (Support.GetEnterprisePrintServers().Contains(server))
                                {
                                    //Try to add the printer now the port is defined on the server.
                                    //first setup the props of the printer
                                    AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.DeviceID, DriverName = theNewPrinter.ENTDriverName, EnableBIDI = false, PortName = theNewPrinter.ENTPortName, Published = false, Shared = false,SourcePrinter=theNewPrinter.ENTSourcePrinter,SourceServer=Support.GetEntGoldPrintServer() };
                                    PrintServer thePrintServer = new PrintServer(@"\\" + theNewPrinter.PrintServer, PrintSystemDesiredAccess.AdministrateServer);
                                    //Use the string return function to determine if the printer was successfully added or not.
                                    logger.Debug("Before creation of new ENT Print Queue : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    outcome.Add(AddNewEnterprisePrinterStringReturn(newPrinter, mySession, thePrintServer));
                                    logger.Debug("After creation of new ENT Print Queue : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    logger.Info("New Printer added.  Printer Name : " + theNewPrinter.DeviceID + " on server " + theNewPrinter.PrintServer + " by user " + User.Identity.Name);
                                    //Clone the device?
                                    if (Support.UseEntGoldPrinter() && theNewPrinter.ENTSourcePrinter != null)
                                    {
                                        var clonePrinterSettings = false;
                                        if (Support.cloneEntPrinterDeviceSettings() && theNewPrinter.cloneENTDevSettings)
                                        {
                                            clonePrinterSettings = true;
                                        }
                                        outcome.Add(clonePrintQueue(newPrinter, server, Support.GetEntGoldPrintServer(), clonePrinterSettings));
                                        logger.Debug("After Print Queue is Cloned : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    }

                                    //if (Support.UseEnterprisePrintTrays())
                                    //{
                                    //    outcome.Add(SetPrinterTray(theNewPrinter.DeviceID, theNewPrinter.PrintServer, theNewPrinter.Tray));
                                    //    logger.Info("Print Tray Set for Printer Name : " + theNewPrinter.DeviceID + " Tray Info: " + theNewPrinter.Tray + " on server " + theNewPrinter.PrintServer + " by user " + User.Identity.Name + " time to install" + newwatch.ElapsedMilliseconds + "ms");
                                    //}
                                    thePrintServer.Dispose();
                                }
                                //Add EPS Print Queue here.
                                else
                                {
                                    //Try to add the printer now the port is defined on the server.
                                    //first setup the props of the printer
                                    AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.DeviceID, DriverName = theNewPrinter.DriverName, EnableBIDI = false, PortName = theNewPrinter.PortName, Published = false, Shared = false, SourcePrinter = theNewPrinter.SourcePrinter };
                                    //Use the string return function to determine if the printer was successfully added or not.
                                    logger.Debug("Before creation of new EPS Print Queue : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    outcome.Add(AddNewPrinterStringReturn(newPrinter, server));
                                    logger.Debug("After creation of new EPS Print Queue : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    logger.Info("New Printer added.  Printer Name : " + theNewPrinter.DeviceID + " on server " + server + " by user " + User.Identity.Name);
                                    if (Support.UseEPSGoldPrinter() && theNewPrinter.SourcePrinter != null)
                                    {
                                        var clonePrinterSettings = false;
                                        if (Support.clonePrinterDeviceSettings() && theNewPrinter.cloneDevSettings)
                                        {
                                            clonePrinterSettings = true;
                                        }
                                        outcome.Add(clonePrintQueue(newPrinter, server, Support.GetEPSGoldPrintServer(), clonePrinterSettings));
                                        logger.Debug("After Print Queue is Cloned : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    }
                                    //if (Support.UsePrintTrays())
                                    //{
                                    //    outcome.Add(SetPrinterTray(theNewPrinter.DeviceID, server, theNewPrinter.Tray));
                                    //    logger.Info("Print Tray Set for Printer Name : " + theNewPrinter.DeviceID + " Tray Info: " + theNewPrinter.Tray + " on server " + server + " by user " + User.Identity.Name + " time to install" + newwatch.ElapsedMilliseconds + "ms");
                                    //}

                                }
                            }
                        }
                        //Cannot connect to the server, so send a message back to the user about it.
                        else
                        {
                            outcome.Add(server + " is not an active or a valid server.  Please verify the server is up or configured correctly in the web.config file.");
                            logger.Info("Current Print Server is not active or invalid.  Print Server: " + server + " by user " + User.Identity.Name);
                        }
                        mySession.Dispose();
                        //End the parallel processing.
                    });
                    newwatch.Stop();
                    //Email users from Web.Config to confirm everything went well!
                    Support.SendEmail("New Printer Added to EPS and Enterprise Print Server", string.Join(Environment.NewLine, outcome) + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to install." + Environment.NewLine + "Created by user: " + User.Identity.Name);
                    //Send a success message the Success View.
                    TempData["SuccessMessage"] = "Congrats, the printer installed correctly!  Enjoy your day.";
                    //return JSON results to the AJAX request of the view.
                    outcome.Add(@"<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + @" ms to install.</h5>");
                    return Json(outcome, JsonRequestBehavior.AllowGet);
                }
                newwatch.Stop();
                //Send email and return results if DNS does not exist for the printer.
                Support.SendEmail("Failed EPS and Enterprise Print Server Install", "Printer: " + theNewPrinter.DeviceID + " failed the DNS lookup or IP Address validation." + Environment.NewLine + "By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist or it is an invalid IP Address.  Please try again.";
                outcome.Add("Hostname of the Printer does not exist or it is an invalid IP Address.  Please try again.");
                outcome.Add("<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to attempt to install." + @"</h5>");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            newwatch.Stop();
            outcome.Add("`r`n" + newwatch.ElapsedMilliseconds + " ms to install.");
            //Return error message that 
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            outcome.Add("Something went wrong with the Model.  Please try again.");
            logger.Error("Failed to initialize the model");
            return Json(outcome, JsonRequestBehavior.AllowGet);

        }

        //Used to get options for a specific printer on a server.
        //The view currently allows to edit a print driver and clear a print queue.
        public ActionResult Details(string printer, string printserver)
        {
            //used to determine if the print server in question is an EPS server.
            //Only limited functionality for non EPS servers is currently defined.
            ViewData["thePrinter"] = printer;
            ViewData["theServer"] = printserver;
            ViewData["isEPSServer"] = Support.GetEPSServers().Exists(s => s == printserver).ToString();
            ViewData["allowADGroupToEditEPS"] = (Support.IsUserAuthorized(Support.ADGroupCanEditEPSPrinter()));
            ViewData["isEntServer"] = Support.GetEnterprisePrintServers().Exists(s => s == printserver).ToString();
            ViewData["allowDelete"] = Support.AllowEPSPrintDeletion();
            ViewData["allowADGroupToDelete"] = (Support.IsUserAuthorized(Support.ADGroupCanDeleteEPSPrinter()));
            ViewData["allowEntPrinterEdit"] = Support.EditEnterprisePrinters();
            ViewData["allowPrintJob"] = (Support.IsUserAuthorized(Support.ADGroupCanPurgePrintQueues()));
            //Return the printer for the View.
            if (Support.GetEPSServers().Exists(s => s == printserver) == true)
            {
                return View(GetPrinterOnAllEPSServers(printer));
            }
            else if (Support.GetEnterprisePrintServers().Exists(s => s == printserver) == true)
            {
                return View(GetPrinterOnAllPrintServers(printer));
            }

            return View(GetPrinterOnAllEPSServers(printer));
        }
        //Used to purge print queues from all EPS Servers.
        public JsonResult PurgePrintQueueAllServers([Bind(Include = "Name")]Printer theNewPrinter)
        {
            List<string> outcome = new List<string>();

            if (!Support.IsUserAuthorized(Support.ADGroupCanPurgePrintQueues()))
            {
                outcome.Add("You are not authorized to purge print jobs.");
                logger.Info("User " + User.Identity.Name + " attempted to purge print queue for: " + theNewPrinter.Name);
                Support.SendEmail("Print Queue Failed to Purged", string.Join(Environment.NewLine, outcome) + Environment.NewLine + "Purged attempted by user: " + User.Identity.Name);
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }

            //if (Support.AdditionalSecurity() == true)
            //{
            //    var theADGroup = Support.ADGroupCanPurgePrintQueues();
            //    var testing = User.Identity;
            //    bool isInRole = User.IsInRole(theADGroup);
            //    if (isInRole == false)
            //    {
            //        outcome.Add("You are not authorized to purge print jobs.");
            //        logger.Info("User " + User.Identity.Name + " attempted to purge print queue for: " + theNewPrinter.Name);
            //        Support.SendEmail("Print Queue Failed to Purged", string.Join(Environment.NewLine, outcome) + Environment.NewLine + "Purged attempted by user: " + User.Identity.Name);
            //        return Json(outcome, JsonRequestBehavior.AllowGet);
            //    }
            //}

            Parallel.ForEach(Support.GetEPSServers(), server =>
            {

                try
                {
                    ClearPrintQueue(theNewPrinter.Name, server);
                    outcome.Add("Successfully purged " + theNewPrinter.Name + " on server " + server);
                    logger.Info("Purged all jobs on " + theNewPrinter.Name + " on server " + server + " by user " + User.Identity.Name);
                }
                catch
                {
                    outcome.Add("Failed to purge " + theNewPrinter.Name + " on server " + server);
                    logger.Info("Failed to purge jobs on " + theNewPrinter.Name + " on server " + server + " by user " + User.Identity.Name);
                }
            });

            Support.SendEmail("Print Queue Purged", string.Join(Environment.NewLine, outcome) + Environment.NewLine + "Purged by user: " + User.Identity.Name);
            return Json(outcome);
        }
        //Used to get options for a specific printer on a server.
        //The view currently allows to edit a print driver and clear a print queue.
        public ActionResult Options(string printer, string printServer)
        {
            //used to determine if the print server in question is an EPS server.
            //Only limited functionality for non EPS servers is currently defined.
            Session["IsEPSServer"] = Support.GetEPSServers().Exists(s => s == printServer).ToString();
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
                    Support.SendEmail("Print Queue Cleared", "Printer: " + theNewPrinter.Name + " on server: " + theNewPrinter.PrintServer + " has been cleared successfully. by user: " + User.Identity.Name);
                    TempData["SuccessMessage"] = "Congrats, the print queue has been cleared!  Enjoy your day.";
                    return RedirectToAction("Success");
                }
                else
                {
                    //something went wrong and it couldn't clear the queue.
                    Support.SendEmail("Print Queue failed to clear", "Printer: " + theNewPrinter.Name + " on server: " + theNewPrinter.PrintServer + " has failed to clear. by user: " + User.Identity.Name);
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
            var myPrinter = GetPrinter(Support.GetEPSServers().First(), printer);
            //Return a list of available print drivers from web.config
            ViewData["printDrivers"] = Support.GetAllPrintDrivers();
            //Determines if an PortName field should be returned to users.
            ViewData["useIP"] = Support.UsePrinterIPAddr();
            ViewData["useTrays"] = Support.UsePrintTrays();
            ViewData["getTrays"] = Support.GetTrays();

            //return view with printer info.
            return View(myPrinter);
        }
        //Will be used at some point to edit non EPS printers, potentially...
        public ActionResult EditEnterprisePrinter(string printer, string printServer)
        {
            var myPrinter = GetPrinter(printServer, printer);
            //Catch to see if the get printer function returns a result.  If not return to index.
            if (myPrinter == null)
            {
                return RedirectToAction("Index");
            }
            ViewData["printDrivers"] = Support.GetAllEnterprisePrintDrivers();
            ViewData["useIP"] = Support.UsePrinterIPAddr();
            ViewData["useTrays"] = Support.UsePrintTrays();
            ViewData["getTrays"] = Support.GetTrays();
            return View(myPrinter);
        }
        //AJAX Json response for editing an EPS Printer.
        //Currently deletes and readds the printer
        public JsonResult EditPrinterJSON([Bind(Include = "Name,Driver,PortName,Tray")]Printer theNewPrinter)
        {
            //Do some timing on the whole process.
            Stopwatch newwatch = new Stopwatch();
            newwatch.Start();

            //Initialize string for the output.
            List<string> outcome = new List<string>();

            //Used to determine if user has correct access to edit a EPS printer.
            //If they don't then it returns an error message for them.
            if (!Support.IsUserAuthorized(Support.ADGroupCanEditEPSPrinter()))
            {
                outcome.Add("You are not authorized to edit printers.");
                logger.Info("User " + User.Identity.Name + " attempted to edit printer: " + theNewPrinter.Name);
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }

            if (ModelState.IsValid)
            {
                //Verify the printer name is still correct.
                if (!Support.ValidatePrinterDNS() || Support.ValidHostname(theNewPrinter.PortName) == true)
                {
                    //Start parallel processing on each EPS server.
                    Parallel.ForEach(Support.GetEPSServers(), (server) =>
                    {
                        //Verify a Cim Session can be created before calling methods that use it... Maybe a bit backwards.
                        CimSession mySession = CimSession.Create(server);
                        var testtheConnection = mySession.TestConnection();
                        if (testtheConnection == true)
                        {
                            //Continue on if a session can be created
                            if (ExistingPrinterPort(theNewPrinter.PortName, mySession) == false)
                            {
                                //Check to see if port doesn't exist and create it if it does not exist.
                                AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.PortName, HostAddress = theNewPrinter.PortName, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                                AddNewPrinterPort(AddThePort, mySession);
                                logger.Info("New Port added.  Port Name : " + theNewPrinter.PortName + " on server " + server + " by user " + User.Identity.Name);
                            }
                            else
                            {
                                //don't do anything if port is actually there.
                            }

                            //PrintServer class requires the 2 wacks in the server name.
                            PrintServer printServer = new PrintServer(@"\\" + server, PrintSystemDesiredAccess.AdministrateServer);
                            //Need to delete out the old printer first, since I haven't found a good way to change print drivers/properities.
                            outcome.Add(DeletePrinter(theNewPrinter.Name, printServer));
                            //create the new printer object with the correct props.
                            AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.Name, DriverName = theNewPrinter.Driver, EnableBIDI = false, PortName = theNewPrinter.PortName, Published = false, Shared = false };
                            //Try and add the printer back in after it's been deleted.
                            outcome.Add(AddNewPrinterStringReturn(newPrinter, mySession, printServer));
                            //outcome.Add(clonePrintQueue(newPrinter, printServer.Name, "localhost"));
                            logger.Info("New Printer added.  Printer Name : " + theNewPrinter.Name + " on server " + server + " by user " + User.Identity.Name);
                            if (Support.UsePrintTrays())
                            {
                                outcome.Add(SetPrinterTray(theNewPrinter.Name, server, theNewPrinter.Tray));
                                logger.Info("Print Tray Set for Printer Name : " + theNewPrinter.Name + " Tray Info: " + theNewPrinter.Tray + " on server " + server + " by user " + User.Identity.Name + " time to install" + newwatch.ElapsedMilliseconds + "ms");
                            }

                        }
                        else
                        {
                            //Print server doesn't exist or is done.  Need to check web.config currently.
                            outcome.Add(server + "Is not a valid server.  Please contact the creator of this thing and have them check the web.config for valid EPS servers.");
                        }
                        mySession.Dispose();
                    });
                    newwatch.Stop();
                    //Finish the Parallel loop and return the results.
                    Support.SendEmail("EPS Printer Edited", string.Join(Environment.NewLine, outcome) + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to delete and install." + Environment.NewLine + "Created by user: " + User.Identity.Name);
                    TempData["SuccessMessage"] = "Congrats, the printer updated correctly!  Enjoy your day.";
                    outcome.Add(@"<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + @" ms to delete and install.</h5>");
                    return Json(outcome, JsonRequestBehavior.AllowGet);

                }
                newwatch.Stop();
                Support.SendEmail("Failed EPS Edit", "Printer: " + theNewPrinter.Name + " failed the DNS lookup or IP address validation." + Environment.NewLine + "By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist or it is not a valid IP address.  Please try again.";
                outcome.Add("Hostname of the Printer does not exist or it is not a valid IP address.  Please try again.");
                outcome.Add("<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to attempt to install." + @"</h5>");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            logger.Error("Failed to initialize the model");
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            outcome.Add("Something went wrong with the Model.  Please try again.");
            return Json(outcome, JsonRequestBehavior.AllowGet);

        }
        //AJAX Json response for editing an EPS Printer.
        //Currently deletes and readds the printer
        public JsonResult EditEnterprisePrinterJSON([Bind(Include = "Name,Driver,PortName,Tray,PrintServer")]Printer theNewPrinter)
        {
            //Do some timing on the whole process.
            Stopwatch newwatch = new Stopwatch();
            newwatch.Start();

            //Initialize string for the output.
            List<string> outcome = new List<string>();

            //Used to determine if user has correct access to edit an Enterprise printer.
            //If they don't then it returns an error message for them.
            if (!Support.IsUserAuthorized(Support.ADGroupCanEditEnterprisePrinter()))
            {
                outcome.Add("You are not authorized to edit Enterprise printers.");
                logger.Info("User " + User.Identity.Name + " attempted to edit printer: " + theNewPrinter.Name);
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }

            //Catch case where a print server may not be submitted
            if (theNewPrinter.PrintServer == null)
            {
                outcome.Add("Print Server is null.  Please try again.");
                logger.Info("User " + User.Identity.Name + " attempted to edit printer: " + theNewPrinter.Name + " but print server is null");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            string server = theNewPrinter.PrintServer.ToString();

            if (ModelState.IsValid)
            {
                //Verify the printer name is still correct.
                if (!Support.ValidatePrinterDNS() || Support.ValidHostname(theNewPrinter.PortName) == true)
                {
                    //Verify a Cim Session can be created before calling methods that use it... Maybe a bit backwards.
                    CimSession mySession = CimSession.Create(server);
                    var testtheConnection = mySession.TestConnection();
                    if (testtheConnection == true)
                    {
                        //Continue on if a session can be created
                        if (ExistingPrinterPort(theNewPrinter.PortName, mySession) == false)
                        {
                            //Check to see if port doesn't exist and create it if it does not exist.
                            AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.PortName, HostAddress = theNewPrinter.PortName, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                            AddNewPrinterPort(AddThePort, mySession);
                            logger.Info("New Port added.  Port Name : " + theNewPrinter.PortName + " on server " + theNewPrinter.PrintServer + " by user " + User.Identity.Name);
                        }
                        else
                        {
                            //don't do anything if port is actually there.
                        }

                        //PrintServer class requires the 2 wacks in the server name.
                        PrintServer printServer = new PrintServer(@"\\" + server, PrintSystemDesiredAccess.AdministrateServer);
                        //Need to delete out the old printer first, since I haven't found a good way to change print drivers/properities.
                        outcome.Add(DeletePrinter(theNewPrinter.Name, printServer));
                        //create the new printer object with the correct props.
                        AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.Name, DriverName = theNewPrinter.Driver, EnableBIDI = false, PortName = theNewPrinter.PortName, Published = false, Shared = false };
                        //Try and add the printer back in after it's been deleted.
                        outcome.Add(AddNewEnterprisePrinterStringReturn(newPrinter, mySession, printServer));
                        logger.Info("New Printer added.  Printer Name : " + theNewPrinter.Name + " on server " + server + " by user " + User.Identity.Name);
                        if (Support.UsePrintTrays())
                        {
                            outcome.Add(SetPrinterTray(theNewPrinter.Name, server, theNewPrinter.Tray));
                            logger.Info("Print Tray Set for Printer Name : " + theNewPrinter.Name + " Tray Info: " + theNewPrinter.Tray + " on server " + server + " by user " + User.Identity.Name + " time to install" + newwatch.ElapsedMilliseconds + "ms");
                        }

                    }
                    else
                    {
                        //Print server doesn't exist or is done.  Need to check web.config currently.
                        outcome.Add(server + "Is not a valid server.  Please contact the creator of this thing and have them check the web.config for valid EPS servers.");
                    }
                    mySession.Dispose();
                    newwatch.Stop();
                    //Finish the Parallel loop and return the results.
                    Support.SendEnterpriseEmail("EPS Printer Edited", string.Join(Environment.NewLine, outcome) + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to delete and install." + Environment.NewLine + "Created by user: " + User.Identity.Name);
                    TempData["SuccessMessage"] = "Congrats, the printer updated correctly!  Enjoy your day.";
                    outcome.Add(@"<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + @" ms to delete and install.</h5>");
                    return Json(outcome, JsonRequestBehavior.AllowGet);

                }
                newwatch.Stop();
                Support.SendEnterpriseEmail("Failed EPS Edit", "Printer: " + theNewPrinter.Name + " failed the DNS lookup or IP address validation." + Environment.NewLine + "By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist or it is not a valid IP address.  Please try again.";
                outcome.Add("Hostname of the Printer does not exist or it is not a valid IP address.  Please try again.");
                outcome.Add("<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to attempt to install." + @"</h5>");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            logger.Error("Failed to initialize the model");
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            outcome.Add("Something went wrong with the Model.  Please try again.");
            return Json(outcome, JsonRequestBehavior.AllowGet);

        }
        //Method to return all printers from a specified print server
        //AJAX Json response for editing an EPS Printer.
        //Currently deletes and readds the printer
        public JsonResult DeleteEPSPrinter([Bind(Include = "Name")]Printer theNewPrinter)
        {
            //Do some timing on the whole process.
            Stopwatch newwatch = new Stopwatch();
            newwatch.Start();

            //Initialize string for the output.
            List<string> outcome = new List<string>();

            //Used to determine if user has correct access to Delete an EPS printer.
            //If they don't then it returns an error message for them.
            if (!Support.IsUserAuthorized(Support.ADGroupCanDeleteEPSPrinter()))
            {
                outcome.Add("You are not authorized to delete printers.");
                logger.Info("User " + User.Identity.Name + " attempted to delete EPS printer: " + theNewPrinter.Name);
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }

            if (ModelState.IsValid)
            {
                //Start parallel processing on each EPS server.
                Parallel.ForEach(Support.GetEPSServers(), (server) =>
                {
                    //Verify a Cim Session can be created before calling methods that use it... Maybe a bit backwards.
                    CimSession mySession = CimSession.Create(server);
                    var testtheConnection = mySession.TestConnection();
                    if (testtheConnection == true)
                    {
                        //Continue on if a session can be created
                        //PrintServer class requires the 2 wacks in the server name.
                        PrintServer printServer = new PrintServer(@"\\" + server, PrintSystemDesiredAccess.AdministrateServer);
                        //Need to delete out the old printer first, since I haven't found a good way to change print drivers/properties.
                        outcome.Add(DeletePrinter(theNewPrinter.Name, printServer));
                        //logger.Info("Printer deleted.  Printer Name : " + theNewPrinter.Name + " on server " + server + " by user " + User.Identity.Name);
                        printServer.Dispose();
                    }
                    else
                    {
                        //Print server doesn't exist or is done.  Need to check web.config currently.
                        outcome.Add(server + "Is not a valid server.  Please contact the creator of this thing and have them check the web.config for valid EPS servers.");
                    }
                    mySession.Dispose();
                });
                newwatch.Stop();
                logger.Info(string.Join(Environment.NewLine, outcome) + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to delete." + Environment.NewLine + "Deleted by user: " + User.Identity.Name);
                //Finish the Parallel loop and return the results.
                Support.SendEmail("EPS Printer Deleted", string.Join(Environment.NewLine, outcome) + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to delete." + Environment.NewLine + "Deleted by user: " + User.Identity.Name);
                TempData["SuccessMessage"] = "Congrats, the printer was deleted!  Enjoy your day.";
                outcome.Add(@"<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + @" ms to delete.</h5>");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            logger.Error("Failed to initialize the model");
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            outcome.Add("Something went wrong with the Model.  Please try again.");
            return Json(outcome, JsonRequestBehavior.AllowGet);

        }
        public JsonResult DeleteEnterprisePrinter([Bind(Include = "Name,PrintServer")]Printer theNewPrinter)
        {
            //Do some timing on the whole process.
            Stopwatch newwatch = new Stopwatch();
            newwatch.Start();

            //Initialize string for the output.
            List<string> outcome = new List<string>();

            //Used to determine if user has correct access to Delete an Enterprise printer.
            //If they don't then it returns an error message for them.
            if (!Support.IsUserAuthorized(Support.ADGroupCanDeleteEnterprisePrinter()))
            {
                outcome.Add("You are not authorized to delete Enterprise printers.");
                logger.Info("User " + User.Identity.Name + " attempted to delete Enterprise printer: " + theNewPrinter.Name);
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }

            //Check to verify the print server field is not null.
            if (theNewPrinter.PrintServer == null)
            {
                outcome.Add("Print Server is null.  Please try again.");
                logger.Info("User " + User.Identity.Name + " attempted to delete printer: " + theNewPrinter.Name + " but print server is null");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }

            string server = theNewPrinter.PrintServer.Trim().ToString();

            if (ModelState.IsValid)
            {
                //Verify a Cim Session can be created before calling methods that use it... Maybe a bit backwards.
                CimSession mySession = CimSession.Create(server);
                var testtheConnection = mySession.TestConnection();
                if (testtheConnection == true)
                {
                    //Continue on if a session can be created
                    //PrintServer class requires the 2 wacks in the server name.
                    PrintServer printServer = new PrintServer(@"\\" + server, PrintSystemDesiredAccess.AdministrateServer);
                    //Need to delete out the old printer first, since I haven't found a good way to change print drivers/properties.
                    outcome.Add(DeletePrinter(theNewPrinter.Name, printServer));
                    printServer.Dispose();
                    //logger.Info("Printer deleted.  Printer Name : " + theNewPrinter.Name + " on server " + server + " by user " + User.Identity.Name);
                }
                else
                {
                    //Print server doesn't exist or is done.  Need to check web.config currently.
                    outcome.Add(server + "Is not a valid server.  Please contact the creator of this thing and have them check the web.config for valid EPS servers.");
                }
                mySession.Dispose();
                newwatch.Stop();
                logger.Info(string.Join(Environment.NewLine, outcome) + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to delete." + Environment.NewLine + "Deleted by user: " + User.Identity.Name);
                //Finish the Parallel loop and return the results.
                Support.SendEmail("EPS Printer Deleted", string.Join(Environment.NewLine, outcome) + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to delete." + Environment.NewLine + "Deleted by user: " + User.Identity.Name);
                TempData["SuccessMessage"] = "Congrats, the printer was deleted!  Enjoy your day.";
                outcome.Add(@"<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + @" ms to delete.</h5>");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            logger.Error("Failed to initialize the model");
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            outcome.Add("Something went wrong with the Model.  Please try again.");
            return Json(outcome, JsonRequestBehavior.AllowGet);

        }
        public List<Printer> GetPrinters(string server)
        {
            //Currently using the PrintServer class instead of a WMI query.
            //Was roughly 3-4 times faster to use PrintServer instead of WMI on 1400 printers.  Still takes about 10 seconds though.

            if (server == "")
            {
                List<Printer> printerList = new List<Printer>();
                return (printerList);
            }

            try
            {
                //PrintServer class requires the 2 wacks in the server name.
                PrintServer printServer = new PrintServer(@"\\" + server);
                //Make a pretty ordered list by name.
                var myPrintQueues = printServer.GetPrintQueues(new EnumeratedPrintQueueTypes[] { EnumeratedPrintQueueTypes.Local }).OrderBy(t => t.Name);
                //Create an empty list to return.
                List<Printer> printerList = new List<Printer>();

                if (Support.ShowNumberPrintJobs())
                {
                    //Loop through and add all items to the custom class.
                    foreach (PrintQueue pq in myPrintQueues)
                    {
                        //pq.Refresh();
                        printerList.Add(new Printer { Name = pq.Name, Driver = pq.QueueDriver.Name, PrintServer = pq.HostingPrintServer.Name.TrimStart('\\'), PortName = pq.QueuePort.Name, NumberJobs = pq.NumberOfJobs });  //Removed the following for performance NumberJobs = pq.NumberOfJobs,
                    }

                }
                else
                {
                    //Loop through and add all items to the custom class.
                    foreach (PrintQueue pq in myPrintQueues)
                    {
                        //pq.Refresh();
                        printerList.Add(new Printer { Name = pq.Name, Driver = pq.QueueDriver.Name, PrintServer = pq.HostingPrintServer.Name.TrimStart('\\'), PortName = pq.QueuePort.Name });  //Removed the following for performance NumberJobs = pq.NumberOfJobs,
                    }
                }
                printServer.Dispose();
                //return the printers added to the custom class.
                return (printerList);
            }
            catch (Exception ex)
            {
                logger.Fatal("Failed to get printers for server: " + server, ex);
                List<Printer> printerList = new List<Printer>();
                return (printerList);
            }
        }
        //Used to purge print queues from all EPS Servers.
        public JsonResult PrintTestPages([Bind(Include = "Name")]Printer thePrinter)
        {
            List<string> outcome = new List<string>();

            Parallel.ForEach(Support.GetEPSServers(), server =>
            {

                try
                {
                    var theOutcome = PrintTestPage(server, thePrinter.Name);
                    if (theOutcome == true)
                    {
                        outcome.Add("Successfully printed a test page for " + thePrinter.Name + " on server " + server);
                    }
                    else
                    {
                        outcome.Add("Failed to print a test page for " + thePrinter.Name + " on server " + server);
                    }
                }
                catch
                {
                    outcome.Add("Failed to print a test page for " + thePrinter.Name + " on server " + server);
                }
            });
            return Json(outcome);
        }
        //Get details for a specific printer on a specific print server.
        static public Printer GetPrinter(string server, string printer)
        {
            Printer printerList = new Printer();
            try
            {
                //Currently using the PrintServer class instead of a WMI query.
                //Was roughly 3-4 times faster to use PrintServer instead of WMI on 1400 printers.  Still takes about 10 seconds though.
                //PrintServer class requires the 2 wacks in the server name.
                PrintServer printServer = new PrintServer(@"\\" + server);
                //Get the one print queue from the print server.
                var myPrintQueues = printServer.GetPrintQueue(printer);
                //refresh and add the print queue to the custom class.
                myPrintQueues.Refresh();
                printerList = new Printer { Name = myPrintQueues.Name, Driver = myPrintQueues.QueueDriver.Name, PrintServer = myPrintQueues.HostingPrintServer.Name, NumberJobs = myPrintQueues.NumberOfJobs, PortName = myPrintQueues.QueuePort.Name };
                myPrintQueues.Dispose();
                printServer.Dispose();

            }
            catch
            {
                printerList = null;
            }
            return (printerList);
        }
        //Get details for a specific printer.
        //Used to find printer on all EPS servers at once.
        static public List<Printer> GetPrinterOnAllEPSServers(string printer)
        {
            //Initialize list to display to end users if it completes or not.
            List<Printer> outcome = new List<Printer>();

            //kick off multiple threads to install printers quickly
            Parallel.ForEach(Support.GetEPSServers(), (server) =>
            {
                try
                {
                    //Currently using the PrintServer class instead of a WMI query.
                    //Was roughly 3-4 times faster to use PrintServer instead of WMI on 1400 printers.  Still takes about 10 seconds though.
                    //PrintServer class requires the 2 wacks in the server name.
                    PrintServer printServer = new PrintServer(@"\\" + server);
                    //Get the one print queue from the print server.
                    var myPrintQueues = printServer.GetPrintQueue(printer);
                    myPrintQueues.Refresh();
                    string theTray;
                    //var theTray = GetCurrentPrintTray(printer, server);
                    if (myPrintQueues.QueueDriver.Name.Contains("ZDesigner") || !Support.UsePrintTrays())
                    {
                        theTray = null;
                    }
                    else
                    {
                        theTray = GetCurrentPrintTray(myPrintQueues);

                    }
                    outcome.Add(new Printer { Name = myPrintQueues.Name.ToUpper(), Driver = myPrintQueues.QueueDriver.Name, PrintServer = server, NumberJobs = myPrintQueues.NumberOfJobs, PortName = myPrintQueues.QueuePort.Name, Status = GetPrinterStatus(myPrintQueues), Tray = theTray });
                    printServer.Dispose();
                    myPrintQueues.Dispose();
                }
                catch
                {
                    outcome.Add(new Printer { Name = printer, PrintServer = server, Status = "Not Installed or Server Down" });
                }
            });
            return (outcome.OrderBy(s => s.PrintServer).ToList());
        }
        //Get details for a specific printer on a specific Print Server.
        //Used for Enterprise Print Servers.
        static public List<Printer> GetPrinterOnPrintServer(string printer, string printserver)
        {
            //Initialize list to display to end users if it completes or not.
            List<Printer> outcome = new List<Printer>();

            try
            {
                //Currently using the PrintServer class instead of a WMI query.
                //Was roughly 3-4 times faster to use PrintServer instead of WMI on 1400 printers.  Still takes about 10 seconds though.
                //PrintServer class requires the 2 wacks in the server name.
                PrintServer printServer = new PrintServer(@"\\" + printserver);
                //Get the one print queue from the print server.
                var myPrintQueues = printServer.GetPrintQueue(printer);
                myPrintQueues.Refresh();
                string theTray;
                //var theTray = GetCurrentPrintTray(printer, server);
                if (myPrintQueues.QueueDriver.Name.Contains("ZDesigner") || !Support.UsePrintTrays())
                {
                    theTray = null;
                }
                else
                {
                    theTray = GetCurrentPrintTray(myPrintQueues);

                }
                outcome.Add(new Printer { Name = myPrintQueues.Name.ToUpper(), Driver = myPrintQueues.QueueDriver.Name, PrintServer = printserver, NumberJobs = myPrintQueues.NumberOfJobs, PortName = myPrintQueues.QueuePort.Name, Status = GetPrinterStatus(myPrintQueues), Tray = theTray });
                printServer.Dispose();
                myPrintQueues.Dispose();
            }
            catch
            {
                outcome.Add(new Printer { Name = printer, PrintServer = printserver, Status = "Not Installed or Server Down" });
            }

            return (outcome.ToList());
        }
        //Get details for a specific printer.
        //Used to find printer on all Print servers at once.
        static public List<Printer> GetPrinterOnAllPrintServers(string printer)
        {
            //Initialize list to display to end users if it completes or not.
            List<Printer> outcome = new List<Printer>();
            //kick off multiple threads to install printers quickly
            //Parallel.ForEach(Support.GetEnterprisePrintServers(), (server) =>
            //{
            //    try
            //    {
            //        //Currently using the PrintServer class instead of a WMI query.
            //        //Was roughly 3-4 times faster to use PrintServer instead of WMI on 1400 printers.  Still takes about 10 seconds though.
            //        //PrintServer class requires the 2 wacks in the server name.
            //        PrintServer printServer = new PrintServer(@"\\" + server);
            //        //Get the one print queue from the print server.
            //        var myPrintQueues = printServer.GetPrintQueue(printer);
            //        myPrintQueues.Refresh();
            //        string theTray;
            //        //var theTray = GetCurrentPrintTray(printer, server);
            //        if (myPrintQueues.QueueDriver.Name.Contains("ZDesigner") || !Support.UsePrintTrays())
            //        {
            //            theTray = null;
            //        }
            //        else
            //        {
            //            theTray = GetCurrentPrintTray(myPrintQueues);

            //        }
            //        outcome.Add(new Printer { Name = myPrintQueues.Name.ToUpper(), Driver = myPrintQueues.QueueDriver.Name, PrintServer = server, NumberJobs = myPrintQueues.NumberOfJobs, PortName = myPrintQueues.QueuePort.Name, Status = "Installed", Tray = theTray });
            //        printServer.Dispose();
            //        myPrintQueues.Dispose();
            //    }
            //    catch
            //    {
            //        outcome.Add(new Printer { Name = printer,PrintServer=server ,Status = "Not Installed or Server Down" });
            //    }
            //});

            foreach (var server in Support.GetEnterprisePrintServers())
            {
                try
                {
                    //Currently using the PrintServer class instead of a WMI query.
                    //Was roughly 3-4 times faster to use PrintServer instead of WMI on 1400 printers.  Still takes about 10 seconds though.
                    //PrintServer class requires the 2 wacks in the server name.
                    PrintServer printServer = new PrintServer(@"\\" + server);
                    //Get the one print queue from the print server.
                    var myPrintQueues = printServer.GetPrintQueue(printer);
                    myPrintQueues.Refresh();
                    string theTray;
                    //var theTray = GetCurrentPrintTray(printer, server);
                    if (myPrintQueues.QueueDriver.Name.Contains("ZDesigner") || !Support.UsePrintTrays())
                    {
                        theTray = null;
                    }
                    else
                    {
                        theTray = GetCurrentPrintTray(myPrintQueues);

                    }
                    outcome.Add(new Printer { Name = myPrintQueues.Name.ToUpper(), Driver = myPrintQueues.QueueDriver.Name, PrintServer = server, NumberJobs = myPrintQueues.NumberOfJobs, PortName = myPrintQueues.QueuePort.Name, Status = "Installed", Tray = theTray });
                    printServer.Dispose();
                    myPrintQueues.Dispose();
                }
                catch
                {
                    outcome.Add(new Printer { Name = printer, PrintServer = server, Status = "Not Installed or Server Down" });
                }
            }


            return (outcome.OrderBy(s => s.PrintServer).OrderBy(s => s.Status).ToList());
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
                //var exist = queryInstance.Where(p => p.CimInstanceProperties["Name"].Value.Equals(portName));
                var exist = queryInstance.Where(p => p.CimInstanceProperties["Name"].Value.ToString().ToUpper().Equals(portName.ToUpper()));
                //If the port exists, then return true.
                if (exist.FirstOrDefault() != null)
                {
                    //return true and exit.
                    mySession.Dispose();
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
        //Return a true/false if printer port is active.  Need to know when adding a printer.
        static private bool ExistingPrinterPort(string portName, CimSession session)
        {
            //This uses Powershell CimSession and WMI to query the information from the server.
            string Namespace = @"root\cimv2";
            string OSQuery = "SELECT * FROM Win32_TCPIPPrinterPort";
            //CimSession mySession = CimSession.Create(serverName);
            //Verify the session created correctly, otherwise it will error out if it fails to connect.
            var testtheConnection = session.TestConnection();
            if (testtheConnection == true)
            {
                //Query WMI for Printer Ports on the server.
                IEnumerable<CimInstance> queryInstance = session.QueryInstances(Namespace, "WQL", OSQuery);
                //Check to see if it exists in the query response.
                var exist = queryInstance.Where(p => p.CimInstanceProperties["Name"].Value.Equals(portName));
                //If the port exists, then return true.
                if (exist.FirstOrDefault() != null)
                {
                    //return true and exit.
                    //mySession.Dispose();
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
        //Add a printer port to a server.
        //Assumes you have verified the print server and ip/dns entries are valid.
        //This method reuses an existing CimSession.
        static private void AddNewPrinterPort(AddPrinterPortClass theNewPrinterPort, CimSession session)
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
            //CimSession Session = CimSession.Create(thePrintServer);
            //Actually create the printer port on the print server.
            CimInstance myPrinter = session.CreateInstance(Namespace, newPrinter);
            //Cleanup
            myPrinter.Dispose();
            //Session.Dispose();
        }
        //Add a printer and return success or failure in a string.
        public string AddNewPrinterStringReturn(AddPrinterClass theNewPrinter, string thePrintServer)
        {
            Stopwatch newwatch = new Stopwatch();
            newwatch.Start();
            logger.Debug("Start of Adding New Printer Function : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            PrintServer printServer = new PrintServer(@"\\" + thePrintServer, PrintSystemDesiredAccess.AdministrateServer);
            PrintPropertyDictionary printProps = new PrintPropertyDictionary { };
            // Share the new printer using Remove/Add methods
            PrintBooleanProperty shared = new PrintBooleanProperty("IsShared", false);
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
                logger.Debug("Start of Adding Installing Print Queue : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                PrintQueue AddingPrinterHere = printServer.InstallPrintQueue(theNewPrinter.DeviceID, theNewPrinter.DriverName, port, "WinPrint", printProps);
                logger.Debug("Installed Print Queue Before Commit : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                printServer.Commit();
                logger.Debug("After Commit of Print Queue : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                printServer.Dispose();
                logger.Debug("After dispose of Print Server : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            }
            catch (System.Printing.PrintSystemException e)
            {
                //printServer.Dispose();
                return (theNewPrinter.DeviceID + " failed to install on " + thePrintServer + " with error message " + e.Message);
            }
            logger.Debug("Finished Installing Printer before Sleep for 1000ms : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            //Have to change some printer properties that are not modified in installation.
            System.Threading.Thread.Sleep(1000);
            //This uses PowerShell CimSession and WMI to query the information from the server.
            //Used to change printer props
            string Namespace = @"root\cimv2";
            string OSQuery = "SELECT * FROM Win32_Printer";
            logger.Debug("Before Cimsession creation : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            CimSession mySession = CimSession.Create(thePrintServer);
            logger.Debug("After Cimsession creation and before testing connection : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            //Verify the session created correctly, otherwise it will error out if it fails to connect.
            var testtheConnection = mySession.TestConnection();
            logger.Debug("After Cimsession Connection Test : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            if (testtheConnection == true)
            {
                logger.Debug("Before Query Instance Setup : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                //Query WMI for Printer Ports on the server.
                IEnumerable<CimInstance> queryInstance = mySession.QueryInstances(Namespace, "WQL", OSQuery);
                logger.Debug("After Query Instance Setup but before Query Instance Where Statement : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                //Check to see if it exists in the query response.
                CimInstance exist = queryInstance.Where(p => p.CimInstanceProperties["Name"].Value.Equals(theNewPrinter.DeviceID)).FirstOrDefault();
                logger.Debug("After Query Instance Statement to see if Print Queue exists : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                //If the printer exists, then return true.
                if (exist != null)
                {
                    logger.Debug("Before Modifying the Properties : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                    //return true and exit.
                    exist.CimInstanceProperties["EnableBIDI"].Value = false;
                    exist.CimInstanceProperties["DoCompleteFirst"].Value = true;
                    exist.CimInstanceProperties["RawOnly"].Value = true;
                    try
                    {
                        mySession.ModifyInstance(Namespace, exist);
                        logger.Debug("After Modifying the WMI Properities and before Disposing of the Session : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                        mySession.Dispose();
                        logger.Debug("After Disposing of CIM Instance : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                        return (theNewPrinter.DeviceID + " was added successfully on " + thePrintServer);
                    }
                    catch
                    {
                        return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + thePrintServer);
                    }
                }
                else
                {
                    //item doesn't exist...
                    mySession.Dispose();
                    return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + thePrintServer);
                }
            }
            else
            {
                //session didn't connect.  Return False
                return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + thePrintServer);
            }

            //return (theNewPrinter.DeviceID + " was added successfully on " + thePrintServer);

        }
        //Add a printer and return success or failure in a string.
        //Uses existing connections to the print server.
        static private string AddNewPrinterStringReturn(AddPrinterClass theNewPrinter, CimSession session, PrintServer printServer)
        {
            //PrintServer printServer = new PrintServer(@"\\" + thePrintServer);
            PrintPropertyDictionary printProps = new PrintPropertyDictionary { };
            // Share the new printer using Remove/Add methods
            PrintBooleanProperty shared = new PrintBooleanProperty("IsShared", false);
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
                //printServer.Dispose();
            }
            catch (System.Printing.PrintSystemException e)
            {
                //printServer.Dispose();
                return (theNewPrinter.DeviceID + " failed to install on " + printServer.Name + " with error message " + e.Message);
            }

            //Have to change some printer properties that are not modified in installation.
            System.Threading.Thread.Sleep(3000);
            //This uses PowerShell CimSession and WMI to query the information from the server.
            //Used to change printer props
            string Namespace = @"root\cimv2";
            string OSQuery = "SELECT * FROM Win32_Printer";
            //CimSession mySession = CimSession.Create(thePrintServer);
            //Verify the session created correctly, otherwise it will error out if it fails to connect.
            var testtheConnection = session.TestConnection();
            if (testtheConnection == true)
            {
                //Query WMI for Printer Ports on the server.
                IEnumerable<CimInstance> queryInstance = session.QueryInstances(Namespace, "WQL", OSQuery);
                //Check to see if it exists in the query response.
                CimInstance exist = queryInstance.Where(p => p.CimInstanceProperties["Name"].Value.Equals(theNewPrinter.DeviceID)).FirstOrDefault();
                //If the port exists, then return true.
                if (exist != null)
                {
                    //return true and exit.
                    exist.CimInstanceProperties["EnableBIDI"].Value = false;
                    exist.CimInstanceProperties["DoCompleteFirst"].Value = true;
                    exist.CimInstanceProperties["RawOnly"].Value = true;
                    try
                    {
                        session.ModifyInstance(Namespace, exist);
                        //mySession.Dispose();
                        return (theNewPrinter.DeviceID + " was added successfully on " + printServer.Name);
                    }
                    catch
                    {
                        //mySession.Dispose();
                        return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + printServer.Name);
                    }
                }
                else
                {
                    //item doesn't exist...
                    //mySession.Dispose();
                    return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + printServer.Name);
                }
            }
            else
            {
                //session didn't connect.  Return False
                return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + printServer);
            }

            //return (theNewPrinter.DeviceID + " was added successfully on " + thePrintServer);

        }

        //Add a printer and return success or failure in a string.
        public string AddNewPrinterStringPowershellReturn(AddPrinterClass theNewPrinter, string thePrintServer)
        {
            Stopwatch newwatch = new Stopwatch();
            newwatch.Start();
            logger.Debug("Start of Adding New Printer Function : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);

            try
            {

            Collection<PSObject> testMailNickName = PowerShell.Create()
                    .AddCommand("Add-Printer")
                    .AddParameter("Name", theNewPrinter.DeviceID)
                    .AddParameter("DriverName", theNewPrinter.DriverName)
                    .AddParameter("PortName", theNewPrinter.PortName)
                    .AddParameter("ComputerName", thePrintServer)
                    .Invoke();
            }
            catch (Exception e)
            {
                //printServer.Dispose();
                return (theNewPrinter.DeviceID + " failed to install on " + thePrintServer + " with error message " + e.Message);
            }
            //PrintServer printServer = new PrintServer(@"\\" + thePrintServer, PrintSystemDesiredAccess.AdministrateServer);
            //PrintPropertyDictionary printProps = new PrintPropertyDictionary { };
            //// Share the new printer using Remove/Add methods
            //PrintBooleanProperty shared = new PrintBooleanProperty("IsShared", false);
            //PrintBooleanProperty BIDI = new PrintBooleanProperty("EnableBIDI", false);
            //PrintBooleanProperty isBIDI = new PrintBooleanProperty("IsBidiEnabled", false);
            //PrintBooleanProperty published = new PrintBooleanProperty("Published", false);
            //PrintBooleanProperty direct = new PrintBooleanProperty("IsDirect", false);
            //PrintBooleanProperty spoolFirst = new PrintBooleanProperty("ScheduleCompletedJobsFirst", true);
            //PrintBooleanProperty doComplete = new PrintBooleanProperty("DoCompleteFirst", true);
            //printProps.Add("IsShared", shared);
            //printProps.Add("EnableBIDI", BIDI);
            //printProps.Add("IsBidiEnabled", isBIDI);
            //printProps.Add("Published", published);
            //printProps.Add("IsDirect", direct);
            //printProps.Add("DoCompleteFirst", doComplete);
            //printProps.Add("ScheduleCompletedJobsFirst", spoolFirst);
            //String[] port = new String[] { theNewPrinter.PortName };

            //try
            //{
            //    logger.Debug("Start of Adding Installing Print Queue : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            //    PrintQueue AddingPrinterHere = printServer.InstallPrintQueue(theNewPrinter.DeviceID, theNewPrinter.DriverName, port, "WinPrint", printProps);
            //    logger.Debug("Installed Print Queue Before Commit : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            //    printServer.Commit();
            //    logger.Debug("After Commit of Print Queue : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            //    printServer.Dispose();
            //    logger.Debug("After dispose of Print Server : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            //}
            //catch (System.Printing.PrintSystemException e)
            //{
            //    //printServer.Dispose();
            //    return (theNewPrinter.DeviceID + " failed to install on " + thePrintServer + " with error message " + e.Message);
            //}
            logger.Debug("Finished Installing Printer before Sleep for 1000ms : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            //Have to change some printer properties that are not modified in installation.
            System.Threading.Thread.Sleep(1000);
            //This uses PowerShell CimSession and WMI to query the information from the server.
            //Used to change printer props
            string Namespace = @"root\cimv2";
            string OSQuery = "SELECT * FROM Win32_Printer";
            logger.Debug("Before Cimsession creation : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            CimSession mySession = CimSession.Create(thePrintServer);
            logger.Debug("After Cimsession creation and before testing connection : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            //Verify the session created correctly, otherwise it will error out if it fails to connect.
            var testtheConnection = mySession.TestConnection();
            logger.Debug("After Cimsession Connection Test : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
            if (testtheConnection == true)
            {
                logger.Debug("Before Query Instance Setup : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                //Query WMI for Printer Ports on the server.
                IEnumerable<CimInstance> queryInstance = mySession.QueryInstances(Namespace, "WQL", OSQuery);
                logger.Debug("After Query Instance Setup but before Query Instance Where Statement : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                //Check to see if it exists in the query response.
                CimInstance exist = queryInstance.Where(p => p.CimInstanceProperties["Name"].Value.Equals(theNewPrinter.DeviceID)).FirstOrDefault();
                logger.Debug("After Query Instance Statement to see if Print Queue exists : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                //If the printer exists, then return true.
                if (exist != null)
                {
                    logger.Debug("Before Modifying the Properties : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                    //return true and exit.
                    exist.CimInstanceProperties["EnableBIDI"].Value = false;
                    exist.CimInstanceProperties["DoCompleteFirst"].Value = true;
                    exist.CimInstanceProperties["RawOnly"].Value = true;
                    try
                    {
                        mySession.ModifyInstance(Namespace, exist);
                        logger.Debug("After Modifying the WMI Properities and before Disposing of the Session : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                        mySession.Dispose();
                        logger.Debug("After Disposing of CIM Instance : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + thePrintServer);
                        return (theNewPrinter.DeviceID + " was added successfully on " + thePrintServer);
                    }
                    catch
                    {
                        return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + thePrintServer);
                    }
                }
                else
                {
                    //item doesn't exist...
                    mySession.Dispose();
                    return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + thePrintServer);
                }
            }
            else
            {
                //session didn't connect.  Return False
                return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + thePrintServer);
            }

            //return (theNewPrinter.DeviceID + " was added successfully on " + thePrintServer);

        }

        //Clone a print queue from a source print queue.
        static private string clonePrintQueue(AddPrinterClass printer, string targetPrintServer, string sourcePrintServer,bool cloneDevSettings)
        {

            //PrintServer class requires the 2 wacks in the server name.
            //Define both source and target print servers to make the changes on.
            PrintServer theTargetPrintServer = new PrintServer(@"\\" + targetPrintServer, PrintSystemDesiredAccess.AdministrateServer);
            PrintServer theSourcePrinterServer = new PrintServer(@"\\" + sourcePrintServer, PrintSystemDesiredAccess.AdministrateServer);
            //LocalPrintServer theSourcePrinterServer = new LocalPrintServer(PrintSystemDesiredAccess.AdministrateServer);

            //verify we can connect to the source/target print queues.  Otherwise it just hangs here if there is an error.
            try
            {
                var test = theSourcePrinterServer.GetPrintQueue(printer.SourcePrinter);
                var testing = theTargetPrintServer.GetPrintQueue(printer.DeviceID);
            }
            catch (Exception e)
            {
                return ("Source Print Queue does not exist for cloning.  Error: " + e.Message.ToString());
            }

            //Get the source print queue printer defaults.
            PrintQueue sourcePrintQueue = theSourcePrinterServer.GetPrintQueue(printer.SourcePrinter);
            //Get the source print ticket
            PrintTicket sourcePrintTicket = sourcePrintQueue.DefaultPrintTicket;
            //get the target Queue
            PrintQueue theTargetPrintQueue = theTargetPrintServer.GetPrintQueue(printer.DeviceID);
            
            if (theTargetPrintQueue.QueueDriver.Name != sourcePrintQueue.QueueDriver.Name)
            {
                return ("Please make sure print drivers are the same for cloning.  Did not clone the Printer Defaults for " + printer.DeviceID + " on print server " + targetPrintServer);
            }


            try
            {

                using (PrintServer ps = theTargetPrintServer)
                {
                    using (PrintQueue pq = new PrintQueue(ps, printer.DeviceID, PrintSystemDesiredAccess.AdministratePrinter))
                    {
                        pq.DefaultPrintTicket = sourcePrintTicket;
                        pq.Commit();

                    }
                }

                if (cloneDevSettings == true)
                {
                    var sourceReg= RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine,sourcePrintServer).OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\"+printer.SourcePrinter+@"\PrinterDriverData");
                    var targetReg = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, targetPrintServer).OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\" + printer.DeviceID + @"\PrinterDriverData",true);
                    sourceReg.CopyTo(targetReg);
                }
                return (printer.DeviceID + " successfully cloned the print queue to " + targetPrintServer + ".");
            }
            catch
            {
                return (printer.DeviceID + " failed to clone the print queue on print server " + targetPrintServer + ".  Please try again.");
            }

        }
        static private string AddNewEnterprisePrinterStringReturn(AddPrinterClass theNewPrinter, CimSession session, PrintServer printServer)
        {
            //PrintServer printServer = new PrintServer(@"\\" + thePrintServer);
            PrintPropertyDictionary printProps = new PrintPropertyDictionary { };
            // Share the new printer using Remove/Add methods
            PrintBooleanProperty direct = new PrintBooleanProperty("IsDirect", false);
            PrintBooleanProperty spoolFirst = new PrintBooleanProperty("ScheduleCompletedJobsFirst", true);
            PrintBooleanProperty doComplete = new PrintBooleanProperty("DoCompleteFirst", true);
            if (Support.UseEnterprisePrinterBiDirectionalSupport() != true)
            {
                PrintBooleanProperty BIDI = new PrintBooleanProperty("EnableBIDI", false);
                printProps.Add("EnableBIDI", BIDI);
            }
            printProps.Add("IsDirect", direct);
            printProps.Add("DoCompleteFirst", doComplete);
            printProps.Add("ScheduleCompletedJobsFirst", spoolFirst);
            String[] port = new String[] { theNewPrinter.PortName };
            PrintQueueAttributes thePrintAttrs;
            if (Support.UseEnterprisePrinterBiDirectionalSupport() != true)
            {
                thePrintAttrs = PrintQueueAttributes.Shared | PrintQueueAttributes.Published | PrintQueueAttributes.ScheduleCompletedJobsFirst;
            }
            else
            {
                thePrintAttrs = PrintQueueAttributes.EnableBidi | PrintQueueAttributes.Shared | PrintQueueAttributes.Published | PrintQueueAttributes.ScheduleCompletedJobsFirst;
            }
            //thePrintAttrs = PrintQueueAttributes.EnableBidi | PrintQueueAttributes.Shared | PrintQueueAttributes.Published | PrintQueueAttributes.ScheduleCompletedJobsFirst;

            try
            {
                //PrintQueue AddingPrinterHere = printServer.InstallPrintQueue(theNewPrinter.DeviceID, theNewPrinter.DriverName, port, "WinPrint", printProps);
                PrintQueue TheNewPrintQueue = printServer.InstallPrintQueue(theNewPrinter.DeviceID, theNewPrinter.DriverName, port, "WinPrint", thePrintAttrs, theNewPrinter.DeviceID, theNewPrinter.Comment, theNewPrinter.Location, "", 1, 1);
                printServer.Commit();
                //printServer.Dispose();
            }
            catch (System.Printing.PrintSystemException e)
            {
                //printServer.Dispose();
                return (theNewPrinter.DeviceID + " failed to install on " + printServer.Name.TrimStart('\\') + " with error message " + e.Message);
            }

            //Have to change some printer properties that are not modified in installation.
            System.Threading.Thread.Sleep(1000);
            //This uses PowerShell CimSession and WMI to query the information from the server.
            //Used to change printer props
            string Namespace = @"root\cimv2";
            string OSQuery = "SELECT * FROM Win32_Printer";
            //CimSession mySession = CimSession.Create(thePrintServer);
            //Verify the session created correctly, otherwise it will error out if it fails to connect.
            var testtheConnection = session.TestConnection();
            if (testtheConnection == true)
            {
                //Query WMI for Printer Ports on the server.
                IEnumerable<CimInstance> queryInstance = session.QueryInstances(Namespace, "WQL", OSQuery);
                //Check to see if it exists in the query response.
                CimInstance exist = queryInstance.Where(p => p.CimInstanceProperties["Name"].Value.Equals(theNewPrinter.DeviceID)).FirstOrDefault();
                //If the port exists, then return true.
                if (exist != null)
                {
                    //return true and exit.
                    if (Support.UseEnterprisePrinterBiDirectionalSupport() != true)
                    {
                        exist.CimInstanceProperties["EnableBIDI"].Value = false;
                    }
                    exist.CimInstanceProperties["DoCompleteFirst"].Value = true;
                    exist.CimInstanceProperties["Published"].Value = true;
                    try
                    {
                        session.ModifyInstance(Namespace, exist);
                        //mySession.Dispose();
                        return (theNewPrinter.DeviceID + " was added successfully on " + printServer.Name.TrimStart('\\'));
                    }
                    catch
                    {
                        //mySession.Dispose();
                        return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + printServer.Name.TrimStart('\\'));
                    }
                }
                else
                {
                    //item doesn't exist...
                    //mySession.Dispose();
                    return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + printServer.Name.TrimStart('\\'));
                }
            }
            else
            {
                //session didn't connect.  Return False
                return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + printServer.Name.TrimStart('\\'));
            }

            //return (theNewPrinter.DeviceID + " was added successfully on " + thePrintServer);

        }
        //Add a printer and return success or failure in a string.
        //Uses existing connections to the print server.
        static private string ClonePrinter(AddPrinterClass theNewPrinter, CimSession session, PrintServer printServer)
        {
            //PrintServer printServer = new PrintServer(@"\\" + thePrintServer);
            PrintPropertyDictionary printProps = new PrintPropertyDictionary { };
            // Share the new printer using Remove/Add methods
            PrintBooleanProperty shared = new PrintBooleanProperty("IsShared", false);
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
                //printServer.Dispose();
            }
            catch (System.Printing.PrintSystemException e)
            {
                //printServer.Dispose();
                return (theNewPrinter.DeviceID + " failed to install on " + printServer.Name + " with error message " + e.Message);
            }

            //Have to change some printer properties that are not modified in installation.
            System.Threading.Thread.Sleep(3000);
            //This uses PowerShell CimSession and WMI to query the information from the server.
            //Used to change printer props
            string Namespace = @"root\cimv2";
            string OSQuery = "SELECT * FROM Win32_Printer";
            //CimSession mySession = CimSession.Create(thePrintServer);
            //Verify the session created correctly, otherwise it will error out if it fails to connect.
            var testtheConnection = session.TestConnection();
            if (testtheConnection == true)
            {
                //Query WMI for Printer Ports on the server.
                IEnumerable<CimInstance> queryInstance = session.QueryInstances(Namespace, "WQL", OSQuery);
                //Check to see if it exists in the query response.
                CimInstance exist = queryInstance.Where(p => p.CimInstanceProperties["Name"].Value.Equals(theNewPrinter.DeviceID)).FirstOrDefault();
                //If the port exists, then return true.
                if (exist != null)
                {
                    //return true and exit.
                    exist.CimInstanceProperties["EnableBIDI"].Value = false;
                    exist.CimInstanceProperties["DoCompleteFirst"].Value = true;
                    exist.CimInstanceProperties["RawOnly"].Value = true;
                    try
                    {
                        session.ModifyInstance(Namespace, exist);
                        //mySession.Dispose();
                        return (theNewPrinter.DeviceID + " was added successfully on " + printServer.Name);
                    }
                    catch
                    {
                        //mySession.Dispose();
                        return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + printServer.Name);
                    }
                }
                else
                {
                    //item doesn't exist...
                    //mySession.Dispose();
                    return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + printServer.Name);
                }
            }
            else
            {
                //session didn't connect.  Return False
                return (theNewPrinter.DeviceID + " was added, but the printer properties were not set on " + printServer);
            }

            //return (theNewPrinter.DeviceID + " was added successfully on " + thePrintServer);

        }
        static private string DeletePrinter(string printer, PrintServer server)
        {
            try
            {
                PrintQueue thePrintQueue = server.GetPrintQueue(printer);
                PrintServer.DeletePrintQueue(thePrintQueue);
                thePrintQueue.Dispose();
                return "Successfully deleted: " + printer + " on server: " + server.Name.TrimStart('\\');
            }
            catch
            {
                return ("Delete failed on server: " + server.Name.TrimStart('\\') + " for printer: " + printer + ".  Printer does not exist.");
            }
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
            if (col.Count == 0)
            {
                s.Dispose();
                return "Delete failed on server: " + server + ".  Printer does not exist.";
            }
            try
            {
                foreach (ManagementObject i in col)
                {
                    i.Delete();
                    return "Successfully deleted: " + i["Name"].ToString() + " on server: " + server;
                }
                s.Dispose();
            }
            catch
            {
                s.Dispose();
                return "Delete failed on server: " + server + ".  Try catch failed";

            }
            return "Delete failed on server: " + server + ". Returned failed at the end of the method.";
        }
        static private bool ClearPrintQueue(string printer, string server)
        {
            try
            {
                using (PrintServer ps = new PrintServer(@"\\" + server))
                {
                    using (PrintQueue pq = new PrintQueue(ps, printer, PrintSystemDesiredAccess.AdministratePrinter))
                    {
                        pq.Purge();
                        return true;
                    }
                }

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
                checkprintServer.Dispose();
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
        static public string GetCurrentPrintTray(string printer, string printserver)
        {
            string thePrintTray = null;

            try
            {
                //PrintServer class requires the 2 wacks in the server name.
                PrintServer printServer = new PrintServer(@"\\" + printserver);
                //Get the one print queue from the print server.
                var myPrintQueues = printServer.GetPrintQueue(printer);
                PrintTicket someTicket = myPrintQueues.DefaultPrintTicket;
                MemoryStream myTicket = new MemoryStream();
                myTicket = myPrintQueues.DefaultPrintTicket.GetXmlStream();
                XmlDocument myXmlDoc = new XmlDocument();
                myXmlDoc.Load(myTicket);
                myTicket.Dispose();
                XmlNode node;
                node = myXmlDoc.DocumentElement;
                foreach (XmlNode node1 in node.ChildNodes)
                    foreach (XmlNode node2 in node1.ChildNodes)
                    {
                        if (node1.Attributes["name"].Value.Contains("JobInputBin"))  //This is looking for psk:JobInputBin.  Works for HP UPD for now....
                        {
                            //tray must be of pskAutoSelect ns0000:Tray1 ns0000:Tray2 ns0000:Tray3
                            thePrintTray = node2.Attributes["name"].Value;
                        }
                    }
                myPrintQueues.Dispose();
                printServer.Dispose();
                if (thePrintTray == null)
                {
                    return ("Error with Tray Lookup.  ThePrintTray is Null");
                }
                int posA = thePrintTray.LastIndexOf(":");
                string search = ":";
                int adjustedPosA = posA + search.Length;
                string finalTray = thePrintTray.Substring(adjustedPosA);
                return (finalTray);

            }
            catch (Exception e)
            {
                return (e.Message);
            }
        }
        //Returns the current print tray a print queue is using.
        static public string GetCurrentPrintTray(PrintQueue myPrintQueues)
        {
            string thePrintTray = null;

            try
            {
                //PrintServer class requires the 2 wacks in the server name.
                //PrintServer printServer = new PrintServer(@"\\" + printserver);
                //Get the one print queue from the print server.
                //var myPrintQueues = printServer.GetPrintQueue(printer);
                PrintTicket someTicket = myPrintQueues.DefaultPrintTicket;
                MemoryStream myTicket = new MemoryStream();
                myTicket = myPrintQueues.DefaultPrintTicket.GetXmlStream();
                XmlDocument myXmlDoc = new XmlDocument();
                myXmlDoc.Load(myTicket);
                myTicket.Dispose();
                XmlNode node;
                node = myXmlDoc.DocumentElement;
                foreach (XmlNode node1 in node.ChildNodes)
                    foreach (XmlNode node2 in node1.ChildNodes)
                    {
                        if (node1.Attributes["name"].Value.Contains("JobInputBin"))  //This is looking for psk:JobInputBin.  Works for HP UPD for now....
                        {
                            //tray must be of pskAutoSelect ns0000:Tray1 ns0000:Tray2 ns0000:Tray3
                            thePrintTray = node2.Attributes["name"].Value;
                        }
                    }
                //myPrintQueues.Dispose();
                //printServer.Dispose();
                if (thePrintTray == null)
                {
                    return ("Error with Tray Lookup.");
                }
                int posA = thePrintTray.LastIndexOf(":");
                string search = ":";
                int adjustedPosA = posA + search.Length;
                string finalTray = thePrintTray.Substring(adjustedPosA);
                return (finalTray);

            }
            catch (Exception e)
            {
                return (e.Message);
            }
            //return ("Shouldn't get this far...");
        }
        static public string SetPrinterTray(string printer, string printserver, string printtray)
        {
            string printTray = null;
            if (printtray.Contains("AutoSelect"))
            {
                printTray = "psk:AutoSelect";
            }
            else if (printtray.Contains("Tray"))
            {
                printTray = "ns0000:" + printtray;
            }
            else
            {
                return ("Printer Tray is not a valid or undefined.  Using default print tray.");
            }

            //PrintServer class requires the 2 wacks in the server name.
            PrintServer printServer = new PrintServer(@"\\" + printserver);
            //Get the one print queue from the print server.
            PrintQueue myPrintQueues = printServer.GetPrintQueue(printer);
            //PrintTicket someTicket = myPrintQueues.DefaultPrintTicket;
            MemoryStream myTicket = myPrintQueues.DefaultPrintTicket.GetXmlStream();
            XmlDocument myXmlDoc = new XmlDocument();
            myXmlDoc.Load(myTicket);
            XmlNode node;
            node = myXmlDoc.DocumentElement;
            foreach (XmlNode node1 in node.ChildNodes)
                foreach (XmlNode node2 in node1.ChildNodes)
                {
                    if (node1.Attributes["name"].Value.Contains("JobInputBin"))  //This is looking for psk:JobInputBin.  Works for HP UPD for now....
                    {
                        //tray must be of pskAutoSelect ns0000:Tray1 ns0000:Tray2 ns0000:Tray3
                        node2.Attributes["name"].Value = printTray;
                    }
                }
            MemoryStream xmlStream = new MemoryStream();
            myXmlDoc.Save(xmlStream);
            xmlStream.Position = 0;
            PrintQueue pqAdmin = new PrintQueue(printServer, printer, PrintSystemDesiredAccess.AdministratePrinter);
            PrintTicket newPrintTicket = new PrintTicket(xmlStream);
            pqAdmin.DefaultPrintTicket = newPrintTicket;
            pqAdmin.Commit();
            printServer.Dispose();
            return ("Printer Tray: " + printtray + " has been set on printer: " + printer + " on print server: " + printserver);
        }
        static public string GetPrinterStatus(PrintQueue pq)
        {
            string statusReport = null;
            if ((pq.QueueStatus & PrintQueueStatus.PaperProblem) == PrintQueueStatus.PaperProblem)
            {
                statusReport = statusReport + "Has a paper problem. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.NoToner) == PrintQueueStatus.NoToner)
            {
                statusReport = statusReport + "Is out of toner. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.DoorOpen) == PrintQueueStatus.DoorOpen)
            {
                statusReport = statusReport + "Has an open door. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.Error) == PrintQueueStatus.Error)
            {
                statusReport = statusReport + "Is in an error state. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.NotAvailable) == PrintQueueStatus.NotAvailable)
            {
                statusReport = statusReport + "Is not available. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.Offline) == PrintQueueStatus.Offline)
            {
                statusReport = statusReport + "Is off line. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.OutOfMemory) == PrintQueueStatus.OutOfMemory)
            {
                statusReport = statusReport + "Is out of memory. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.PaperOut) == PrintQueueStatus.PaperOut)
            {
                statusReport = statusReport + "Is out of paper. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.OutputBinFull) == PrintQueueStatus.OutputBinFull)
            {
                statusReport = statusReport + "Has a full output bin. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.PaperJam) == PrintQueueStatus.PaperJam)
            {
                statusReport = statusReport + "Has a paper jam. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.Paused) == PrintQueueStatus.Paused)
            {
                statusReport = statusReport + "Is paused. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.TonerLow) == PrintQueueStatus.TonerLow)
            {
                statusReport = statusReport + "Is low on toner. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.UserIntervention) == PrintQueueStatus.UserIntervention)
            {
                statusReport = statusReport + "Needs user intervention. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.Busy) == PrintQueueStatus.Busy)
            {
                statusReport = statusReport + "Currently Busy";
            }
            else
            {
                statusReport = "Available";
            }
            return (statusReport);
        }
        static public bool PrintTestPage(string printServer, string printer)
        {
            try
            {
                ManagementScope scope = new ManagementScope(String.Format(@"\\{0}\ROOT\CIMV2", printServer));
                scope.Connect();
                ManagementPath managementPath = new ManagementPath("Win32_Process");
                ManagementClass processClass = new ManagementClass(scope, managementPath, new ObjectGetOptions());
                ManagementBaseObject inParams = processClass.GetMethodParameters("Create");
                inParams["CommandLine"] = @"C:\Windows\System32\rundll32.exe printui.dll,PrintUIEntry /q /k /n" + printer;
                ManagementBaseObject outParams = processClass.InvokeMethod("Create", inParams, null);
                uint rtn = System.Convert.ToUInt32(outParams["returnValue"]);
                uint processID = System.Convert.ToUInt32(outParams["processId"]);

                System.Threading.Thread.Sleep(1000);
                var processName = "rundll32.exe";
                var query = new SelectQuery("select * from Win32_process where name = '" + processName + "' AND processid = '" + processID + "'");
                using (var searcher = new ManagementObjectSearcher(scope, query))
                {
                    foreach (ManagementObject process in searcher.Get()) // this is the fixed line
                    {
                        process.InvokeMethod("Terminate", null);
                    }
                }
                return true;
            }
            catch
            {

                return false;
            }

        }
        //Used for parallel and AJAX processing of new EPS printer creation.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public JsonResult CreatePrinterJSONexeVersion([Bind(Include = "DeviceID,DriverName,PortName,Tray,Comments,Location,SourcePrinter")]AddPrinterClass theNewPrinter)
        {

            //Do some timing on the whole process.
            Stopwatch newwatch = new Stopwatch();
            newwatch.Start();

            

            //Initialize list to display to end users if it completes or not.
            List<string> outcome = new List<string>();

            //Used to determine if user has correct access to create a printer.
            //If they don't then it returns an error message for them.
            if (!Support.IsUserAuthorized(Support.ADGroupCanAddEPSPrinter()))
            {
                outcome.Add("You are not authorized to add printers.");
                logger.Info("User " + User.Identity.Name + " attempted to create EPS printer: " + theNewPrinter.DeviceID);
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }

            //Make sure incoming data is good to go.
            if (ModelState.IsValid)
            {
                //Validate the printer exists in DNS or Valid IP Address
                //Web.Config determines if it is really used or not.
                if (!Support.ValidatePrinterDNS() || Support.ValidHostname(theNewPrinter.PortName) == true)
                {
                    //List<string> theservers = new List<string>();
                    //theservers.Add("tst-eps-1");
                    var theservers = Support.GetEPSServers();
                    //foreach (var server in theservers)
                    Parallel.ForEach(Support.GetEPSServers(), server =>
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
                                logger.Info("New Port added.  Port Name : " + theNewPrinter.PortName + " on server " + server + " by user " + User.Identity.Name);
                            }
                            else
                            {
                                logger.Info("New Port attempted to be added but already exists.  Port Name : " + theNewPrinter.PortName + " on server " + server + " by user " + User.Identity.Name);
                                //If printer port already exists, don't do anything!
                            }
                            if (CheckCurrentPrinter(theNewPrinter.DeviceID, server))
                            {
                                outcome.Add(theNewPrinter.DeviceID + " already exists on " + server);
                                logger.Info("New Printer attempted to be added but already exists.  Printer Name : " + theNewPrinter.DeviceID + " on server " + server + " by user " + User.Identity.Name);
                            }
                            else
                            {
                                //Try to add the printer now the port is defined on the server.
                                //first setup the props of the printer
                                AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.DeviceID, DriverName = theNewPrinter.DriverName, EnableBIDI = false, PortName = theNewPrinter.PortName, Published = false, Shared = false, SourcePrinter = theNewPrinter.SourcePrinter != null ? theNewPrinter.SourcePrinter:"nothing",SourceServer=theNewPrinter.SourceServer!=null?theNewPrinter.SourceServer:"nothing"  };
                                //Use the string return function to determine if the printer was successfully added or not.
                                string ServerType;
                                if (Support.GetEnterprisePrintServers().Contains(server))
                                {
                                    ServerType = "Enterprise";
                                }
                                else
                                {
                                    ServerType = "EPS";
                                }

                                if (Support.UseEXEForPrinterCreation() == true)
                                {
   
                                    Process cmd = new Process();
                                    cmd.StartInfo.FileName = HostingEnvironment.MapPath("~/ExternalReferences/EPSPrintMgmtConsole.exe");  //@"C:\Users\hm3307\Documents\Visual Studio 2015\Projects\EPSPrintMgmt\EPSPrintMgmtConsole\bin\Debug\EPSPrintMgmtConsole.exe";
                                    cmd.StartInfo.Arguments = "-n " + newPrinter.DeviceID + " -d \"" + newPrinter.DriverName + "\" -p " + newPrinter.PortName + " -t " + ServerType + " -w " + server +" -g " + Support.UseEPSGoldPrinter() + " -h " + newPrinter.SourcePrinter+ " -i " + newPrinter.SourcePrinter;
                                    cmd.StartInfo.CreateNoWindow = true;
                                    cmd.StartInfo.UseShellExecute = false;
                                    //cmd.StartInfo.WorkingDirectory = HostingEnvironment.MapPath("~/ExternalReferences/" + server.ToString() + "/");
                                    cmd.Start();
                                    cmd.WaitForExit();
                                    outcome.Add("Printer: " + newPrinter.DeviceID + " " + Support.ExeReturnCodeParser(cmd.ExitCode) + " on server: " + server);
                                    logger.Debug("exit code: " + Support.ExeReturnCodeParser(cmd.ExitCode));
                                }
                                else
                                {
                                    //outcome.Add(AddNewPrinterStringReturn(newPrinter, server));
                                   var theoutcomehere= AddNewPrinterStringReturn(newPrinter, server);
                                    outcome.Add(theoutcomehere);
                                }


                                logger.Info("New Printer added.  Printer Name : " + theNewPrinter.DeviceID + " on server " + server + " by user " + User.Identity.Name);
                                //if (Support.UseEPSGoldPrinter() && theNewPrinter.SourcePrinter != null)
                                //{
                                //    outcome.Add(clonePrintQueue(newPrinter, server, "localhost"));
                                //}
                                //if (Support.UsePrintTrays())
                                //{
                                //    outcome.Add(SetPrinterTray(theNewPrinter.DeviceID, server, theNewPrinter.Tray));
                                //    logger.Info("Print Tray Set for Printer Name : " + theNewPrinter.DeviceID + " Tray Info: " + theNewPrinter.Tray + " on server " + server + " by user " + User.Identity.Name + " time to install" + newwatch.ElapsedMilliseconds + "ms");
                                //}
                            }
                        }
                        //Cannot connect to the server, so send a message back to the user about it.
                        else
                        {
                            outcome.Add(server + " is not an active or a valid server.  Please verify the server is up or configured correctly in the web.config file.");
                            logger.Info("Current Print Server is not active or invalid.  Print Server: " + server + " by user " + User.Identity.Name);
                        }
                        mySession.Dispose();
                        //End the parallel processing.
                    });

                    newwatch.Stop();
                    //Email users from Web.Config to confirm everything went well!
                    Support.SendEmail("New Printer Added to EPS", string.Join(Environment.NewLine, outcome) + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to install." + Environment.NewLine + "Created by user: " + User.Identity.Name);
                    //Send a success message the Success View.
                    TempData["SuccessMessage"] = "Congrats, the printer installed correctly!  Enjoy your day.";
                    //return JSON results to the AJAX request of the view.
                    outcome.Add(@"<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + @" ms to install.</h5>");
                    return Json(outcome, JsonRequestBehavior.AllowGet);
                }
                newwatch.Stop();
                //Send email and return results if DNS does not exist for the printer.
                Support.SendEmail("Failed EPS Install", "Printer: " + theNewPrinter.DeviceID + " failed the DNS lookup or IP Address validation." + Environment.NewLine + "By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist or it is an invalid IP Address.  Please try again.";
                outcome.Add("Hostname of the Printer does not exist or it is an invalid IP Address.  Please try again.");
                outcome.Add("<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to attempt to install." + @"</h5>");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            newwatch.Stop();
            outcome.Add("`r`n" + newwatch.ElapsedMilliseconds + " ms to install.");
            //Return error message that 
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            outcome.Add("Something went wrong with the Model.  Please try again.");
            logger.Error("Failed to initialize the model");
            return Json(outcome, JsonRequestBehavior.AllowGet);

        }

        //Used for parallel and AJAX processing of new EPS and Enterprise printer creation.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public JsonResult CreateEPSAndEnterprisePrinterJSONexeVersion([Bind(Include = "DeviceID,DriverName,PortName,Tray,Comments,Location,PrintServer,SourceServer,SourcePrinter,ENTDriverName,ENTPortName,ENTSourcePrinter,ENTSourceServer,ENTShared,ENTPublished,ENTEnableBIDI,ENTTray")]AddEPSandENTPrinterClass theNewPrinter)
        {
            //Do some timing on the whole process.
            Stopwatch newwatch = new Stopwatch();
            newwatch.Start();

            //Initialize list to display to end users if it completes or not.
            List<string> outcome = new List<string>();

            //Used to determine if user has correct access to create a printer.
            //If they don't then it returns an error message for them.
            if (!Support.IsUserAuthorized(Support.ADGroupCanAddEPSPrinter()) || !Support.IsUserAuthorized(Support.ADGroupCanAddEnterprisePrinter()))
            {
                outcome.Add("You are not authorized to add EPS and Enterprise printers.");
                logger.Info("User " + User.Identity.Name + " attempted to create EPS and Enterprise printer: " + theNewPrinter.DeviceID);
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }

            //Make sure incoming data is good to go.
            if (ModelState.IsValid)
            {
                //Validate the printer exists in DNS or Valid IP Address
                //Web.Config determines if it is really used or not.
                if (!Support.ValidatePrinterDNS() || Support.ValidHostname(theNewPrinter.PortName) == true)
                {
                    logger.Debug(newwatch.ElapsedMilliseconds + " ms has passed.  Before Parallel foreach");
                    var allPrintServers = Support.GetEPSServers();
                    allPrintServers.Add(theNewPrinter.PrintServer);
                    //kick off multiple threads to install printers quickly for EPS Print Servers.
                    Parallel.ForEach(allPrintServers, server =>
                    {
                        //determine print server type used when adding a new print queue.
                        string ServerType;
                        if (Support.GetEnterprisePrintServers().Contains(server))
                        {
                            ServerType = "Enterprise";
                        }
                        else
                        {
                            ServerType = "EPS";
                        }

                        logger.Debug("Before 1st CimSession creation: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                        //Used to test if it can connect via Cim first.  If it cannot it skips that server.
                        CimSession mySession = CimSession.Create(server);
                        var testtheConnection = mySession.TestConnection();
                        if (testtheConnection == true)
                        {
                            logger.Debug("CimSession creation completed Successfully: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                            //Start the process of installing a printer.
                            //Checks to see if the Printer port already exists on the server.

                            //Start with Enterprise Print Servers.  Have different parameters passed for Enterprise Print Servers.
                            if (Support.GetEnterprisePrintServers().Contains(server))
                            {
                                logger.Debug("Before ENT Port Creation.  Testing port : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                if (ExistingPrinterPort(theNewPrinter.ENTPortName, server) == false)
                                {
                                    logger.Debug("After ENT Port Check and Before Port Creation: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    //Adds printer port, currently just using the name of the printer for the port name.  So DNS must be setup first in order to add printer port.
                                    AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.ENTPortName, HostAddress = theNewPrinter.ENTPortName, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                                    AddNewPrinterPort(AddThePort, server);
                                    logger.Debug("After ENT Port Creation: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    logger.Info("New Port added.  Port Name : " + theNewPrinter.ENTPortName + " on server " + server + " by user " + User.Identity.Name);
                                }
                                else
                                {
                                    logger.Debug("After ENT Port Check with No Ports to add : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    logger.Info("New Port attempted to be added but already exists.  Port Name : " + theNewPrinter.ENTPortName + " on server " + server + " by user " + User.Identity.Name);
                                    //If printer port already exists, don't do anything!
                                }
                            }
                            //Check the EPS Print Servers Port Info
                            else
                            {
                                logger.Debug("Before EPS Port Creation.  Testing port : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                if (ExistingPrinterPort(theNewPrinter.PortName, server) == false)
                                {
                                    logger.Debug("After EPS Port Check and Before Port Creation: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    //Adds printer port, currently just using the name of the printer for the port name.  So DNS must be setup first in order to add printer port.
                                    AddPrinterPortClass AddThePort = new AddPrinterPortClass { Name = theNewPrinter.PortName, HostAddress = theNewPrinter.PortName, PortNumber = 9100, Protocol = 1, SNMPEnabled = false };
                                    AddNewPrinterPort(AddThePort, server);
                                    logger.Debug("After EPS Port Creation: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    logger.Info("New Port added.  Port Name : " + theNewPrinter.PortName + " on server " + server + " by user " + User.Identity.Name);
                                }
                                else
                                {
                                    logger.Debug("After EPS Port Check with No Ports to add : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    logger.Info("New Port attempted to be added but already exists.  Port Name : " + theNewPrinter.PortName + " on server " + server + " by user " + User.Identity.Name);
                                    //If printer port already exists, don't do anything!
                                }
                            }

                            logger.Debug("Before Testing if printer exists : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                            if (CheckCurrentPrinter(theNewPrinter.DeviceID, server))
                            {
                                logger.Debug("After Testing if printer exists and nothing to do: " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                outcome.Add(theNewPrinter.DeviceID + " already exists on " + server);
                                logger.Info("New Printer attempted to be added but already exists.  Printer Name : " + theNewPrinter.DeviceID + " on server " + server + " by user " + User.Identity.Name);
                            }
                            else
                            {
                                //determine if this is an EPS or Enterprise Print Server
                                //Add Enterprise Print Queue First.
                                if (Support.GetEnterprisePrintServers().Contains(server))
                                {
                                    //Try to add the printer now the port is defined on the server.
                                    //first setup the props of the printer
                                    AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.DeviceID, DriverName = theNewPrinter.ENTDriverName, EnableBIDI = false, PortName = theNewPrinter.ENTPortName, Published = false, Shared = false };
                                    PrintServer thePrintServer = new PrintServer(@"\\" + theNewPrinter.PrintServer, PrintSystemDesiredAccess.AdministrateServer);
                                    //Use the string return function to determine if the printer was successfully added or not.
                                    logger.Debug("Before creation of new ENT Print Queue : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);

                                    //determine if you should use the dedicated EXE or not.
                                    if (Support.UseEXEForPrinterCreation() == true)
                                    {
                                        Process cmd = new Process();
                                        cmd.StartInfo.FileName = HostingEnvironment.MapPath("~/ExternalReferences/EPSPrintMgmtConsole.exe");
                                        cmd.StartInfo.Arguments = "-n " + newPrinter.DeviceID + " -d \"" + newPrinter.DriverName + "\" -p " + newPrinter.PortName + " -t " + ServerType + " -w " + server + " -g " + Support.UseEPSGoldPrinter() + " -h " + newPrinter.SourcePrinter + " -i " + newPrinter.SourceServer; ;
                                        cmd.StartInfo.CreateNoWindow = true;
                                        cmd.StartInfo.UseShellExecute = false;
                                        cmd.Start();
                                        cmd.WaitForExit();
                                        outcome.Add("Printer: " + newPrinter.DeviceID + " " + Support.ExeReturnCodeParser(cmd.ExitCode) + " on server: " + server);
                                        logger.Debug("exit code: " + Support.ExeReturnCodeParser(cmd.ExitCode));
                                    }
                                    else
                                    {
                                        outcome.Add(AddNewEnterprisePrinterStringReturn(newPrinter, mySession, thePrintServer));
                                        logger.Debug("After creation of new ENT Print Queue : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                        logger.Info("New Printer added.  Printer Name : " + theNewPrinter.DeviceID + " on server " + theNewPrinter.PrintServer + " by user " + User.Identity.Name);
                                    }
                                    if (Support.UseEnterprisePrintTrays())
                                    {
                                        outcome.Add(SetPrinterTray(theNewPrinter.DeviceID, theNewPrinter.PrintServer, theNewPrinter.Tray));
                                        logger.Info("Print Tray Set for Printer Name : " + theNewPrinter.DeviceID + " Tray Info: " + theNewPrinter.Tray + " on server " + theNewPrinter.PrintServer + " by user " + User.Identity.Name + " time to install" + newwatch.ElapsedMilliseconds + "ms");
                                    }
                                    thePrintServer.Dispose();
                                }
                                //Add EPS Print Queue here.
                                else
                                {
                                    //Try to add the printer now the port is defined on the server.
                                    //first setup the props of the printer
                                    AddPrinterClass newPrinter = new AddPrinterClass { DeviceID = theNewPrinter.DeviceID, DriverName = theNewPrinter.DriverName, EnableBIDI = false, PortName = theNewPrinter.PortName, Published = false, Shared = false, SourcePrinter = theNewPrinter.SourcePrinter };
                                    //Use the string return function to determine if the printer was successfully added or not.
                                    logger.Debug("Before creation of new EPS Print Queue : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);

                                    //determine if you should use the dedicated EXE or not.
                                    if (Support.UseEXEForPrinterCreation() == true)
                                    {
                                        Process cmd = new Process();
                                        cmd.StartInfo.FileName = HostingEnvironment.MapPath("~/ExternalReferences/EPSPrintMgmtConsole.exe");
                                        cmd.StartInfo.Arguments = "-n " + newPrinter.DeviceID + " -d \"" + newPrinter.DriverName + "\" -p " + newPrinter.PortName + " -t " + ServerType + " -w " + server + " -g " + Support.UseEPSGoldPrinter() + " -h " + newPrinter.SourcePrinter + " -i " + newPrinter.SourceServer;
                                        cmd.StartInfo.CreateNoWindow = true;
                                        cmd.Start();
                                        cmd.WaitForExit();
                                        outcome.Add("Printer: " + newPrinter.DeviceID + " " + Support.ExeReturnCodeParser(cmd.ExitCode) + " on server: " + server);
                                        logger.Debug("exit code: " + Support.ExeReturnCodeParser(cmd.ExitCode));
                                    }
                                    else
                                    {
                                    outcome.Add(AddNewPrinterStringReturn(newPrinter, server));
                                    logger.Debug("After creation of new EPS Print Queue : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    logger.Info("New Printer added.  Printer Name : " + theNewPrinter.DeviceID + " on server " + server + " by user " + User.Identity.Name);
                                    }

                                    //if (Support.UseEPSGoldPrinter() && theNewPrinter.SourcePrinter != null)
                                    //{
                                    //    outcome.Add(clonePrintQueue(newPrinter, server, "localhost"));
                                    //    logger.Debug("After Print Queue is Cloned : " + newwatch.ElapsedMilliseconds + " ms has passed on server: " + server);
                                    //}
                                    //if (Support.UsePrintTrays())
                                    //{
                                    //    outcome.Add(SetPrinterTray(theNewPrinter.DeviceID, server, theNewPrinter.Tray));
                                    //    logger.Info("Print Tray Set for Printer Name : " + theNewPrinter.DeviceID + " Tray Info: " + theNewPrinter.Tray + " on server " + server + " by user " + User.Identity.Name + " time to install" + newwatch.ElapsedMilliseconds + "ms");
                                    //}

                                }
                            }
                        }
                        //Cannot connect to the server, so send a message back to the user about it.
                        else
                        {
                            outcome.Add(server + " is not an active or a valid server.  Please verify the server is up or configured correctly in the web.config file.");
                            logger.Info("Current Print Server is not active or invalid.  Print Server: " + server + " by user " + User.Identity.Name);
                        }
                        mySession.Dispose();
                        //End the parallel processing.
                    });
                    newwatch.Stop();
                    //Email users from Web.Config to confirm everything went well!
                    Support.SendEmail("New Printer Added to EPS and Enterprise Print Server", string.Join(Environment.NewLine, outcome) + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to install." + Environment.NewLine + "Created by user: " + User.Identity.Name);
                    //Send a success message the Success View.
                    TempData["SuccessMessage"] = "Congrats, the printer installed correctly!  Enjoy your day.";
                    //return JSON results to the AJAX request of the view.
                    outcome.Add(@"<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + @" ms to install.</h5>");
                    return Json(outcome, JsonRequestBehavior.AllowGet);
                }
                newwatch.Stop();
                //Send email and return results if DNS does not exist for the printer.
                Support.SendEmail("Failed EPS and Enterprise Print Server Install", "Printer: " + theNewPrinter.DeviceID + " failed the DNS lookup or IP Address validation." + Environment.NewLine + "By user: " + User.Identity.Name);
                TempData["RedirectToError"] = "Hostname of the Printer does not exist or it is an invalid IP Address.  Please try again.";
                outcome.Add("Hostname of the Printer does not exist or it is an invalid IP Address.  Please try again.");
                outcome.Add("<h5>" + Environment.NewLine + newwatch.ElapsedMilliseconds + "ms to attempt to install." + @"</h5>");
                return Json(outcome, JsonRequestBehavior.AllowGet);
            }
            newwatch.Stop();
            outcome.Add("`r`n" + newwatch.ElapsedMilliseconds + " ms to install.");
            //Return error message that 
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            outcome.Add("Something went wrong with the Model.  Please try again.");
            logger.Error("Failed to initialize the model");
            return Json(outcome, JsonRequestBehavior.AllowGet);

        }
    }
}