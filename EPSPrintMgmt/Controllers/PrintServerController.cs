using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;
using EPSPrintMgmt.Models;
using System.Management;
using System.Printing;
using System.Net;

namespace EPSPrintMgmt.Controllers
{
    public class PrintServerController : Controller
    {
        // GET: PrintServer
        public ActionResult Index()
        {
            List<MyPrintServer> myPrintServers = new List<MyPrintServer>();
            foreach(var server in Support.GetAllPrintServers())
            {
                string printerCount;
                string ipAddress;
                try
                {
                    PrintServer printServer = new PrintServer(@"\\"+server, PrintSystemDesiredAccess.AdministrateServer);
                    printerCount = printServer.GetPrintQueues().Count().ToString();
                }
                catch
                {
                    printerCount = "N/A";
                }

                try
                {
                    ipAddress = Dns.GetHostEntry(server).AddressList[0].ToString();
                }
                catch
                {
                    ipAddress = "N/A";
                }

                myPrintServers.Add(new MyPrintServer { Name = server.ToString() ,PrinterCount=printerCount ,IP=ipAddress});
            }
            return View(myPrintServers.OrderBy(o=>o.Name));
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
        public ActionResult Options(string printServer)
        {
            MyPrintServer myPrintServers = new MyPrintServer { Name=printServer};
            return View(myPrintServers);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Options([Bind(Include = "Name")]MyPrintServer myPrintServer)
        {
            if (ModelState.IsValid)
            {
                if (FlushDNSCache(myPrintServer.Name.ToString()) == true)
                {
                    Support.SendEmail("DNS flushed from Print Server.","DNS has been flushed on the following computer: "+myPrintServer.Name+ "  by user: " + User.Identity.Name);
                    TempData["SuccessMessage"] = "Congrats, DNS has been flushed from "+myPrintServer.Name+"!  Enjoy your day.";
                    return RedirectToAction("Success");
                }
                else
                {
                    Support.SendEmail("Failed to flush DNS", "DNS has failed to flush on the following computer: " + myPrintServer.Name + "  by user: " + User.Identity.Name);
                    TempData["RedirectToError"] = "Could not flush DNS on "+myPrintServer.Name+".  Please try again or logon to the server directly to clear it.";
                    return RedirectToAction("Error");
                }
            }
            TempData["RedirectToError"] = "Something went wrong with the Model.  Please try again.";
            return RedirectToAction("Error");
        }

        static private bool FlushDNSCache(string myPrintServer)
        {
            object[] theProcessToRun = { "cmd.exe /C ipconfig /flushdns" };
            //object[] theProcessToRun = { "notepad.exe" };
            ConnectionOptions theConnection = new ConnectionOptions();
            //theConnection.Username = "username";
            //theConnection.Password = "password";
            theConnection.EnablePrivileges = true;
            ManagementScope theScope = new ManagementScope("\\\\" + myPrintServer + "\\root\\cimv2", theConnection);
            ManagementClass theClass = new ManagementClass(theScope, new ManagementPath("Win32_Process"), new ObjectGetOptions());
            var output = theClass.InvokeMethod("Create", theProcessToRun);
            if (output.ToString() == "0")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}