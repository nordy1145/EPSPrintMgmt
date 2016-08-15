using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;
using EPSPrintMgmt.Models;
using System.Management;

namespace EPSPrintMgmt.Controllers
{
    public class PrintServerController : Controller
    {
        // GET: PrintServer
        public ActionResult Index()
        {
            List<MyPrintServer> myPrintServers = new List<MyPrintServer>();
            foreach(var server in GetAllPrintServers())
            {
                myPrintServers.Add(new MyPrintServer { Name = server.ToString() });
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
                    SendEmail("DNS flushed from Print Server.","DNS has been flushed on the following computer: "+myPrintServer.Name+ "  by user: " + User.Identity.Name);
                    TempData["SuccessMessage"] = "Congrats, DNS has been flushed from "+myPrintServer.Name+"!  Enjoy your day.";
                    return RedirectToAction("Success");
                }
                else
                {
                    SendEmail("Failed to flush DNS", "DNS has failed to flush on the following computer: " + myPrintServer.Name + "  by user: " + User.Identity.Name);
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


    }
}