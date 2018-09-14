using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Configuration;

namespace EPSPrintMgmt.Models
{

    public class Support
    {
        public static void SendEmail(string subject, string body)
        {
            MailMessage message = new MailMessage(GetEmailFrom(), GetEmailTo(), subject, body);

            SmtpClient mailClient = new SmtpClient(GetRelayServer());
            try
            {
                mailClient.Send(message);
                mailClient.Dispose();
            }
            catch
            {
                //do something useful some day...
            }
        }
        public static void SendEnterpriseEmail(string subject, string body)
        {
            MailMessage message = new MailMessage(GetEmailFrom(), GetEnterpriseEmailTo(), subject, body);

            SmtpClient mailClient = new SmtpClient(GetRelayServer());
            try
            {
                mailClient.Send(message);
                mailClient.Dispose();
            }
            catch
            {
                //do something useful some day...
            }
        }
        static private string GetRelayServer()
        {
            string relayServer = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("MailRelay")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return (relayServer);
        }
        static private string GetEmailTo()
        {
            string relayServer = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EmailTo")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return (relayServer);
        }
        static private string GetEnterpriseEmailTo()
        {
            string relayServer = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EmailEnterpriseTo")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return (relayServer);
        }
        static private string GetEmailFrom()
        {
            string relayServer = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EmailFrom")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return (relayServer);
        }
        //Currently checks to validate it's an actual IP address or a valid DNS entry.
        //Doesn't Ping the IP address, just makes sure it's a valid IPV4 or IPV6 address.
        public static bool ValidHostname(string hostname)
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
            catch //(System.Net.Sockets.SocketException e)
            {
                //Console.WriteLine(e.Message);
                return false;
            }
            return true;
        }
        static public bool UsePrinterIPAddr()
        {
            string useIPAddr = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("UsePrinterIPAddress")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(useIPAddr.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;
        }
        static public List<string> GetEPSServers()
        {
            List<string> epsServers = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EPSServer")).Select(k => ConfigurationManager.AppSettings[k]).ToList();
            return (epsServers);
        }
        static public List<string> GetAllPrintServers()
        {
            List<string> allServers = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EPSServer") ||k.StartsWith("EnterpriseServer")).Select(k => ConfigurationManager.AppSettings[k]).ToList();

            return (allServers);
        }
        static public List<string> GetEnterprisePrintServers()
        {
            //List<string> allServers = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EnterpriseServer")).Select(k => ConfigurationManager.AppSettings[k]).OrderBy(x => x).ToList();
            List<string> allServers = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EnterpriseServer")).Select(k => ConfigurationManager.AppSettings[k]).ToList();
            return (allServers);
        }
        static public List<string> GetAllPrintDrivers()
        {
            List<string> printDrivers = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("PrintDriver")).Select(k => ConfigurationManager.AppSettings[k]).ToList();
            return (printDrivers);
        }
        static public List<string> GetAllEnterprisePrintDrivers()
        {
            List<string> printDrivers = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EnterprisePD")).Select(k => ConfigurationManager.AppSettings[k]).OrderBy(k=>k).ToList();
            return (printDrivers);
        }
        static public List<string> GetEPSGoldPrinters()
        {
            List<string> epsGoldPrinters = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EPSGoldPrinter")).Select(k => ConfigurationManager.AppSettings[k]).ToList();
            return (epsGoldPrinters);
        }
        static public string GetEPSGoldPrintServer()
        {
            string goldServer = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EPSGoldPrintServer")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return (goldServer);
        }
        static public bool UseEPSGoldPrinter()
        {
            string useEPSGold = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("UseEPSGoldPrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(useEPSGold.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool clonePrinterDeviceSettings()
        {
            string clonePrinterDevSets = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("CloneDeviceSettings")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(clonePrinterDevSets.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public List<string> GetEntGoldPrinters()
        {
            List<string> epsGoldPrinters = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EntGoldPrinter")).Select(k => ConfigurationManager.AppSettings[k]).ToList();
            return (epsGoldPrinters);
        }
        static public string GetEntGoldPrintServer()
        {
            string goldServer = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EntGoldPrintServer")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return (goldServer);
        }
        static public bool UseEntGoldPrinter()
        {
            string useEPSGold = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("UseEntGoldPrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(useEPSGold.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool cloneEntPrinterDeviceSettings()
        {
            string clonePrinterDevSets = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("CloneEntDeviceSettings")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(clonePrinterDevSets.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public List<string> GetTrays()
        {
            List<string> trays = new List<string>(new string[] { "AutoSelect", "Tray1", "Tray2", "Tray3", "Tray4", "Tray5", "Tray6" });
            return (trays);
        }
        static public bool UsePrintTrays()
        {
            string usePrintTrays = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("UsePrintTrays")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(usePrintTrays.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool AutoPrintWindowsTestPage()
        {
            string printTest = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("AutoPrintWindowsTestPage")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(printTest.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool UseEnterprisePrintTrays()
        {
            string usePrintTrays = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("UseEnterprisePrintTrays")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(usePrintTrays.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool UseEnterprisePrinterBiDirectionalSupport()
        {
            string UseEnterprisePrinterBiDirectionalSupport = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EnterprisePrinterBiDirectionalSupport")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(UseEnterprisePrinterBiDirectionalSupport.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool EnterpriseUsePrinterIPAddr()
        {
            string useIPAddr = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("UseEnterprisePrinterIPAddress")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(useIPAddr, "true", true) == 0)
            {
                return true;
            }
            return false;
        }
        static public bool ValidatePrinterDNS()
        {
            string validDNS = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("ValidatePrinterDNS")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(validDNS.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool ShowNumberPrintJobs()
        {
            string validDNS = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("ShowNumberOfJobs")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(validDNS.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool AllowEPSPrintDeletion()
        {
            string deleteEPSPrinter = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("AllowEPSPrinterDeletion")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(deleteEPSPrinter, "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool AddEnterprisePrinters()
        {
            string useEPSAndEnterprisePrinters = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("AllowEnterprisePrintCreation")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(useEPSAndEnterprisePrinters.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool EditEnterprisePrinters()
        {
            string editEntPrint = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EditEnterprisePrinters")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(editEntPrint.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool AddEPSAndEnterprisePrinters()
        {
            string useEPSAndEnterprisePrinters = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("AllowEPSAndEnterprisePrintCreation")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(useEPSAndEnterprisePrinters.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool AdditionalSecurity()
        {
            string theSecurity = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("AdditionalSecurity")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(theSecurity.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;
        }
        static public string ADGroupCanDeleteEPSPrinter()
        {
            string adGroup = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("ADGrouptoDeleteEPSPrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return adGroup;
        }
        static public string ADGroupCanDeleteEnterprisePrinter()
        {
            string adGroup = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("ADGrouptoDeleteEnterprisePrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return adGroup;
        }
        static public string ADGroupCanAddEPSPrinter()
        {
            string adGroup = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("ADGrouptoAddEPSPrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return adGroup;
        }
        static public string ADGroupCanAddEnterprisePrinter()
        {
            string adGroup = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("ADGrouptoAddEnterprisePrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return adGroup;
        }
        static public string ADGroupCanAddEPSAndEnterprisePrinter()
        {
            string adGroup = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("ADGrouptoAddEPSAndEnterprisePrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return adGroup;
        }
        static public string ADGroupCanPurgePrintQueues()
        {
            string adGroup = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("ADGrouptoPurgePrintQueues")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return adGroup;
        }
        static public string ADGroupCanEditEPSPrinter()
        {
            string adGroup = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("ADGrouptoEditEPSPrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return adGroup;
        }
        static public string ADGroupCanEditEnterprisePrinter()
        {
            string adGroup = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("ADGrouptoEditEnterprisePrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return adGroup;
        }
        static public string ADGroupCanViewPrintServers()
        {
            string adGroup = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("ADGrouptoViewPrintServers")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return adGroup;
        }
        static public bool IsUserAuthorized(string adGroup)
        {
            log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            if (Support.AdditionalSecurity() == true)
            {
                bool isInRole = HttpContext.Current.User.IsInRole(adGroup);
                var httpTest = HttpContext.Current.User.Identity;
                foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(httpTest))
                {
                    string name = descriptor.Name;
                    object value = descriptor.GetValue(httpTest);
                    logger.Debug(name+"="+value);
                }
                foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(httpTest.AuthenticationType))
                {
                    string name = descriptor.Name;
                    object value = descriptor.GetValue(httpTest.AuthenticationType);
                    logger.Debug(name + "=" + value);
                }
                foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(httpTest.IsAuthenticated))
                {
                    string name = descriptor.Name;
                    object value = descriptor.GetValue(httpTest.IsAuthenticated);
                    logger.Debug(name + "=" + value);
                }
                foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(httpTest.Name))
                {
                    string name = descriptor.Name;
                    object value = descriptor.GetValue(httpTest.Name);
                    logger.Debug(name + "=" + value);
                }

                logger.Debug("User info.  User ID: " + HttpContext.Current.User.Identity.Name.ToString() + " User is Authenticated: " + HttpContext.Current.User.Identity.IsAuthenticated.ToString() + " User Authenticated Type" + HttpContext.Current.User.Identity.AuthenticationType.ToString() + " AD Group check: "+adGroup + " Is User in AD Group Check: " + isInRole.ToString());
                if (isInRole == false)
                {
                    return false;
                }
            }
            return true;
        }
        static public string ExeReturnCodeParser(int returnCode)
        {
            switch (returnCode)
            {
                case 1:
                    return ("success");
                case 2:
                    return ("success without props");
                case 400:
                    return ("failed");
                case 404:
                    return ("failed missing parameters");
                case 500:
                    return ("no idea how we got there");
                        
            }
            return ("");
        }
        static public void UpdateAppSettingsConfig()
        {
            var saveConfig = false;
            var config = WebConfigurationManager.OpenWebConfiguration("~");
            //Update AppSettings.config file with defaults.
            //Start with EPS Servers
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EPSServer")).Select(k => ConfigurationManager.AppSettings[k]).ToList().Count() == 0)
            {
                config.AppSettings.Settings.Add("EPSServer1", "You-eps-server-Name");
                config.AppSettings.Settings.Add("EPSServer2", "You-eps-server-Name");
                config.AppSettings.Settings.Add("EPSServer3", "You-eps-server-Name");
                config.AppSettings.Settings.Add("EPSServer4", "You-eps-server-Name");
                saveConfig = true;
            }
            //Add Defaults for Enterprise Print Servers.
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EnterpriseServer")).Select(k => ConfigurationManager.AppSettings[k]).ToList().Count() == 0)
            {
                config.AppSettings.Settings.Add("EnterpriseServer1", "You-Enterprise-Print-Server-Name");
                config.AppSettings.Settings.Add("EnterpriseServer2", "You-Enterprise-Print-Server-Name");
                saveConfig = true;
            }
            //Add Defaults for EPS Print Drivers
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("PrintDriver")).Select(k => ConfigurationManager.AppSettings[k]).ToList().Count() == 0)
            {
                config.AppSettings.Settings.Add("PrintDriver1", "HP Universal Printing PCL 5 (v5.6.5)");
                config.AppSettings.Settings.Add("PrintDriver2", "HP Universal Printing PS (v5.7.0)");
                config.AppSettings.Settings.Add("PrintDriver3", "ZDesigner GX420d");
                saveConfig = true;
            }
            //Add Defaults for Enterprise Print Drivers
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EnterprisePD")).Select(k => ConfigurationManager.AppSettings[k]).ToList().Count() == 0)
            {
                config.AppSettings.Settings.Add("EnterprisePD1", "HP Universal Printing PCL 5 (v5.6.5)");
                config.AppSettings.Settings.Add("EnterprisePD2", "Xerox GPD PCL6 V3.2.303.16.0");
                saveConfig = true;
            }
            //Add the ability to clone EPS Print Driver
            if(ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("UseEPSGoldPrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("UseEPSGoldPrinter", "false");
                saveConfig = true;
            }
            //Add the ability to clone EPS Printer Device Settings
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("CloneDeviceSettings")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("CloneDeviceSettings", "false");
                saveConfig = true;
            }
            //Define your EPS Gold Print Server
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EPSGoldPrintServer")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("EPSGoldPrintServer", "Source-EPS-Printserver-for-Cloning-EPS-Print-Queues");
                saveConfig = true;
            }
            //Add Defaults EPS Gold Print Queues
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EPSGoldPrinter")).Select(k => ConfigurationManager.AppSettings[k]).ToList().Count() == 0)
            {
                config.AppSettings.Settings.Add("EPSGoldPrinter1", "GoldPrinter1-HP");
                config.AppSettings.Settings.Add("EPSGoldPrinter2", "GoldPrinter2-Cannon");
                config.AppSettings.Settings.Add("EPSGoldPrinter3", "GoldPrinter3-Zebra");
                saveConfig = true;
            }
            //Add the ability to clone Enterprise Print Driver
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("UseEntGoldPrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("UseEntGoldPrinter", "false");
                saveConfig = true;
            }
            //Add the ability to clone Enterprise Printer Device Settings
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("CloneEntDeviceSettings")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("CloneEntDeviceSettings", "false");
                saveConfig = true;
            }
            //Define your Enterprise Gold Print Server
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EntGoldPrintServer")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("EntGoldPrintServer", "Source-Enterprise-Printserver-for-Cloning-Enterprise-Print-Queues");
                saveConfig = true;
            }
            //Add Defaults Enterprise Gold Print Queues
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EntGoldPrinter")).Select(k => ConfigurationManager.AppSettings[k]).ToList().Count() == 0)
            {
                config.AppSettings.Settings.Add("EntGoldPrinter1", "GoldPrinter1-HP");
                config.AppSettings.Settings.Add("EntGoldPrinter2", "GoldPrinter2-Cannon");
                config.AppSettings.Settings.Add("EntGoldPrinter3", "GoldPrinter3-Xerox");
                saveConfig = true;
            }
            //Define your Enterprise Gold Print Server
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EntGoldPrintServer")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("EntGoldPrintServer", "Source-Enterprise-Printserver-for-Cloning-Enterprise-Print-Queues");
                saveConfig = true;
            }
            //Define Email Relay Settings
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("MailRelay")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("MailRelay", "emailrelay.hostname.com");
                saveConfig = true;
            }
            //Define Email To Settings
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EmailTo")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("EmailTo", "emailaddress@host.com");
                saveConfig = true;
            }
            //Define Email To Settings
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EmailFrom")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("EmailFrom", "EPSPrintMgmt@host.com");
                saveConfig = true;
            }
            //Define Email To Settings
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EmailEnterpriseTo")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("EmailEnterpriseTo", "emailaddress@host.com");
                saveConfig = true;
            }
            //Define Organization Name
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("OrgName")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("OrgName", "Your Org Name Here");
                saveConfig = true;
            }
            //Use IP or DNS for Printers
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("UsePrinterIPAddress")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("UsePrinterIPAddress", "false");
                saveConfig = true;
            }
            //Do you want to validate the DNS settings
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("ValidatePrinterDNS")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("ValidatePrinterDNS", "true");
                saveConfig = true;
            }
            //Use Print Trays.  THis is really deprecated... 
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("UsePrintTrays")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("UsePrintTrays", "false");
                saveConfig = true;
            }
            //Use Enterprise Print trays.  Again going away...
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("UseEnterprisePrintTrays")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("UseEnterprisePrintTrays", "false");
                saveConfig = true;
            }
            //Show number of print jobs
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("ShowNumberOfJobs")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("ShowNumberOfJobs", "false");
                saveConfig = true;
            }
            //Allow EPS Print Queues to be deleted.
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("AllowEPSPrinterDeletion")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("AllowEPSPrinterDeletion", "false");
                saveConfig = true;
            }
            //Ability to create Enterprise Print Queues
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("AllowEnterprisePrintCreation")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("AllowEnterprisePrintCreation", "true");
                saveConfig = true;
            }
            //Edit Enterprise PRint queues.  
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EditEnterprisePrinters")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("EditEnterprisePrinters", "false");
                saveConfig = true;
            }
            //Bidirectional Support for Enterprise Print Queues
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("EnterprisePrinterBiDirectionalSupport")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("EnterprisePrinterBiDirectionalSupport", "true");
                saveConfig = true;
            }
            //
            //Setup Field to allow IP addresses for Enterprise Print queue
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("UseEnterprisePrinterIPAddress")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("UseEnterprisePrinterIPAddress", "true");
                saveConfig = true;
            }
            //Allow the ability to add EPS and Enterprise print queues at the same tim.
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("AllowEPSAndEnterprisePrintCreation")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("AllowEPSAndEnterprisePrintCreation", "true");
                saveConfig = true;
            }
            //Allow EPS Print Queues to be deleted
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("AllowEPSPrinterDeletion")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("AllowEPSPrinterDeletion", "true");
                saveConfig = true;
            }
            ////Use external EXE.  This will be deprecated....
            //if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("UseEXEForPrinterCreation")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            //{
            //    config.AppSettings.Settings.Add("UseEXEForPrinterCreation", "false");
            //    saveConfig = true;
            //}
            //Setup additional Security
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("AdditionalSecurity")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("AdditionalSecurity", "false");
                saveConfig = true;
            }
            //AD Group to allow EPS Printer Deletion
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("ADGrouptoDeleteEPSPrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("ADGrouptoDeleteEPSPrinter", @"Domain\ADGroup");
                saveConfig = true;
            }
            //AD Group to allow Enterprise Printer Deletion
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("ADGrouptoDeleteEnterprisePrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("ADGrouptoDeleteEnterprisePrinter", @"Domain\ADGroup");
                saveConfig = true;
            }
            //AD Group to all new EPS Print Queues
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("ADGrouptoAddEPSPrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("ADGrouptoAddEPSPrinter", @"Domain\ADGroup");
                saveConfig = true;
            }
            //AD Group to all new Enterprise Print Queues
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("ADGrouptoAddEnterprisePrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("ADGrouptoAddEnterprisePrinter", @"Domain\ADGroup");
                saveConfig = true;
            }
            //AD Group that can Edit EPS Print Queues
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("ADGrouptoEditEPSPrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("ADGrouptoEditEPSPrinter", @"Domain\ADGroup");
                saveConfig = true;
            }
            //AD Group that can Edit Enterprise Print Queues
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("ADGrouptoEditEnterprisePrinter")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("ADGrouptoEditEnterprisePrinter", @"Domain\ADGroup");
                saveConfig = true;
            }
            //AD Group that can Purge Print Queues
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("ADGrouptoPurgePrintQueues")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("ADGrouptoPurgePrintQueues", @"Domain\ADGroup");
                saveConfig = true;
            }
            //AD Group that can View Print Servers
            if (ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("ADGrouptoViewPrintServers")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault() == null)
            {
                config.AppSettings.Settings.Add("ADGrouptoViewPrintServers", @"Domain\ADGroup");
                saveConfig = true;
            }


            //Only save the config if it's needed.
            if (saveConfig == true)
            {
            config.Save();

            }
            
        }

        //Infoblox Specific items
        static public bool ReserverInfobloxIP()
        {
            string reservewInfobloxIP = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("AbilityToReserveDHCPinInfoblox")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(reservewInfobloxIP.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;
        }
        //Auto Reserver IP after adding Enterprise Print queue
        static public bool AutoAddIPInfoblox()
        {
            string autoAdd = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("AutoReserveDHCPinInfoblox")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(autoAdd.ToLower(), "true", true) == 0)
            {
                return true;
            }
            return false;
        }
        //Get Infoblox Server Name
        static public string GetInfobloxServerName()
        {
            string ServerName = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("InfobloxServerName")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return (ServerName);
        }
        //Get Infoblox User Name
        static public string GetInfobloxUserName()
        {
            string UserName = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("InfobloxUsername")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return (UserName);
        }
        //Get Infoblox Password Name
        static public string GetInfobloxPassword()
        {
            string userPassword = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("InfobloxPassword")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            return (userPassword);
        }
        public static async Task<InfobloxIPInformation> getIPInfo(string ipAddress)
        {
            using (var client = new HttpClient())
            {
                ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
                InfobloxIPInformation theIP = null;
                var byteArray = Encoding.ASCII.GetBytes(GetInfobloxUserName()+":"+GetInfobloxPassword());
                client.BaseAddress = new Uri("https://"+GetInfobloxServerName()+"/");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));

                HttpResponseMessage response = await client.GetAsync("wapi/v2.5/ipv4address?ip_address="+ipAddress+"&_return_as_object=1").ConfigureAwait(false);
                if (response.IsSuccessStatusCode)
                {
                    theIP = await response.Content.ReadAsAsync<InfobloxIPInformation>();
                }
                //var contents = await response.Content.ReadAsStringAsync();

                return (theIP);
            }
        }

        public static async Task<string> updateInfoblox(InfobloxReservedIPInfo theAddress)
        {
            using (var client = new HttpClient())
            {
                ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls | SecurityProtocolType.Tls12;
                //ServicePointManager.UseNagleAlgorithm = true;
                string theIP = null;
                var byteArray = Encoding.ASCII.GetBytes(GetInfobloxUserName() + ":" + GetInfobloxPassword());
                client.BaseAddress = new Uri("https://"+GetInfobloxServerName()+"/");

                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                var testinstdsf = JsonConvert.SerializeObject(theAddress);
                var yesplease = JToken.Parse(testinstdsf).ToString();
                var content = new StringContent(JsonConvert.SerializeObject(theAddress), Encoding.UTF8, "application/json");

                var response = await client.PostAsync("wapi/v2.5/fixedaddress", content).ConfigureAwait(false);
                //var response = await client.PostAsJsonAsync(new Uri("https://infoblox_a/wapi/v2.5/fixedaddress"), testinstdsf);
                if (response.IsSuccessStatusCode)
                {
                    theIP = await response.Content.ReadAsStringAsync();
                }
                //var contents = await response.Content.ReadAsStringAsync();

                return (theIP);
            }
        }
        public static async Task<string> RestartInfobloxGrid()
        {
            using (var client = new HttpClient())
            {
                ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls | SecurityProtocolType.Tls12;
                string theIP = null;
                var byteArray = Encoding.ASCII.GetBytes(GetInfobloxUserName()+":"+GetInfobloxPassword());
                client.BaseAddress = new Uri("https://"+GetInfobloxServerName()+"/");

                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                var restartOptions = new InfobloxGridRestartClass { restart_option = "RESTART_IF_NEEDED" };
                var somejson = JsonConvert.SerializeObject(restartOptions);
                var content = new StringContent(JsonConvert.SerializeObject(restartOptions), Encoding.UTF8, "application/json");

                var response = await client.PostAsync(new Uri("https://"+Support.GetInfobloxServerName()+ "/wapi/v2.5/grid/b25lLmNsdXN0ZXIkMA:HCMC-Infoblox?_function=restartservices"), content);
                if (response.IsSuccessStatusCode)
                {
                    theIP = await response.Content.ReadAsStringAsync();
                }

                return (theIP);
            }
        }
        public static async Task<string> GetInfobloxAdvancedInfo(string theAddress)
        {
            using (var client = new HttpClient())
            {
                ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
                string theIP = null;
                var byteArray = Encoding.ASCII.GetBytes(GetInfobloxUserName() + ":" + GetInfobloxPassword());
                client.BaseAddress = new Uri("https://" + GetInfobloxServerName() + "/");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));

                HttpResponseMessage response = await client.GetAsync("wapi/v2.5/search?search_string:~=" + theAddress + "&_return_as_object=1").ConfigureAwait(false);
                if (response.IsSuccessStatusCode)
                {
                    theIP = await response.Content.ReadAsStringAsync();
                }
                //var contents = await response.Content.ReadAsStringAsync();

                return (theIP);
            }
        }
        public static async Task<string> GetInfobloxLeaseInfo(string theAddress)
        {
            using (var client = new HttpClient())
            {
                ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
                string theIP = null;
                var byteArray = Encoding.ASCII.GetBytes(GetInfobloxUserName() + ":" + GetInfobloxPassword());
                client.BaseAddress = new Uri("https://" + GetInfobloxServerName() + "/");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));

                HttpResponseMessage response = await client.GetAsync("wapi/v2.5/lease?address=" + theAddress + "&_return_fields=binding_state,hardware,client_hostname,fingerprint&_return_as_object=1").ConfigureAwait(false);
                if (response.IsSuccessStatusCode)
                {
                    theIP = await response.Content.ReadAsStringAsync();
                }
                //var contents = await response.Content.ReadAsStringAsync();

                return (theIP);
            }
        }
        public static async Task<string> GetInfobloxIPInfo(string theAddress)
        {
            using (var client = new HttpClient())
            {
                ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
                string theIP = null;
                var byteArray = Encoding.ASCII.GetBytes(GetInfobloxUserName() + ":" + GetInfobloxPassword());
                client.BaseAddress = new Uri("https://" + GetInfobloxServerName() + "/");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));

                HttpResponseMessage response = await client.GetAsync("wapi/v2.5/ipv4address?status=USED&ip_address=" + theAddress + "&_return_as_object=1").ConfigureAwait(false);
                if (response.IsSuccessStatusCode)
                {
                    theIP = await response.Content.ReadAsStringAsync();
                }
                //var contents = await response.Content.ReadAsStringAsync();

                return (theIP);
            }
        }
    }
}