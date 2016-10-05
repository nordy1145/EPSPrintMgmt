using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;

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
            if (string.Compare(useIPAddr, "true", true) == 0)
            {
                return true;
            }
            return false;
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
        static public List<string> GetEnterprisePrintServers()
        {
            List<string> allServers = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EnterpriseServer")).Select(k => ConfigurationManager.AppSettings[k]).OrderBy(x=>x).ToList();
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
        static public List<string> GetTrays()
        {
            List<string> trays = new List<string>(new string[] { "AutoSelect", "Tray1", "Tray2", "Tray3", "Tray4", "Tray5", "Tray6" });
            return (trays);
        }
        static public bool UsePrintTrays()
        {
            string usePrintTrays = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("UsePrintTrays")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(usePrintTrays, "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool ValidatePrinterDNS()
        {
            string validDNS = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("ValidatePrinterDNS")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(validDNS, "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool ShowNumberPrintJobs()
        {
            string validDNS = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("ShowNumberOfJobs")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(validDNS, "true", true) == 0)
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
        static public bool EditEnterprisePrinters()
        {
            string editEntPrint = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("EditEnterprisePrinters")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(editEntPrint, "true", true) == 0)
            {
                return true;
            }
            return false;

        }
        static public bool AdditionalSecurity()
        {
            string theSecurity = ConfigurationManager.AppSettings.AllKeys.Where(k => k.Contains("AdditionalSecurity")).Select(k => ConfigurationManager.AppSettings[k]).FirstOrDefault();
            if (string.Compare(theSecurity, "true", true) == 0)
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
    }
}