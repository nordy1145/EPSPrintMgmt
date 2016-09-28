using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Printing;
using System.ComponentModel;


namespace EPSPrintMgmt.Models
{
    public class Printer
    {
        [DisplayName("Printer")]
        public string Name { get; set; }
        [DisplayName("Print Driver")]
        public string Driver { get; set; }
        [DisplayName("Print Server")]
        public string PrintServer { get; set; }
        [DisplayName("Number of Jobs")]
        public int NumberJobs { get; set; }
        [DisplayName("Port Name")]
        public string PortName { get; set; }
        public string Tray { get; set; }
        public bool IsEPS { get; set; }
        public string Status { get; set; }
        public string Comments { get; set; }
        public string Location { get; set; }
    }

    public class AddPrinterClass
    {
        [DisplayName("Driver Name")]
        public string DriverName { get; set; }
        [DisplayName("Port Name")]
        public string PortName { get; set; }
        public string Location { get; set; }
        public string Comment { get; set; }
        [DisplayName("Printer Name")]
        public string DeviceID { get; set; }
        public Boolean Shared { get; set; }
        public Boolean Published { get; set; }
        public Boolean EnableBIDI { get; set; }
        public string Tray { get; set; }
        public bool IsEPS { get; set; }
        public bool IsEnterprise { get; set; }
    }

    public class AddPrinterPortClass
    {
        public string Name { get; set; }
        public bool SNMPEnabled { get; set; }
        public int Protocol { get; set; }
        public int PortNumber { get; set; }
        public string HostAddress { get; set; }
    }
}