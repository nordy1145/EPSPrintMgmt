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
    public class ClonePrinter:Printer
    {
        public string sourcePrintQueue { get; set; }
        public string sourcePrintServer { get; set; }
    }
    public class EditPrinter:Printer
    {
        public string OldDriver { get; set; }
        public string OldPortName { get; set; }
        public string OldTry { get; set; }
        public string OldComments { get; set; }
        public string OldLocation { get; set; }
        public string OldPrintServer { get; set; }
    }
    public class AddPrinterClass
    {
        [DisplayName("Driver Name")]
        public string DriverName { get; set; }
        [DisplayName("DNS or IP Address")]
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
        [DisplayName("Print Server")]
        public string PrintServer { get; set; }
        [DisplayName("Source Printer")]
        public string SourcePrinter { get; set; }
        public string SourceServer { get; set; }
        [DisplayName("Copy Printer Device Settings")]
        public bool cloneDevSettings { get; set; }
    }
    public class AddEPSandENTPrinterClass : AddPrinterClass
    {
        [DisplayName("Enterprise Driver Name")]
        public string ENTDriverName { get; set; }
        [DisplayName("DNS or IP Address")]
        public string ENTPortName { get; set; }
        [DisplayName("Enterprise Source Printer")]
        public string ENTSourcePrinter { get; set; }
        public string ENTSourceServer { get; set; }
        public Boolean ENTShared { get; set; }
        public Boolean ENTPublished { get; set; }
        public Boolean ENTEnableBIDI { get; set; }
        public string ENTTray { get; set; }
        [DisplayName("Copy Enterprise Printer Device Settings")]
        public bool cloneENTDevSettings { get; set; }
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