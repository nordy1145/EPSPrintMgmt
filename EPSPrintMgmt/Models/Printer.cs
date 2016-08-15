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
        public string Name { get; set; }
        public string Driver { get; set; }
        public string PrintServer { get; set; }
        public int NumberJobs { get; set; }
    }

    public class AddPrinterClass
    {
        public string DriverName { get; set; }
        public string PortName { get; set; }
        public string Location { get; set; }
        public string Comment { get; set; }
        [DisplayName("Printer Name")]
        public string DeviceID { get; set; }
        public Boolean Shared { get; set; }
        public Boolean Published { get; set; }
        public Boolean EnableBIDI { get; set; }

    }

    public class AddPrinterPortClass
    {
        public string Name { get; set; }
        public bool SNMPEnabled { get; set; }
        public int Protocol { get; set; }
        public int PortNumber { get; set; }
        public string HostAddress { get; set; }
    }

    public class PrintTest : System.Printing.PrintQueue
    {
        public PrintTest(PrintServer ps, string Test):base(ps, Test)
        {
            return;
        }
    }

}