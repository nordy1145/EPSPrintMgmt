using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel;

namespace EPSPrintMgmt.Models
{
    public class MyPrintServer
    {
        public string Name { get; set; }
        public string IP { get; set; }
        [DisplayName("Printer Count")]
        public string PrinterCount { get; set; }
    }
}