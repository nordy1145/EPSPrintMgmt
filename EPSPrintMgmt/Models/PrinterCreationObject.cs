using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPSPrintMgmt.Models
{
    public class PrinterCreationObject
    {
        public AddPrinterClass printer { get; set; }
        public string server { get; set; }
        public string user { get; set; }
    }
    public class PrinterCreationObjectReturn : PrinterCreationObject
    {
        public PrinterCreationObjectReturn(string json)
        {
            JObject jObject = JObject.Parse(json);
            printer = jObject["theNewPrinter"];
            server = (string)jObject["thePrintServer"];
            user = (string)jObject["user"];
        }

    }
}