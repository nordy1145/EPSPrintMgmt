using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPSPrintMgmt.Models
{
    public class PrinterCreation
    {

        public string userName { get; set; }
        public string result { get; set; }
        public string processingTime { get; set; }
        public string printer { get; set; }
        public string server { get; set; }
        public string comment { get; set; }
        public DateTime? startTime { get; set; }
    }
    public class PrinterCreationReturn : PrinterCreation
    {
        public PrinterCreationReturn(string json, DateTime? theStartTime)
        {
            JObject jObject = JObject.Parse(json);
            userName = (string)jObject["userName"];
            result = (string)jObject["result"];
            processingTime = (string)jObject["processingTime"];
            printer = (string)jObject["printer"];
            server = (string)jObject["server"];
            comment = (string)jObject["comment"];
            startTime = theStartTime;
        }

    }
}