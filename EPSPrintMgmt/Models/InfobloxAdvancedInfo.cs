using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPSPrintMgmt.Models
{
    public class InfobloxAdvancedInfo
    {
        [JsonProperty("result")]
        public IList<InfobloxAdvancedInfoResult> Result { get; set; }
    }
    public class InfobloxAdvancedInfoResult
    {

        [JsonProperty("_ref")]
        public string Ref { get; set; }

        [JsonProperty("ipv4addr")]
        public string Ipv4addr { get; set; }

        [JsonProperty("network_view")]
        public string NetworkView { get; set; }

        [JsonProperty("address")]
        public string Address { get; set; }
    }

}