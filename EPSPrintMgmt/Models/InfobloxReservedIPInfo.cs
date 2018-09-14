using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPSPrintMgmt.Models
{
    public class InfobloxReservedIPInfo
    {
        [JsonProperty("ipv4addr")]
        public string Ipv4addr { get; set; }

        [JsonProperty("mac")]
        public string Mac { get; set; }

        [JsonProperty("match_client")]
        public string MatchClient { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }
        [JsonProperty("comment")]
        public string Comment { get; set; }
    }
}