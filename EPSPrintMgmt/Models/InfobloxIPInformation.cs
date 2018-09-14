using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPSPrintMgmt.Models
{
    public partial class InfobloxIPInformation
    {
        [JsonProperty("result")]
        public Result[] Result { get; set; }
    }

    public partial class Result
    {
        [JsonProperty("dhcp_client_identifier")]
        public string DhcpClientIdentifier { get; set; }

        [JsonProperty("ip_address")]
        public string IpAddress { get; set; }

        [JsonProperty("is_conflict")]
        public bool IsConflict { get; set; }

        [JsonProperty("lease_state")]
        public string LeaseState { get; set; }

        [JsonProperty("mac_address")]
        public string MacAddress { get; set; }

        [JsonProperty("names")]
        public string[] Names { get; set; }

        [JsonProperty("network")]
        public string Network { get; set; }

        [JsonProperty("network_view")]
        public string NetworkView { get; set; }

        [JsonProperty("objects")]
        public string[] Objects { get; set; }

        [JsonProperty("_ref")]
        public string Ref { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; }

        [JsonProperty("types")]
        public string[] Types { get; set; }

        [JsonProperty("usage")]
        public string[] Usage { get; set; }
    }
}