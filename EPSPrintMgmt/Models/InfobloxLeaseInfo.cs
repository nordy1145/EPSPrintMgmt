using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPSPrintMgmt.Models
{
    public class InfobloxLeaseInfo
    {
        [JsonProperty("result")]
        public IList<InfobloxLeaseInfoResult> Result { get; set; }
    }
    public class InfobloxLeaseInfoResult
    {

        [JsonProperty("_ref")]
        public string Ref { get; set; }

        [JsonProperty("binding_state")]
        public string BindingState { get; set; }

        [JsonProperty("client_hostname")]
        public string ClientHostname { get; set; }

        [JsonProperty("fingerprint")]
        public string Fingerprint { get; set; }

        [JsonProperty("hardware")]
        public string Hardware { get; set; }
    }
}