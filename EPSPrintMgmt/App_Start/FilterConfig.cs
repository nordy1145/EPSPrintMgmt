using EPSPrintMgmt.Models;
using Hangfire;
using System.Web;
using System.Web.Mvc;

namespace EPSPrintMgmt
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
            GlobalJobFilters.Filters.Add(new ProlongExpirationTimeAttribute());
        }
    }
}
