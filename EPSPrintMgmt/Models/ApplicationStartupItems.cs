using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EPSPrintMgmt.Models
{
    public class ApplicationStartupItems
    {
    }
    public class UserProfilePictureActionFilter : ActionFilterAttribute
    {

        public override void OnResultExecuting(ResultExecutingContext filterContext)
        {
            filterContext.Controller.ViewBag.IsAbleAddEPSandENTPrinters = Support.AddEPSAndEnterprisePrinters();
            filterContext.Controller.ViewBag.IsAbleAddENTPrinters = Support.AddEnterprisePrinters();
            
            //filterContext.Controller.ViewBag.IsAdmin = MembershipService.IsAdmin;

            //var userProfile = MembershipService.GetCurrentUserProfile();
            //if (userProfile != null)
            //{
            //    filterContext.Controller.ViewBag.Avatar = userProfile.Picture;
            //}
        }

    }
}