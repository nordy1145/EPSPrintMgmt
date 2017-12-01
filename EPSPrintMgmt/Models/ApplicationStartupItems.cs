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
            bool canAddEntAndEps;
            if (Support.AddEPSAndEnterprisePrinters() == true)
            {
                if (Support.AdditionalSecurity())
                {
                    if (Support.IsUserAuthorized(Support.ADGroupCanAddEPSAndEnterprisePrinter()))
                    {
                        canAddEntAndEps = true;
                    }
                    else
                    {
                        canAddEntAndEps = false;
                    }
                }
                else
                {
                    canAddEntAndEps = true;
                }
            }
            else
            {
                canAddEntAndEps = false;
            }
            //filterContext.Controller.ViewBag.IsAbleAddEPSandENTPrinters = (Support.IsUserAuthorized(Support.ADGroupCanAddEPSAndEnterprisePrinter()));
            filterContext.Controller.ViewBag.IsAbleAddEPSandENTPrinters = (canAddEntAndEps);
            //filterContext.Controller.ViewBag.IsAbleAddEPSandENTPrinters = Support.AddEPSAndEnterprisePrinters();
            bool canAddENT;
            if (Support.AddEnterprisePrinters() == true)
            {
                if (Support.AdditionalSecurity())
                {
                    if (Support.IsUserAuthorized(Support.ADGroupCanAddEnterprisePrinter()))
                        {
                        canAddENT = true;
                    }
                    else
                    {
                        canAddENT = false;
                    }
                }
                else
                {
                    canAddENT = true;
                }
            }else
            {
                canAddENT = false;
            }
            //filterContext.Controller.ViewBag.IsAbleAddENTPrinters = (Support.IsUserAuthorized(Support.ADGroupCanAddEnterprisePrinter())&&Support.AddEnterprisePrinters());
            filterContext.Controller.ViewBag.IsAbleAddENTPrinters = (canAddENT);

            //filterContext.Controller.ViewBag.IsAbleAddENTPrinters = Support.AddEnterprisePrinters();
            filterContext.Controller.ViewBag.IsAbleAddEPSPrinters = (Support.IsUserAuthorized(Support.ADGroupCanAddEPSPrinter()));

            filterContext.Controller.ViewBag.IsAbleSeePrintJobs = (Support.IsUserAuthorized(Support.ADGroupCanPurgePrintQueues()));

            filterContext.Controller.ViewBag.IsAbleSeePrintServers = (Support.IsUserAuthorized(Support.ADGroupCanViewPrintServers()));


            //filterContext.Controller.ViewBag.IsAdmin = MembershipService.IsAdmin;

            //var userProfile = MembershipService.GetCurrentUserProfile();
            //if (userProfile != null)
            //{
            //    filterContext.Controller.ViewBag.Avatar = userProfile.Picture;
            //}
        }

    }
}