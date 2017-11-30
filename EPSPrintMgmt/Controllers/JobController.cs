using EPSPrintMgmt.Models;
using Hangfire;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EPSPrintMgmt.Controllers
{
    public class JobController : Controller
    {
        //// GET: Job
        //public ActionResult Index()
        //{
        //    int jobRangeFrom = 0;
        //    int jobReturnCount = 100;

        //    Hangfire.Storage.IMonitoringApi monitor = JobStorage.Current.GetMonitoringApi();
        //    var theList = monitor.SucceededListCount();
        //    Hangfire.Storage.Monitoring.JobList<Hangfire.Storage.Monitoring.SucceededJobDto> theFullLIst = monitor.SucceededJobs(jobRangeFrom, jobReturnCount);
        //    List<PrinterCreation> output = new List<PrinterCreation>();
        //    foreach (var i in theFullLIst)
        //    {
        //        try
        //        {
        //        output.Add(new PrinterCreationReturn(i.Value.Result.ToString(),i.Value.SucceededAt.Value.ToLocalTime()));

        //        }
        //        catch
        //        {
        //            try {
        //                output.Add(new PrinterCreation { result = i.Value.Result.ToString() });
        //            }
        //            catch
        //            {

        //            }
        //            Console.WriteLine("error....");
        //        }
        //    }
        //    return View(theFullLIst);
        //}
        //public ActionResult MorePosts(int? start,int? count)
        //{
        //    int jobRangeFrom = 0;
        //    int jobReturnCount = 100;
        //    Hangfire.Storage.IMonitoringApi monitor = JobStorage.Current.GetMonitoringApi();
        //    var theList = monitor.SucceededListCount();
        //    Hangfire.Storage.Monitoring.JobList<Hangfire.Storage.Monitoring.SucceededJobDto> theFullLIst = monitor.SucceededJobs(jobRangeFrom, jobReturnCount);
        //    return PartialView(theFullLIst);
        //}
        //public ActionResult Test(int? page)
        //{
        //    //var dummyItems = Enumerable.Range(1, 150).Select(x => "Item " + x);

        //    Hangfire.Storage.IMonitoringApi monitor = JobStorage.Current.GetMonitoringApi();
        //    var allCounts = monitor.GetStatistics();
        //    var queuedJobs = monitor.EnqueuedCount("Queued");
        //    var failedJobs = monitor.EnqueuedCount("Failed");
        //    var processingJobs = monitor.EnqueuedCount("Processing");
        //    var enqueuedJobs = monitor.EnqueuedCount("Enqueued");
        //    var succeededJobs = monitor.EnqueuedCount("Succeeded");
        //    //var queuedJobs = monitor.EnqueuedCount("Queued");
        //    var failedallJobs = monitor.FailedCount();
        //    int totalCount = Convert.ToInt32( monitor.SucceededListCount());
        //    var pager = new Pager(totalCount, page);
        //    Hangfire.Storage.Monitoring.JobList<Hangfire.Storage.Monitoring.SucceededJobDto> theFullLIst = monitor.SucceededJobs((pager.CurrentPage - 1) * pager.PageSize, pager.PageSize);
        //    List<PrinterCreation> allJobs = new List<PrinterCreation>();
        //    foreach(var i in theFullLIst)
        //    {
        //        try
        //        {
        //            allJobs.Add(new PrinterCreationReturn(i.Value.Result.ToString(),i.Value.SucceededAt.Value.ToLocalTime()));
        //        }
        //        catch
        //        {
        //            try
        //            {
        //                allJobs.Add(new PrinterCreation { result = i.Value.Result.ToString() });
        //            }
        //            catch
        //            {
        //            }
        //        }
        //    }

        //    var viewModel = new JobsSucceededView
        //    {
        //        //Items = dummyItems.Skip((pager.CurrentPage - 1) * pager.PageSize).Take(pager.PageSize),
        //        Pager = pager,
        //        SucceededJobs = allJobs.AsEnumerable()
                
        //    };

        //    return View(viewModel);
        //}
        public ActionResult FailedJobs(int? page)
        {
            Hangfire.Storage.IMonitoringApi monitor = JobStorage.Current.GetMonitoringApi();
            int totalCount = Convert.ToInt32(monitor.FailedCount());
            var pager = new Pager(totalCount, page);
            Hangfire.Storage.Monitoring.JobList<Hangfire.Storage.Monitoring.FailedJobDto> theFullLIst = monitor.FailedJobs((pager.CurrentPage - 1) * pager.PageSize, pager.PageSize);
            var viewModel = new JobsFailedVIew {
                Pager=pager,
                FailedJobs = theFullLIst,
            };
            Response.AddHeader("Refresh", "60");
            return View(viewModel);
        }
        public ActionResult QueuedJobs(int? page)
        {
            Hangfire.Storage.IMonitoringApi monitor = JobStorage.Current.GetMonitoringApi();
            int totalCount = Convert.ToInt32(monitor.EnqueuedCount("default"));
            var pager = new Pager(totalCount, page);
            Hangfire.Storage.Monitoring.JobList<Hangfire.Storage.Monitoring.EnqueuedJobDto> theFullLIst = monitor.EnqueuedJobs("default",(pager.CurrentPage - 1) * pager.PageSize, pager.PageSize);
            var viewModel = new JobsQueuedView
            {
                Pager = pager,
                QueuedJobs = theFullLIst
            };
            return View(viewModel);
        }
        public ActionResult SucceededJobs(int? page)
        {
            Hangfire.Storage.IMonitoringApi monitor = JobStorage.Current.GetMonitoringApi();
            int totalCount = Convert.ToInt32(monitor.SucceededListCount());
            var pager = new Pager(totalCount, page);

            Hangfire.Storage.Monitoring.JobList<Hangfire.Storage.Monitoring.SucceededJobDto> theFullLIst = monitor.SucceededJobs((pager.CurrentPage - 1) * pager.PageSize, pager.PageSize);
            
            List<PrinterCreation> allJobs = new List<PrinterCreation>();
            foreach (var i in theFullLIst)
            {
                try
                {
                    allJobs.Add(new PrinterCreationReturn(i.Value.Result.ToString(),i.Value.SucceededAt.Value.ToLocalTime()));
                }
                catch
                {
                    try
                    {
                        allJobs.Add(new PrinterCreation { result = i.Value.Result.ToString() });
                    }
                    catch
                    {
                    }
                }
            }

            var viewModel = new JobsSucceededView
            {
                Pager = pager,
                SucceededJobs = allJobs
            };
            Response.AddHeader("Refresh", "60");
            return View(viewModel);
        }
        public ActionResult ProcessingJobs(int? page)
        {
            Hangfire.Storage.IMonitoringApi monitor = JobStorage.Current.GetMonitoringApi();
            int totalCount = Convert.ToInt32(monitor.EnqueuedCount("default"));
            var pager = new Pager(totalCount, page);
            Hangfire.Storage.Monitoring.JobList<Hangfire.Storage.Monitoring.ProcessingJobDto> theFullLIst = monitor.ProcessingJobs((pager.CurrentPage - 1) * pager.PageSize, pager.PageSize);
            var viewModel = new JobsProcessingView
            {
                Pager = pager,
                ProcessingJobs = theFullLIst
            };
            Response.AddHeader("Refresh", "60");
            return View(viewModel);
        }
    }
}