using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Hangfire;
using Hangfire.Common;
using Hangfire.States;
using Hangfire.Storage;

namespace EPSPrintMgmt.Models
{
    public class ProlongExpirationTimeAttribute : JobFilterAttribute, IApplyStateFilter
    {
        public void OnStateUnapplied(ApplyStateContext context, IWriteOnlyTransaction transaction)
        {
            context.JobExpirationTimeout = TimeSpan.FromDays(30);
        }

        public void OnStateApplied(ApplyStateContext context, IWriteOnlyTransaction transaction)
        {
            context.JobExpirationTimeout = TimeSpan.FromDays(30);
        }
    }
}