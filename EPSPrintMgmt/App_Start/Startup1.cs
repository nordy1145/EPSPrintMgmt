using System;
using System.Threading.Tasks;
using Microsoft.Owin;
using Owin;
using Hangfire;

[assembly: OwinStartup(typeof(EPSPrintMgmt.App_Start.Startup1))]
[assembly: log4net.Config.XmlConfigurator(ConfigFile = "Web.config", Watch = true)]

namespace EPSPrintMgmt.App_Start
{
    public class Startup1
    {
        public void Configuration(IAppBuilder app)
        {
            // For more information on how to configure your application, visit http://go.microsoft.com/fwlink/?LinkID=316888
            GlobalConfiguration.Configuration.UseSqlServerStorage("HangfireDBContext");
            //GlobalJobFilters.Filters.Add(new ProlongExpirationTimeAttribute());
            var options = new BackgroundJobServerOptions { WorkerCount = Environment.ProcessorCount * 5 };
            //app.UseHangfireServer(options);
            app.UseHangfireDashboard();
            app.UseHangfireServer(options);
        }
    }
}
