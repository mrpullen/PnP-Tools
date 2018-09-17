using OfficeDevPnP.Core.Diagnostics;
using SharePoint.Scanning.Common;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePoint.Scanning.SiteFinderScanner
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Retrieve
            var fieldConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).GetSection("ChangeDataFields") as FieldDataSection;
            // Validate commandline options
            var options = new SiteFinderOptions(fieldConfig);
            options.ValidateOptions(args);

            //Instantiate scan job
            SiteFinderScanJob job = new SiteFinderScanJob(options);

            // I'm debugging
            //job.UseThreading = false;
            Log.Info("Data Change Scanner", "Starting");
            job.Execute();
            Log.Info("Data Change Scanner", "Finished");

            // Sample on how to add custom log entry
        }
    }
}
