using OfficeDevPnP.Core.Framework.TimerJobs;
using SharePoint.Scanning.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePoint.Scanning.DataChangeScanner
{
    public class UpdateDataJob: ScanJob
    {
        public UpdateDataJob(UpdateDataOptions options) : base(options as BaseOptions, "UpdateDataJob", "v1.0")
        {
            TimerJobRun += UpdateData_TimerJobRun;
        }

        private void UpdateData_TimerJobRun(object sender, TimerJobRunEventArgs e)
        {
            if(e.WebClientContext == null || e.SiteClientContext == null)
            {
                ScanError error = new ScanError()
                {
                    Error = "No valid ClientContext objects",
                    SiteURL = e.Url,
                    SiteColUrl = e.Url
                };
                this.ScanErrors.Push(error);
                Console.WriteLine("Error on UpdateDataJob for site {1}: {0}", "No valid ClientContext objects", e.Url);

                // bail out
                return;
            }

        }


    }
}
