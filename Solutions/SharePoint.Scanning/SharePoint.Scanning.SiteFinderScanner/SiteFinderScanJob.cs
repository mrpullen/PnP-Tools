using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.TimerJobs;
using SharePoint.Scanning.Common;
using SharePoint.Scanning.Framework;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SharePoint.Scanning.SiteFinderScanner
{
    public class SiteFinderScanJob: ScanJob
    {
        private Int32 SitesToScan = 0;
        public ConcurrentDictionary<string, Scan> SiteScanResults;
        public SiteFinderOptions options;
        public SiteFinderScanJob(SiteFinderOptions options): base(options as BaseOptions, "SiteFinder","1.0")
        {
            this.options = options;
            TimerJobRun += SiteFinderScanJob_TimerJobRun;
        }

        private void SiteFinderScanJob_TimerJobRun(object sender, TimerJobRunEventArgs e)
        {
            // Validate ClientContext objects
            if (e.WebClientContext == null || e.SiteClientContext == null)
            {
                ScanError error = new ScanError()
                {
                    Error = "No valid ClientContext objects",
                    SiteURL = e.Url,
                    SiteColUrl = e.Url
                };
                this.ScanErrors.Push(error);
                Console.WriteLine("Error for site {1}: {0}", "No valid ClientContext objects", e.Url);

                // bail out
                return;
            }

            // thread safe increase of the sites counter
            IncreaseScannedSites();

            try
            {
                Console.WriteLine("Processing site {0}...", e.Url);

                #region Basic sample
               // Set the first site collection done flag + perform telemetry
                SetFirstSiteCollectionDone(e.WebClientContext);
                // add your custom scan logic here, ensure the catch errors as we don't want to terminate scanning
                e.WebClientContext.Load(e.WebClientContext.Web, p => p.Title);
                e.WebClientContext.Load(e.WebClientContext.Web.Fields, flds => flds.Include<Field>(field => field.Title, Field=>Field.Id, field=>field.InternalName, field=>field.StaticName));
                e.WebClientContext.ExecuteQueryRetry();
                var fields = options.FieldConfig.Fields.GetFieldNameValues();

                // Now if we find any of the FieldNAmes we are looking to set, we shoudl scan this site.
                if (e.WebClientContext.Web.Fields.Any(f => fields.Contains(f.Title))) {
                    Scan result = new Scan()
                    {
                        SiteColUrl = e.Url,
                        SiteURL = e.Url
                    };


                    // Store the scan result
                    if (!SiteScanResults.TryAdd(e.Url, result))
                    {
                        ScanError error = new ScanError()
                        {
                            SiteURL = e.Url,
                            SiteColUrl = e.Url,
                            Error = "Could not add scan result for this site"
                        };
                        this.ScanErrors.Push(error);
                    }
                }
                #endregion

            }
            catch (Exception ex)
            {
                ScanError error = new ScanError()
                {
                    Error = ex.Message,
                    SiteColUrl = e.Url,
                    SiteURL = e.Url,
                    Field1 = "put additional info here"
                };
                this.ScanErrors.Push(error);
                Console.WriteLine("Error for site {1}: {0}", ex.Message, e.Url);
            }

            // Output the scanning progress
            try
            {
                TimeSpan ts = DateTime.Now.Subtract(this.StartTime);
                Console.WriteLine($"Thread: {Thread.CurrentThread.ManagedThreadId}. Processed {this.ScannedSites} of {this.SitesToScan} site collections ({Math.Round(((float)this.ScannedSites / (float)this.SitesToScan) * 100)}%). Process running for {ts.Days} days, {ts.Hours} hours, {ts.Minutes} minutes and {ts.Seconds} seconds.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error showing progress: {ex.ToString()}");
            }

        }

        public override List<string> ResolveAddedSites(List<string> addedSites)
        {
            // Use search approach to determine which sites to process
           List<string> sites = new List<string>(100000);

            string tenantAdmin = "";
            if (!string.IsNullOrEmpty(this.TenantAdminSite))
            {
                tenantAdmin = this.TenantAdminSite;
            }
            this.Realm = GetRealmFromTargetUrl(new Uri(tenantAdmin));

            //Enumerate all sites. 
            SPOSitePropertiesEnumerable spp = null;
            using (ClientContext ccAdmin = this.CreateClientContext(tenantAdmin))
            {
                Tenant tenant = new Tenant(ccAdmin);
                int startIndex = 0;

                string site = String.Empty;
                while (spp == null || spp.Count > 0)
                {
                    spp = tenant.GetSiteProperties(startIndex, true);
                    ccAdmin.Load(spp);
                    ccAdmin.ExecuteQuery();
                    
                    foreach (SiteProperties sp in spp)
                    {
                        sites.Add(sp.Url);
                    }
                    startIndex++;
                }
            }

                return sites;
        }
    }
}
