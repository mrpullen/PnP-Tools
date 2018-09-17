using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.UserProfiles;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.TimerJobs;
using SharePoint.Scanning.Common;
using SharePoint.Scanning.Framework;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SharePoint.Scanning.DataChangeScanner
{
    public class ChangeDataScanJob : ScanJob
    {
        #region Variables
        //public Int32 UserProfileLock = 0;
        //public Int32 TaxonomySessionLock = 0;
        internal List<Mode> ScanModes;
        internal bool UseSearchQuery = false;
        private Int32 SitesToScan = 0;
        internal FieldDataSection fieldConfig;
        public ConcurrentDictionary<string, ChangeDataScanResult> ScanResults;
        public ConcurrentDictionary<string, Dictionary<string, string>> UserProfileData;
        public ConcurrentDictionary<string, Dictionary<string, TaxonomyFieldValue>> TermSetLookup;
        #endregion

        public ChangeDataScanJob(ChangeDataOptions options) : base(options as BaseOptions, "ChangeDataScanJob", "1.0")
        {
            // Configure job specific settings
            ScanModes = options.ScanModes;
            UseSearchQuery = !options.DontUseSearchQuery;
            ExpandSubSites = false; // false is default value, shown her for demo purposes
            ScanResults = new ConcurrentDictionary<string, ChangeDataScanResult>(options.Threads, 10000);
            UserProfileData = new ConcurrentDictionary<string, Dictionary<string, string>>(options.Threads, 20000);
            TermSetLookup = new ConcurrentDictionary<string, Dictionary<string, TaxonomyFieldValue>>(options.Threads, 2000);
            fieldConfig = options.FieldConfig; //ConfigurationManager.GetSection("ChangeDataFields") as FieldDataSection;

            LoadTaxonomyTerms();

             // Connect the eventhandler
             TimerJobRun += ChangeDataScanResults_TimerJobRun;
        }

        public ChangeQuery GetChangeQuery(ChangeToken tokenStart)
        {
            var changeQuery = new ChangeQuery(false, false);
            if (tokenStart != null)
            {
                changeQuery.ChangeTokenStart = tokenStart;
            }
            //TODO: Update to support this pulling the confguration of the query from the app.config.
            changeQuery.Item = true;
            changeQuery.Add = true;

            return changeQuery;
        }

        public Dictionary<string, string> GetProfileData(ClientContext context, string userName)
        {
            if (this.UserProfileData.ContainsKey(userName))
            {
                return this.UserProfileData[userName];
            }
            else
            {
                PeopleManager peopleManager = new PeopleManager(context);
                PersonProperties personProperties = peopleManager.GetPropertiesFor(userName);
                context.Load(personProperties, p => p.AccountName, p => p.UserProfileProperties);
                context.ExecuteQueryRetry();
                var dictionary = new Dictionary<string, string>();
                var profileProperties = fieldConfig.Fields.GetProfilePropertyValues();

                foreach (var profileProperty in profileProperties)
                {
                    if (personProperties.UserProfileProperties.ContainsKey(profileProperty))
                    {
                        var profilePropertyValue = personProperties.UserProfileProperties[profileProperty];
                        if (!String.IsNullOrEmpty(profilePropertyValue))
                        {
                            dictionary.Add(profileProperty, profilePropertyValue);
                        }
                    }
                }
                if (this.UserProfileData.TryAdd(userName, dictionary))
                {
                    return UserProfileData[userName];
                }
                else
                {
                    Log.Debug("ChangeDataScanJob -- GetProfileData", "Unable to add user profile information to UserPRofile Data ConcurrentDictionary", null);
                    return dictionary;

                }
            }
        }

        public void LoadTaxonomyTerms()
        {
            string tenantAdmin = "";
            if (!string.IsNullOrEmpty(this.TenantAdminSite))
            {
                tenantAdmin = this.TenantAdminSite;
            }
            this.Realm = GetRealmFromTargetUrl(new Uri(tenantAdmin));


            using (ClientContext ccAdmin = this.CreateClientContext(tenantAdmin))
            {
                var termGroup = fieldConfig.TermGroup;
                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ccAdmin);
                var taxTermStore = taxonomySession.GetDefaultKeywordsTermStore();
                var taxTermGroup = taxTermStore.GetTermGroupByName(termGroup);
                var taxTermSets = taxTermGroup.TermSets;
                ccAdmin.Load(taxTermGroup);
                ccAdmin.Load(taxTermSets, t => t.Include(tSets => tSets.Terms, tSets=>tSets.Name));
                ccAdmin.ExecuteQueryRetry();
                
                foreach(var taxTermSet in taxTermSets)
                {
                    var dictionary = new Dictionary<string, TaxonomyFieldValue>();
                    var terms = taxTermSet.Terms;
                    foreach (var term in terms )
                    {
                        
                        var value = new TaxonomyFieldValue();
                        value.Label = term.Name;
                        value.TermGuid = term.Id.ToString();
                        value.WssId = -1;
                        dictionary.Add(term.Name, value);
                    }

                    if (TermSetLookup.TryAdd(taxTermSet.Name, dictionary))
                    {
                       
                    }
                    else
                    {
                        Log.Error("ChangeDataScanJob -- Failed to load LoadTaxonomyTerms", "Failed to load dictionary for terms", null);
                    }
                }
                 
                }

        }


        
        public TaxonomyFieldValue GetTaxonomyFieldValue(string termSetName, string termName)
        {
            if(TermSetLookup.ContainsKey(termSetName))
            {
                var dictionary = this.TermSetLookup[termSetName];
                if(dictionary.ContainsKey(termName))
                {
                    var result = dictionary[termName];// as 
                    return result;
                }
            }
            return null;

        }
  
        #region Scanner implementation
        /// <summary>
        /// Grab the number of sites that need to be scanned...will be needed to show progress when we're having a long run
        /// </summary>
        /// <param name="addedSites"></param>
        /// <returns></returns>
        public override List<string> ResolveAddedSites(List<string> addedSites)
        {
            if (!this.UseSearchQuery)
            {
                var sites = base.ResolveAddedSites(addedSites);
                this.SitesToScan = sites.Count;
                return sites;
            }
            else
            {
                // Use search approach to determine which sites to process
                List<string> searchedSites = new List<string>(100);

                string tenantAdmin = "";
                if (!string.IsNullOrEmpty(this.TenantAdminSite))
                {
                    tenantAdmin = this.TenantAdminSite;
                }
                else
                {
                    if (string.IsNullOrEmpty(this.Tenant))
                    {
                        this.Tenant = new Uri(addedSites[0]).DnsSafeHost.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries)[0];
                    }

                    tenantAdmin = $"https://{this.Tenant}-admin.sharepoint.com";
                }

                this.Realm = GetRealmFromTargetUrl(new Uri(tenantAdmin));


                using (ClientContext ccAdmin = this.CreateClientContext(tenantAdmin))
                {
                    List<string> propertiesToRetrieve = new List<string>
                    {
                        "Title",
                        "SPSiteUrl",
                        "FileExtension",
                        "OriginalPath"
                    };

                    // Get sites with a given web template
                    //var results = Search(ccAdmin.Web, "contentclass:STS_Web (WebTemplate:ACCSVC OR WebTemplate:ACCSRV)", propertiesToRetrieve);
                    // Get sites that contain a certain set of files
                    var results = this.Search(ccAdmin.Web, "((fileextension=htm OR fileextension=html) AND contentclass=STS_ListItem_DocumentLibrary)", propertiesToRetrieve);
                    foreach (var site in results)
                    {
                        if (!string.IsNullOrEmpty(site["SPSiteUrl"]) && !searchedSites.Contains(site["SPSiteUrl"]))
                        {
                            searchedSites.Add(site["SPSiteUrl"]);
                        }
                    }
                }

                this.SitesToScan = searchedSites.Count;
                return searchedSites;
            }
        }

        private void ChangeDataScanResults_TimerJobRun(object sender, TimerJobRunEventArgs e)
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

                //Simply return every list item from the entire site collection.. that has required fields.
                if (this.ScanModes.Contains(Mode.SiteFullSync))
                {

                    #region Sub site iteration sample                
                    // Set the first site collection done flag + perform telemetry
                    SetFirstSiteCollectionDone(e.WebClientContext);

                    // Manually iterate over the content
                    IEnumerable<string> expandedSites = GetAllSubSites(e.SiteClientContext.Site);
                    bool isFirstSiteInList = true;
                    string siteCollectionUrl = "";

                    foreach (string site in expandedSites)
                    {
                        // thread safe increase of the webs counter
                        IncreaseScannedWebs();

                        // Clone the existing ClientContext for the sub web
                        using (ClientContext ccWeb = e.SiteClientContext.Clone(site))
                        {
                            Console.WriteLine("Processing site {0}...", site);

                            if (isFirstSiteInList)
                            {
                                // Perf optimization: do one call per site to load all the needed properties
                                var spSite = (ccWeb as ClientContext).Site;
                                ccWeb.Load(spSite, p => p.RootWeb, p => p.Url);
                                ccWeb.Load(spSite.RootWeb, p => p.Id);

                                isFirstSiteInList = false;
                            }

                            // Perf optimization: do one call per web to load all the needed properties
                            ccWeb.Load(ccWeb.Web, p => p.Id, p => p.Title);
                            ccWeb.Load(ccWeb.Web, p => p.WebTemplate, p => p.Configuration);
                            ccWeb.Load(ccWeb.Web, p => p.Lists.Include(li=>li.Id, li => li.Fields, li => li.UserCustomActions, li => li.Title, li => li.Hidden, li => li.DefaultViewUrl, li => li.BaseTemplate, li => li.RootFolder, li => li.ListExperienceOptions));
                            ccWeb.ExecuteQueryRetry();

                            // Fill site collection url
                            if (string.IsNullOrEmpty(siteCollectionUrl))
                            {
                                siteCollectionUrl = ccWeb.Site.Url;
                            }

                            // Need to know if this is a sub site?
                            if (ccWeb.Web.IsSubSite())
                            {
                                // Sub site specific logic
                            }

                            ChangeDataScanResult result = new ChangeDataScanResult()
                            {
                                SiteColUrl = e.Url,
                                SiteURL = site,
                                SiteId = Guid.Empty,
                                WebId = Guid.Empty,
                                ListId = Guid.Empty,
                                UniqueId = Guid.Empty
                            };

                            // Store the scan result
                            if (!ScanResults.TryAdd(site, result))
                            {
                                ScanError error = new ScanError()
                                {
                                    SiteURL = site,
                                    SiteColUrl = e.Url,
                                    Error = "Could not add scan result for this web"
                                };
                                this.ScanErrors.Push(error);
                            }
                        }
                    }
                    #endregion


                }
                //only return list items that are new based on change token from entire site collection. easy peasy.
                if (this.ScanModes.Contains(Mode.SiteDeltaSync))
                {
                    try
                    {
                        #region Site Delta Sync
                        var siteContext = e.SiteClientContext;
                    

                        var site = e.SiteClientContext.Site;
                        var web = site.RootWeb;
                        var lists = web.Lists;
                   
                        siteContext.Load(site, s => s.Id, s => s.CurrentChangeToken);
                        siteContext.Load(web, p=>p.Id, p=>p.AllProperties.FieldValues, p=>p.Title, p=>p.Lists.Include(l => l.CurrentChangeToken, l => l.Fields, l => l.Id,l=>l.Title));
                    
                  
                        siteContext.ExecuteQueryRetry();
                        ChangeToken changeToken = null;
                        if (web.AllProperties.FieldValues.Keys.Contains("ChangeDataToken")) {
                            changeToken = new ChangeToken() { StringValue = web.AllProperties["ChangeDataToken"].ToString() };
                        }
                   
                        var changeQuery = this.GetChangeQuery(changeToken);
                        if(site.CurrentChangeToken.GreaterThan(changeToken)) {
                        

                        }
                    
                    // TODO: Check if the site changeToken is greater than the changeToken we've pulled from Site properties.
                    var siteChanges = siteContext.GetAllSiteChanges(changeQuery);
                    var fieldMaps = fieldConfig.Fields.GetFieldNameValues();
                  
                        foreach (ChangeItem siteChange in siteChanges)
                        {
                            ChangeDataScanResult result = new ChangeDataScanResult()
                            {
                                SiteColUrl = e.Url,
                                SiteURL = web.Url,
                                SiteId = siteChange.SiteId,
                                WebId = siteChange.WebId,
                                ListId = siteChange.ListId,
                                UniqueId = siteChange.UniqueId
                            };

                            // Store the scan result
                            if (!ScanResults.TryAdd(e.Url, result))
                            {
                                ScanError error = new ScanError()
                                {
                                    SiteURL = e.Url,
                                    SiteColUrl = e.Url,
                                    Error = "Could not add scan result for this web"
                                };
                                this.ScanErrors.Push(error);
                            }
                        }

                        //Set Site Token
                        web.AllProperties["ChangeDataToken"] = site.CurrentChangeToken.StringValue;
                        web.Update();
                    }
                    catch(Exception exc)
                    {
                        ScanError error = new ScanError()
                        {
                            Error = exc.Message,
                            SiteColUrl = e.Url,
                            SiteURL = e.Url,
                            Field1 = "Error in processing changes in SiteDeltaSync"
                        };
                        this.ScanErrors.Push(error);
                        Console.WriteLine("Error for site {1}: {0}", exc.Message, e.Url);
                        //Log.Error(exc, "ChangeDataScanJob-ChangeItem", exc.Message, null);
                    }

                    #endregion
                }

                #region Basic sample
                /*
                // Set the first site collection done flag + perform telemetry
                SetFirstSiteCollectionDone(e.WebClientContext);

                // add your custom scan logic here, ensure the catch errors as we don't want to terminate scanning
                e.WebClientContext.Load(e.WebClientContext.Web, p => p.Title);
                e.WebClientContext.ExecuteQueryRetry();
                ScanResult result = new ScanResult()
                {
                    SiteColUrl = e.Url,
                    SiteURL = e.Url,
                    SiteTitle = e.WebClientContext.Web.Title
                };

                // Store the scan result
                if (!ScanResults.TryAdd(e.Url, result))
                {
                    ScanError error = new ScanError()
                    {
                        SiteURL = e.Url,
                        SiteColUrl = e.Url,
                        Error = "Could not add scan result for this site"
                    };
                    this.ScanErrors.Push(error);
                }
                */
                #endregion

                #region Search based sample
                /**
                // Set the first site collection done flag + perform telemetry
                SetFirstSiteCollectionDone(e.WebClientContext);

                // Need to use search inside this site collection?
                List<string> propertiesToRetrieve = new List<string>
                {
                    "Title",
                    "SPSiteUrl",
                    "FileExtension",
                    "OriginalPath"
                };
                var searchResults = this.Search(e.SiteClientContext.Web, $"((fileextension=htm OR fileextension=html) AND contentclass=STS_ListItem_DocumentLibrary AND Path:{e.Url.TrimEnd('/')}/*)", propertiesToRetrieve);
                foreach (var searchResult in searchResults)
                {

                    ScanResult result = new ScanResult()
                    {
                        SiteColUrl = e.Url,
                        FileName = searchResult["OriginalPath"]
                    };

                    // Get web url
                    var webUrlData = Web.GetWebUrlFromPageUrl(e.SiteClientContext, result.FileName);
                    e.SiteClientContext.ExecuteQueryRetry();
                    result.SiteURL = webUrlData.Value;

                    // Store the scan result, use FileName as unique key in this sample
                    if (!ScanResults.TryAdd(result.FileName, result))
                    {
                        ScanError error = new ScanError()
                        {
                            SiteURL = e.Url,
                            SiteColUrl = e.Url,
                            Error = "Could not add scan result for this web"
                        };
                        this.ScanErrors.Push(error);
                    }
                }
                **/
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

        public override DateTime Execute()
        {
            // Triggers the run of the scanning...will result in ReferenceScanJob_TimerJobRun being called per site collection or per site
            var start = base.Execute();

            // Handle the export of the job specific scanning data
            string outputfile = string.Format("{0}\\ChangeDataScanJobResults.csv", this.OutputFolder);
            string[] outputHeaders = new string[] { "Site Collection Url", "Site Url", "Title", "File name" };
            Console.WriteLine("Outputting change data scan results to {0}", outputfile);

            using (StreamWriter outfile = new StreamWriter(outputfile))
            {
                outfile.Write(string.Format("{0}\r\n", string.Join(this.Separator, outputHeaders)));
                foreach (var item in this.ScanResults)
                {
                    outfile.Write(string.Format("{0}\r\n", string.Join(this.Separator, ToCsv(item.Value.SiteColUrl), ToCsv(item.Value.SiteURL), ToCsv(item.Value.ListId.ToString()), ToCsv(item.Value.UniqueId.ToString()))));
                }
            }

            Console.WriteLine("=====================================================");
            Console.WriteLine("All done. Took {0} for {1} sites", (DateTime.Now - start).ToString(), this.ScannedSites);
            Console.WriteLine("=====================================================");

            /**
             *  Call UpdateDataJob with all the scan results. This will take care of updating the data. 
             *  Or we could call this in another step and pass the parameter to the output file. 
             *  Then we can scan without updating and update without scanning.
             * */


            return start;
        }
        #endregion
    }
}
