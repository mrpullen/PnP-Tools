using CommandLine;
using SharePoint.Scanning.Common;
using SharePoint.Scanning.Framework;
using System;
using System.Collections.Generic;

namespace SharePoint.Scanning.DataChangeScanner
{

    /// <summary>
    /// Mode
    /// FullSync get all data from each list - return the data if the item will support the data update.
    /// DeltaSync gets only new data from each list - using Change data.
    /// </summary>
    public enum Mode
    {
        //Sync Sites
        SiteFullSync = 0,
        SiteDeltaSync = 1,
        //Sync Lists 
        ListFullSync = 2,
        ListDeltaSync = 3,
        //Sync Items by Search
        SearchFullSync = 4,
        SearchDeltaSync = 5
    }

    public class ChangeDataOptions : BaseOptions
    {
        public FieldDataSection FieldConfig { get; set; }

        public ChangeDataOptions(FieldDataSection fieldConfig)
        {
            this.FieldConfig = fieldConfig;
        }

        public List<Mode> ScanModes { get; set; }
        
        [Option('m', "mode", HelpText = "Execution mode. Use following modes: SiteFullSync, SiteDeltaSync, ListFullSync, ListDeltaSync. Omit or use scan for a full scan", DefaultValue = Mode.SiteFullSync, Required = true)]
        public Mode Mode { get; set; }

        [Option('q', "dontusesearchquery", HelpText = "Use site enumeration instead of search to find the impacted files", DefaultValue = false)]
        public bool DontUseSearchQuery { get; set; }

        [Option('l', "lists", HelpText = "Execution mode. Use following lists: Documents, Posts. Omit will cause no work to be done on ListFullSync / ListDeltaSync modes", DefaultValue = "", Required = false)]
        public List<string> Lists { get; set; }

        public override void ValidateOptions(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine(this.GetUsage());
                Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);

            }
            base.ValidateOptions(args);

            ScanModes = this.GetAllOptions(args);

            if (ScanModes.Count == 0)
            {
                Console.WriteLine("You need at least one Mode set.");
                Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);

            }
        }


        [HelpOption]
        public string GetUsage()
        {
            var help = this.GetUsage("Data Change Reference Scanner");
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("==========================================================");
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("See the PnP-Tools repo for more information at:");
            help.AddPreOptionsLine("https://github.com/SharePoint/PnP-Tools/tree/master/Solutions/SharePoint.Scanning");
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("Specifying your urls to scan + url to tenant admin (needed for SPO Dedicated):");
            help.AddPreOptionsLine("==============================================================================");
            help.AddPreOptionsLine("Using app-only:");
            help.AddPreOptionsLine("referencescanner.exe -m <mode> -r <urls> -a <tenant admin site> -i <your client id> -s <your client secret>");
            help.AddPreOptionsLine("e.g. referencescanner.exe -m scan -r https://team.contoso.com/*,https://mysites.contoso.com/* -a https://contoso-admin.contoso.com -i 7a5c1615-997a-4059-a784-db2245ec7cc1 -s eOb6h+s805O/V3DOpd0dalec33Q6ShrHlSKkSra1FFw=");
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("Using credentials:");
            help.AddPreOptionsLine("referencescanner.exe -m <mode> -r <urls> -a <tenant admin site> -u <your user id> -p <your user password>");
            help.AddPreOptionsLine("e.g. referencescanner.exe -m scan -r https://team.contoso.com/*,https://mysites.contoso.com/* -a https://contoso-admin.contoso.com -u spadmin@contoso.com -p pwd");

            help.AddOptions(this);
            return help;
        }

        public List<Mode> GetAllOptions(string[] args)
        {
            System.Collections.Generic.List<Mode> lstmode = new System.Collections.Generic.List<Mode>();
            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i].ToLower();

                if (arg == "-m" || arg == "--mode")
                {
                    string mode = args[i + 1];
                    if (mode != null)
                    {
                        string[] modes = mode.Split(',');
                        foreach (string m in modes)
                        {
                            switch (m.Trim().ToLower())
                            {
                                case "fullsync":
                                    if (!lstmode.Contains(Mode.SiteFullSync)) { lstmode.Add(Mode.SiteFullSync); }
                                    break;
                                case "deltasync":
                                    if (!lstmode.Contains(Mode.SiteDeltaSync)) lstmode.Add(Mode.SiteDeltaSync);
                                    break;
                            }
                        }

                        return lstmode;
                    }
                }
            }

            

            return lstmode;
        }
    }
}
