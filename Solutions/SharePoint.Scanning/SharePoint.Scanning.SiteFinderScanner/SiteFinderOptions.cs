using SharePoint.Scanning.Common;
using SharePoint.Scanning.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePoint.Scanning.SiteFinderScanner
{
    public class SiteFinderOptions : BaseOptions
    {
        public FieldDataSection FieldConfig { get; set; }
        public SiteFinderOptions(FieldDataSection fieldConfig)
        {
            FieldConfig = fieldConfig;
        }


    }
}
