using SharePoint.Scanning.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePoint.Scanning.DataChangeScanner
{
    public class ChangeDataScanResult: Scan
    {
        public Guid SiteId { get; set; }
        public Guid WebId { get; set; }
        public Guid ListId { get; set; }
        public Guid UniqueId { get; set; }
    }
}
