using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPMigrationTool
{
    class MigrationSite
    {
        //properties
        public string OldSiteUrl { get; set; }
        public string OldSiteName { get; set; }
        public string NewSiteUrl { get; set; }
        public string NewSiteName { get; set; }
        public string siteRelativeUrl { get; set; }

        //constructor
        public MigrationSite(string OldSiteUrl, string OldSiteName, string siteRelativeUrl, string NewSiteUrl, string NewSiteName)
        {
            this.OldSiteUrl = OldSiteUrl;
            this.OldSiteName = OldSiteName;
            this.NewSiteUrl = NewSiteUrl;
            this.NewSiteName = NewSiteName;
            this.siteRelativeUrl = siteRelativeUrl;
        }
    }
}
