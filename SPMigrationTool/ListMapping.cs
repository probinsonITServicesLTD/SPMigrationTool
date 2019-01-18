using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPMigrationTool
{
    class ListMapping
    {
        public string ListName { get; set; }
        public string ListType { get; set; }
        public string MappedListname { get; set; }

        public ListMapping(string ListName, string ListType, string MappedListname)
        {
            this.ListName = ListName;
            this.ListType = ListType;
            this.MappedListname = MappedListname;
        }
    }
}
