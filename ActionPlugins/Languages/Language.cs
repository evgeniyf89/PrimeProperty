using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.Languages
{
    public static class Language
    {
        private static readonly string _logicalName = "sl_language";
        private static EntityReference GetRef(string id, string name = null) => new EntityReference(_logicalName, new Guid(id)) { Name = name };
        public static EntityReference English => GetRef("a0523a24-20db-eb11-bacb-000d3a2c3636", "ENG");
    }
}
