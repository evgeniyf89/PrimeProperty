using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.SharePoint.Model
{
    public class SharePointData
    {
        public SharePointData(ICredentials credentials, string url)
        {
            Credentials = credentials;
            Url = url;
        }
        public ICredentials Credentials { get;}
        public string Url { get; }
    }
}
