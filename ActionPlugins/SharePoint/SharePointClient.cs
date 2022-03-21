using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.SharePoint
{
    class SharePointClient : ISharePointClient
    {
        private readonly ClientContext _context;
        public SharePointClient(string sharepointUrl, ICredentials sharepointCredentials)
        {
            _context = new ClientContext(sharepointUrl)
            {
                Credentials = sharepointCredentials
            };
        }

        public void Dispose()
        {
            _context.Dispose();
        }

        public byte[] GetFileByAbsoluteUrl(string url)
        {
            var web = _context.Web;
            var file = web.GetFileByUrl(url);
            _context.Load(file,
                i => i.Exists,
                u => u.Name,
                u => u.ListItemAllFields["FileDirRef"]);
            _context.ExecuteQuery();
            if (!file.Exists) throw new Exception($"In SharePoint doesn't exist file bu url = {url}");
            var stream = file.OpenBinaryStream();
            _context.ExecuteQuery();

            using (var mStream = new MemoryStream())
            {
                stream.Value.CopyTo(mStream);
                return mStream.ToArray();
            }
        }
    }
}
