﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.SharePoint
{
    interface ISharePointClient : IDisposable
    {
        byte[] GetFileByAbsoluteUrl(string url);
    }
}
