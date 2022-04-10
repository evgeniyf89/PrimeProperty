using SoftLine.ActionPlugins.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.WordTemplateConstructor
{
    interface IWordDocumentFormater
    {
        byte[] Form(byte[] wordTemplate, PrintFormSettings settings);
    }
}
