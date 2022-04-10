using SoftLine.ActionPlugins.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.WordTemplateConstructor
{
    interface IParametersContstructor
    {
        PrintFormSettings GetPrintFormSettings(int templateid, Guid entityid, Guid languageId);
    }
}
