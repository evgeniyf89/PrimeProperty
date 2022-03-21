using SoftLine.ActionPlugins.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.Word_Template_Constructor
{
    interface IParametersContstructor
    {
        PrintFormSettings GetPrintFormSettings(int templateid);
    }
}
