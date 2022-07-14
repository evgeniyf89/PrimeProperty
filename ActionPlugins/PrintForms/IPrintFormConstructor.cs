using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.PrintForms
{
    public interface IPrintFormConstructor<T>
    {
        T GetForms(InputPrintFormData inputData);
    }
}
