using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.PrintForms.Project_Price_Form.Model
{
    public class FlatTableColumn
    {
        public FlatTableColumn(string header, string dimension, IEnumerable<string> value, bool isInvisibility = false)
        {
            Header = header;
            Dimension = dimension;
            Value = value.ToArray();
            IsInvisibility = isInvisibility;
        }

        public string Header { get; }

        public string Dimension { get; }

        public string[] Value { get; }

        public bool IsInvisibility { get; }

    }
}
