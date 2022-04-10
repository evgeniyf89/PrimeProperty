using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using Microsoft.Xrm.Sdk;
using SoftLine.ActionPlugins.OptionSets;

namespace SoftLine.ActionPlugins.Model
{
    public class PrintFormSettings
    {
        public bool IsOpenToWord { get; set; }

        public bool IsOpen { get; set; }

        public int TemplateID { get; set; }

        public bool IsSaveToSharePoint { get; set; }

        [Required(ErrorMessage = "Template name in SharePoint is empty")]
        public string TemplateNameInSharePoint { get; set; }

        [Required(ErrorMessage = "Document format is empty")]
        public OptionValue DocumentFormat { get; set; }

        [Required(ErrorMessage = "Template name in SharePoint is empty")]
        public string TemplateFolderInSharePoint { get; set; }

        public IEnumerable<WordParameters> Parameters { get; set; }

        public IEnumerable<Image> Images { get; set; }

        public bool ValidData(out string message)
        {
            var results = new List<ValidationResult>();
            var context = new ValidationContext(this);
            var isValid = Validator.TryValidateObject(this, context, results);
            message = isValid
                ? string.Empty
                : string.Join(Environment.NewLine, results);
            return isValid;
        }

    }



    public class OptionValue
    {
        public OptionValue(string name, int value)
        {
            Name = name;
            Value = value;
        }
        public string Name { get; set; }
        public int Value { get; set; }
    }

    public class WordParameters
    {
        public string Format { get; set; }
        public bool IsToWord { get; set; }
        public DocumentParameterFieldType FieldType { get; set; }
        public object Value { get; set; }
        public string WordDocVariable { get; set; }
        public bool IsRequired { get; set; }

        public string CrmParameter { get; set; }
        public string Name { get; set; }
    }

    public class Image
    {
        public string WordDocVariable { get; set; }

        public byte[] Value { get; set; } 
    }
}
