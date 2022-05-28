using Microsoft.Xrm.Sdk;

namespace SoftLine.ActionPlugins.PrintForms
{
   public class InputPrintFormData
    {
        public EntityReference Market { get; set; }

        public EntityReference Language { get; set; }

        public bool IsWithLogo { get; set; }

        public OptionSetValue PromotionType { get; set; }

        public EntityReference TargetEntityRef { get; set; }
    }
}
