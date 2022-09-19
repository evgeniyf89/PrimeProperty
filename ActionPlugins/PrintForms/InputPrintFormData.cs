using Microsoft.Xrm.Sdk;
using SoftLine.ActionPlugins.PrintForms.Metadata;

namespace SoftLine.ActionPlugins.PrintForms
{
   public class InputPrintFormData
    {
        public OptionSetValue Market { get; set; }

        public EntityReference Language { get; set; }

        public bool IsWithLogo { get; set; }

        public OptionSetValue PromotionType { get; set; }

        public System.Guid[] TargetEntityIds { get; set; }
        public int PrintFormId { get; set; }
        public Unit Unit { get; set; }

        public bool? IsWithPrice { get; set; }

        public bool? IsWithId { get; set; }
        public bool? IsOnePage { get; set; }
    }
}
