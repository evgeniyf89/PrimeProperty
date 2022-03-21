using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SoftLine.ActionPlugins.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.Word_Template_Constructor
{
    class ParametersContstructor : IParametersContstructor
    {
        private readonly IOrganizationService _service;
        public ParametersContstructor(IOrganizationService service)
        {
            _service = service;
        }

        public Entity RetivePrintFormSettings(int templateid)
        {
            var query = $@"<fetch top='1' no-lock='true' >
                           <entity name='sl_print_form_settings' >
                             <attribute name='sl_print_form_settingsid' />
                             <attribute name='sl_name' />
                             <attribute name='sl_issavesharepointbit' />
                             <attribute name='sl_template_file_name' />
                             <attribute name='sl_template_folder' />
                             <attribute name='sl_open_in_wordbit' />  
                             <attribute name='sl_document_formatcode' />  
                             <filter type='and' >
                               <condition attribute='sl_identifier_template' operator='eq' value='1' />
                               <condition attribute='statecode' operator='eq' value='0' />
                             </filter>
                           </entity>
                         </fetch>";
            var setting = _service.RetrieveMultiple(new FetchExpression(query))
                .Entities
                .FirstOrDefault();
            return setting ?? throw new Exception($"No template with Template ID = {templateid}");
        }

        public PrintFormSettings GetPrintFormSettings(int templateid)
        {
            var entity = RetivePrintFormSettings(templateid);            
            return Convert(entity);
        }

        private PrintFormSettings Convert(Entity printFormEntity)
        {
            var containsDocumentFormat = printFormEntity.Contains("sl_document_formatcode");

            var documentFormatValue = containsDocumentFormat
                ? ((OptionSetValue)printFormEntity["sl_document_formatcode"]).Value
                : default;
            var documentFormatText = containsDocumentFormat
                ? printFormEntity.FormattedValues["sl_document_formatcode"]
                : default;
            var documentFormat = !string.IsNullOrEmpty(documentFormatText)
              ? new OptionValue(documentFormatText, documentFormatValue)
              : default;
            return new PrintFormSettings()
            {
                DocumentFormat = documentFormat,
                TemplateID = printFormEntity.GetAttributeValue<int>("sl_identifier_template"),
                IsOpen = printFormEntity.GetAttributeValue<bool>("sl_open_documentbit"),
                IsOpenToWord = printFormEntity.GetAttributeValue<bool>("sl_open_in_wordbit"),
                IsSaveToSharePoint = printFormEntity.GetAttributeValue<bool>("sl_issavesharepointbit"),
                TemplateFolderInSharePoint = printFormEntity.GetAttributeValue<string>("sl_template_folder"),
                TemplateNameInSharePoint = printFormEntity.GetAttributeValue<string>("sl_template_file_name"),
            };
        }
    }
}
