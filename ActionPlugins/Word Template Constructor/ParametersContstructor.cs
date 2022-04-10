using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using SoftLine.ActionPlugins.DocumentQuery;
using SoftLine.ActionPlugins.Extensions;
using SoftLine.ActionPlugins.Model;
using SoftLine.ActionPlugins.OptionSets;
using SoftLine.ActionPlugins.SharePoint;
using System;
using System.Linq;

namespace SoftLine.ActionPlugins.WordTemplateConstructor
{
    class ParametersContstructor : IParametersContstructor
    {
        private readonly IOrganizationService _service;
        private readonly ISharePointClient _spClient;
        public ParametersContstructor(IOrganizationService service, ISharePointClient spClient)
        {
            _service = service;
            _spClient = spClient;
        }

        private Entity RetivePrintFormSettings(int templateid)
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
                               <condition attribute='sl_identifier_template' operator='eq' value='{templateid}' />
                               <condition attribute='statecode' operator='eq' value='0' />
                             </filter>
                           </entity>
                         </fetch>";
            var setting = _service.RetrieveMultiple(new FetchExpression(query))
                .Entities
                .FirstOrDefault();
            return setting ?? throw new Exception($"No template with Template ID = {templateid}");
        }

        public PrintFormSettings GetPrintFormSettings(int templateid, Guid entityid, Guid languageId)
        {
            var entity = RetivePrintFormSettings(templateid);        
            return Convert(entity, entityid, languageId);
        }


        private PrintFormSettings Convert(Entity printFormEntity, Guid entityid, Guid languageId)
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
            var settings = new PrintFormSettings()
            {
                DocumentFormat = documentFormat,
                TemplateID = printFormEntity.GetAttributeValue<int>("sl_identifier_template"),
                IsOpen = printFormEntity.GetAttributeValue<bool>("sl_open_documentbit"),
                IsOpenToWord = printFormEntity.GetAttributeValue<bool>("sl_open_in_wordbit"),
                IsSaveToSharePoint = printFormEntity.GetAttributeValue<bool>("sl_issavesharepointbit"),
                TemplateFolderInSharePoint = printFormEntity.GetAttributeValue<string>("sl_template_folder"),
                TemplateNameInSharePoint = printFormEntity.GetAttributeValue<string>("sl_template_file_name"),
            };
            var isValidData = settings.ValidData(out string errorMessage);        
            if (isValidData)
            {
                settings.Parameters = GetAllWordParameters(printFormEntity, entityid, languageId);
            }          
            return settings;
        }

        private WordParameters[] GetAllWordParameters(Entity printFormEntity, Guid entityid, Guid languageId)
        {
            var requests = RetriveRequests(printFormEntity.Id);
            return requests
                .Select(x =>
                {
                    var query = (x["sl_request_text"] as string).Replace("@Entity", entityid.ToString());
                    var requestMethod = (DocumentParameterRequestMethod)x.GetAttributeValue<OptionSetValue>("sl_request_methodcode").Value;
                    switch (requestMethod)
                    {
                        case DocumentParameterRequestMethod.FetchXml:
                            {
                                var data = _service.RetrieveMultiple(new FetchExpression(query))
                                                        .Entities[0];
                                var parameters = RetriveRequestParameter(x.Id);
                                return parameters
                                                .Select(p =>
                                                {
                                                    var type = (DocumentParameterFieldType)p.GetAttributeValue<OptionSetValue>("sl_field_typecode").Value;
                                                    var crmParameter = p.GetAttributeValue<string>("sl_parameter_crm");
                                                    return new WordParameters()
                                                    {
                                                        WordDocVariable = p.GetAttributeValue<string>("sl_document_parameter"),
                                                        FieldType = type,
                                                        CrmParameter = crmParameter,
                                                        IsRequired = p.GetAttributeValue<bool>("sl_required_fieldbit"),
                                                        IsToWord = p.GetAttributeValue<bool>("sl_parameterf_documentbit"),
                                                        Name = p.GetAttributeValue<string>("sl_name"),
                                                        Value = GetDataValue(data, crmParameter, type)
                                                    };
                                                });
                            }
                        case DocumentParameterRequestMethod.ServerCode:
                            {
                                var queryServerClass = JsonConvert.DeserializeObject<ServerCodeClass>(query);
                                var enemy = Activator.CreateInstance(null, queryServerClass.Class);
                                var documentRunQuery = (IDocumentRunQuery)enemy.Unwrap();
                                documentRunQuery.Init(_service);
                                return documentRunQuery.RunQuery(languageId);
                            }
                        default:
                            return Array.Empty<WordParameters>();
                    }

                })
                .SelectMany(x => x)
                .ToArray();
        }

        private DataCollection<Entity> RetriveRequests(Guid printFormEntityid)
        {
            var query = $@"<fetch no-lock='true' >
                           <entity name='sl_request_document' >                            
                             <attribute name='sl_request_text' />
                             <attribute name='sl_request_methodcode' />
                             <link-entity name='sl_request_document_sl_print_form_setti' from='sl_request_documentid' to='sl_request_documentid' visible='false' intersect='true' >
                               <link-entity name='sl_print_form_settings' from='sl_print_form_settingsid' to='sl_print_form_settingsid' alias='ag' >
                                 <filter type='and' >
                                   <condition attribute='sl_print_form_settingsid' operator='eq' value='{printFormEntityid}' />
                                 </filter>
                               </link-entity>
                             </link-entity>
                           </entity>
                         </fetch>";
            return _service.RetrieveMultiple(new FetchExpression(query)).Entities;
        }

        private DataCollection<Entity> RetriveRequestParameter(Guid requestid)
        {
            var query = $@"<fetch no-lock='true' >
                          <entity name='sl_document_parameter' >
                            <attribute name='sl_parameterf_documentbit' />
                            <attribute name='sl_field_typecode' />
                            <attribute name='sl_name' />
                            <attribute name='sl_document_parameter' />
                            <attribute name='sl_parameter_crm' />
                            <attribute name='sl_parameterf_documentbit' />
                            <attribute name='sl_required_fieldbit' />
                            <link-entity name='sl_document_parameter_request_document' from='sl_document_parameterid' to='sl_document_parameterid' visible='false' intersect='true' >
                              <link-entity name='sl_request_document' from='sl_request_documentid' to='sl_request_documentid' alias='ac' >
                                <filter type='and' >
                                  <condition attribute='sl_request_documentid' operator='eq' value='{requestid}' />
                                </filter>
                              </link-entity>
                            </link-entity>
                          </entity>
                        </fetch>";
            return _service.RetrieveMultiple(new FetchExpression(query)).Entities;
        }

        private object GetDataValue(Entity data, string attrName, DocumentParameterFieldType type)
        {
            switch (type)
            {
                case DocumentParameterFieldType.String:
                    return data.GetValue<string>(attrName);
                case DocumentParameterFieldType.OptionSetValue:
                    return data.FormattedValues.Contains(attrName) ? data.FormattedValues[attrName] : throw new Exception("No data type specified");
                case DocumentParameterFieldType.Int:
                    return data.GetValue<int>(attrName).ToString();
                case DocumentParameterFieldType.Decimal:
                    return data.GetValue<decimal>(attrName).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
                case DocumentParameterFieldType.Date:
                    var date = data.GetValue<DateTime?>(attrName);
                    return date.HasValue ? date.Value.ToString("dd.MM.yyyy") : default;
                case DocumentParameterFieldType.Image:
                    var url = data.GetValue<string>(attrName);
                    return  _spClient.GetFileByAbsoluteUrl(url);
                default:
                    throw new Exception($"No data type specified for {attrName}");

            }
        }
    }
}
