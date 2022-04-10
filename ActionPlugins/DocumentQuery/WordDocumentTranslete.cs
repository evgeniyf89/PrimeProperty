using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SoftLine.ActionPlugins.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.DocumentQuery
{
    class WordDocumentTranslete : IDocumentRunQuery
    {
        private IOrganizationService _crmService;
        public void Init(IOrganizationService crmService)
        {
            _crmService = crmService;
        }

        public WordParameters[] RunQuery(Guid entityId)
        {
            var documentPatemeterRecodrIdMap = new Dictionary<string, string>()
            {
                ["01001"] = "ПодробнееНаименование",
                ["01003"] = "ОтчетНаименование",
                ["01004"] = "ГородНаименование",
                ["01005"] = "РайонНаименование",
                ["01006"] = "СтадияНаименование",
                ["01007"] = "ДоМоряНаименование",
                ["01008"] = "ЦенаОтНаименование",
                ["01009"] = "ОкончаниеСтроительстваНаименование",
                ["01010"] = "ОсновныеПреимуществаНаименование",
                ["01014"] = "ЭтажЗаголовок",
                ["01015"] = "ТипЗаголовок",
                ["01016"] = "СпальниЗаголовок",
                ["01017"] = "ВнутренняяЗаголовок",
                ["01018"] = "КрытаяВерандаЗаголовок",
                ["01019"] = "ОткрытаяВерандаЗаголовок",
                ["01020"] = "ТерассаЗаголовок",
                ["01021"] = "ОбщееПользованиеЗаголовок",
                ["01022"] = "ОбщаяПлощадьЗаголовок",
                ["01023"] = "БассейнЗаголовок",
                ["01024"] = "ЦенаЗаголовок",
                ["01025"] = "НДСЗаголовок",
                ["01026"] = "УчастокЗаголовок",
                ["01027"] = "КладовкаЗаголовок",
                ["01028"] = "ПаркинЗаголовок",
                ["01029"] = "M2Заголовок",
                ["01030"] = "IBP",
                ["01037"] = "Т",
                ["01038"] = "Ф",
                ["01039"] = "Е",
                ["01040"] = "Логотип",

            };
            var strBuilder = new StringBuilder();
            foreach (var key in documentPatemeterRecodrIdMap.Keys)
            {
                strBuilder.Append($"<value>{key}</value>");
            }
            var query = $@"<fetch distinct='true' no-lock='true' >
                           <entity name='sl_word_for_report' >
                             <attribute name='sl_word' />
                             <attribute name='sl_id' />
                              <filter>
                                <condition attribute='sl_id' operator='in' >                                 
                                  {strBuilder}
                                </condition>
                              </filter>
                             <link-entity name='sl_word_translation' from='sl_wordid' to='sl_word_for_reportid' link-type='outer' alias='translation' >
                               <attribute name='sl_name' />
                               <filter>
                                 <condition attribute='sl_languageid' operator='eq' value='{entityId}' />
                               </filter>
                             </link-entity>
                           </entity>
                         </fetch>";
            var translations = _crmService
                .RetrieveMultiple(new FetchExpression(query))
                .Entities
                .GroupBy(x => x.GetAttributeValue<string>("sl_id"))
                .ToArray();
            if (translations.Length != documentPatemeterRecodrIdMap.Count)
            {
                var wordIds = translations
                    .Select(x => x.Key)
                    .ToArray();
                var noParameters = documentPatemeterRecodrIdMap.Keys
                    .Where(x => !wordIds.Contains(x));
                var ids = string.Join("; ", noParameters);
                throw new Exception($"No translation word with id: {ids}");
            }

            return translations
                .Select(x =>
                {
                    var firstTranslate = x.First();
                    var traslate = firstTranslate.GetAttributeValue<AliasedValue>("translation.sl_name")?.Value as string
                        ?? firstTranslate.GetAttributeValue<string>("sl_word");
                    return new WordParameters()
                    {
                        WordDocVariable = documentPatemeterRecodrIdMap[x.Key],
                        Value = traslate
                    };
                })
                .ToArray();
        }
    }
}
