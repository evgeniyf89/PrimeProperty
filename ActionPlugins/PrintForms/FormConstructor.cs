using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SoftLine.ActionPlugins.Extensions;
using SoftLine.ActionPlugins.Languages;
using SoftLine.ActionPlugins.OptionSets;
using SoftLine.ActionPlugins.PrintForms.Metadata;
using SoftLine.ActionPlugins.PrintForms.Project_Price_Form.Model;
using SoftLine.ActionPlugins.PrintForms.ProjectPriceForm.Model;
using SoftLine.ActionPlugins.SharePoint;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.PrintForms
{
    public abstract class FormConstructor<T>
    {
        internal IOrganizationService CrmService;
        internal ISharePointClient SharePointClient;

        private Dictionary<string, Entity[]> _metadata;
        private Dictionary<Guid, Entity> _cityTranslate;
        private Dictionary<Guid, Entity> _areaTranslate;

        public CultureInfo CultureInfo { get => new CultureInfo("en-US"); } 
        public FormConstructor(IOrganizationService crmService, ISharePointClient sharePointClient)
        {
            CrmService = crmService;
            SharePointClient = sharePointClient;
        }
        public abstract T GetForms(InputPrintFormData inputData);

        // public abstract T GetForm(InputPrintFormData inputData);

        public virtual string FindFourParagraph(string text)
        {
            if (string.IsNullOrEmpty(text)) return default;
            var i = 0;
            var paragraphIndex = 0;
            var findWord = "</li>";
            while (i < 4)
            {
                paragraphIndex = text.IndexOf(findWord, paragraphIndex);
                if (paragraphIndex == -1)
                    return text;
                i++;
                paragraphIndex += findWord.Length;
            }
            return text.Substring(0, paragraphIndex);
        }

        public virtual DataCollection<Entity> RetriveLinkProjectData(InputPrintFormData inputData)
        {
            var query = $@"<fetch no-lock='true' >
                           <entity name='sl_project' >
                             <attribute name='sl_name' />
                             <attribute name='sl_stagecode' />
                             <attribute name='sl_distance_sea' />
                             <attribute name='sl_completion_date' />
                             <attribute name='sl_major_benefit' />
                             <attribute name='sl_idbit' />
                             <attribute name='sl_foorbit' />
                             <attribute name='sl_typebit' />
                             <attribute name='sl_bedbit' />
                             <attribute name='sl_indoorbit' />
                             <attribute name='sl_covverbit' />
                             <attribute name='sl_uncovverbit' />
                             <attribute name='sl_roof_terracebit' />
                             <attribute name='sl_plotbit' />
                             <attribute name='sl_commonbit' />
                             <attribute name='sl_storagebit' />
                             <attribute name='sl_total_areabit' />
                             <attribute name='sl_poolbit' />
                             <attribute name='sl_parkingbit' />
                             <attribute name='sl_pricebit' />
                             <attribute name='sl_project_categorycode' />
                             <filter type='and'>
                                <condition attribute='sl_projectid' operator='in'>
                                 {string.Join(string.Empty, inputData.TargetEntityIds.Select(x => $"<value>{x}</value>"))} 
                                </condition>
                              </filter>
                             <link-entity name='sl_picture' from='sl_projectid' to='sl_projectid' link-type='outer' alias='picture' >
                               <attribute name='sl_url' />
                               <attribute name='sl_weight' />
                               <filter type='and' >
                                 <condition attribute='sl_typecode' operator='eq' value='102690000' />
                                 <condition attribute='sl_upload_formatid' operator='eq' uiname='317 x 178' uitype='sl_upload_format' value='31563CFC-9F9B-EC11-B400-000D3AC21CBD' />
                               </filter>
                               <order attribute='sl_weight' />
                             </link-entity>
                             <link-entity name='sl_city' from='sl_cityid' to='sl_cityid' link-type='outer' alias='city' >
                               <attribute name='sl_name' />
                             </link-entity>
                             <link-entity name='sl_area' from='sl_areaid' to='sl_areaid' link-type='outer' alias='area' >
                               <attribute name='sl_name' />
                             </link-entity>
                             <link-entity name='sl_projectlanguage' from='sl_projectid' to='sl_projectid' link-type='outer' alias='projectlanguage'>
                              <attribute name='sl_major_benefit' />
                              <filter type='and'>
                                <condition attribute='sl_languageid' operator='eq' value='{inputData.Language.Id}' />
                              </filter>
                            </link-entity>
                           </entity>
                         </fetch>";
            return CrmService.RetrieveMultiple(new FetchExpression(query))
                .Entities;
        }

        public virtual Entity RetriveAccount()
        {
            var query = $@"<fetch top='1' no-lock='true' >
                           <entity name='account' >
                             <attribute name='telephone1' />
                             <attribute name='websiteurl' />
                             <attribute name='fax' />
                             <attribute name='emailaddress1' />
                             <attribute name='sl_address' />
                             <filter type='and' >
                               <condition attribute='sl_ppg_rolecode' operator='eq' value='588610001' />
                             </filter>
                           </entity>
                         </fetch>";
            return CrmService.RetrieveMultiple(new FetchExpression(query))
                .Entities
                .FirstOrDefault();
        }

        public virtual DataCollection<Entity> RetriveFlats(int? promotionType, Guid projectid)
        {
            var attrs = new[]
            {
                "sl_number", "sl_floor", "sl_property_typeid", "sl_number_of_bedrooms",
                "sl_indoor","sl_covered_veranda","sl_open_veranda","sl_roof_terrace",
                "sl_common_area","sl_total_area", "sl_type_of_poolcode","sl_ibpbit","sl_id",
                "sl_plot_size","sl_storage_size","sl_parking","sl_ltr_statusccode","sl_sale_statuscode"
            };
            var filterCondition = promotionType == (int)PromotionType.PurchaseProperty
                ? "sl_sale_statuscode"
                : promotionType == (int)PromotionType.LtRent ? "sl_ltr_statusccode" : default;
            if (string.IsNullOrEmpty(filterCondition)) return new EntityCollection().Entities;
            var query = $@"<fetch no-lock='true' >
                   <entity name='sl_unit' >                         
                     {string.Join(string.Empty, attrs.Select(x => $"<attribute name='{x}' />"))}
                      <filter type='and' >
                        <condition attribute='{filterCondition}' operator='ne' value='588610000' />
                        <condition attribute='sl_projectid' operator='eq' value='{projectid}' />
                      </filter>
                    </entity>
                  </fetch>";
            return CrmService.RetrieveMultiple(new FetchExpression(query)).Entities;
        }

        private Dictionary<string, Entity[]> RetriveMetadata(Guid languageid)
        {
            var keys = Enum.GetValues(typeof(FormMetadata))
                .Cast<FormMetadata>()
                .Select(x => x.StringValue())
                .ToArray();
            var strBuilder = new StringBuilder();
            foreach (var key in keys)
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
                                 <condition attribute='sl_languageid' operator='eq' value='{languageid}' />
                               </filter>
                             </link-entity>
                           </entity>
                         </fetch>";
            return CrmService
                .RetrieveMultiple(new FetchExpression(query))
                .Entities
                .GroupBy(x => x.GetAttributeValue<string>("sl_id"))
                .ToDictionary(x => x.Key, y => y.ToArray());
        }

        public virtual List<string> GetPoolData(InputPrintFormData inputData, IEnumerable<Entity> flats)
        {
            var poolAttrName = "sl_type_of_poolcode";
            if (Language.English.Equals(inputData.Language))
                return flats
                    .Select(x => x.FormattedValues.Contains(poolAttrName) ? x.FormattedValues[poolAttrName] : default)
                    .ToList();
            var distinctPoolType = flats
                .Select(x => x.GetAttributeValue<OptionSetValue>(poolAttrName)?.Value)
                .Where(x => x.HasValue)
                .Distinct();
            if (distinctPoolType.Count() == 0) return new List<string>();
            var query = $@"<fetch no-lock='true'>
                           <entity name='sl_word_for_report'>
                             <attribute name='sl_word' />
                             <attribute name='sl_value' />
                             <filter type='and'>
                               <condition attribute='sl_option_set' operator='eq' value='sl_unit.sl_type_of_poolcode' />
                               <condition attribute='sl_value' operator='in'>
                                 {string.Join(string.Empty, distinctPoolType.Select(x => $"<value>{x}</value>"))}
                               </condition>
                             </filter>
                             <link-entity name='sl_word_translation' from='sl_wordid' to='sl_word_for_reportid' link-type='outer' alias='ab'>
                               <attribute name='sl_name' />
                                <filter>
                                   <condition attribute='sl_languageid' operator='eq' value='{inputData.Language.Id}' />
                                </filter>
                             </link-entity>
                           </entity>
                         </fetch>";
            var translations = CrmService.RetrieveMultiple(new FetchExpression(query)).Entities;
            return flats
                .Select(x =>
                {
                    var value = x.GetAttributeValue<OptionSetValue>(poolAttrName)?.Value;
                    if (!value.HasValue) return default;
                    var translation = translations.FirstOrDefault(t => t.GetAttributeValue<int>("sl_value") == value);
                    return translation?.GetValue<string>("ab.sl_name") ?? translation.GetAttributeValue<string>("sl_word");
                })
                .ToList();
        }

        public virtual Dictionary<Guid, string> GetPoolDataByFlatId(InputPrintFormData inputData, IEnumerable<Entity> flats)
        {
            var poolAttrName = "sl_type_of_poolcode";
            if (Language.English.Equals(inputData.Language))
                return flats
                    .ToDictionary(y => y.Id, x => x.FormattedValues.Contains(poolAttrName) ? x.FormattedValues[poolAttrName] : default);
            var distinctPoolType = flats
                .Select(x => x.GetAttributeValue<OptionSetValue>(poolAttrName)?.Value)
                .Where(x => x.HasValue)
                .Distinct();
            if (distinctPoolType.Count() == 0) return new Dictionary<Guid, string>();
            var query = $@"<fetch no-lock='true'>
                           <entity name='sl_word_for_report'>
                             <attribute name='sl_word' />
                             <attribute name='sl_value' />
                             <filter type='and'>
                               <condition attribute='sl_option_set' operator='eq' value='sl_unit.sl_type_of_poolcode' />
                               <condition attribute='sl_value' operator='in'>
                                 {string.Join(string.Empty, distinctPoolType.Select(x => $"<value>{x}</value>"))}
                               </condition>
                             </filter>
                             <link-entity name='sl_word_translation' from='sl_wordid' to='sl_word_for_reportid' link-type='outer' alias='ab'>
                               <attribute name='sl_name' />
                                <filter>
                                   <condition attribute='sl_languageid' operator='eq' value='{inputData.Language.Id}' />
                                </filter>
                             </link-entity>
                           </entity>
                         </fetch>";
            var translations = CrmService.RetrieveMultiple(new FetchExpression(query)).Entities;
            return flats
                .ToDictionary(y => y.Id, x =>
                   {
                       var value = x.GetAttributeValue<OptionSetValue>(poolAttrName)?.Value;
                       if (!value.HasValue) return default;
                       var translation = translations.FirstOrDefault(t => t.GetAttributeValue<int>("sl_value") == value);
                       return translation?.GetValue<string>("ab.sl_name") ?? translation.GetAttributeValue<string>("sl_word");
                   });
        }

        public virtual Dictionary<Guid, string> RetriveListings(InputPrintFormData inputData, IEnumerable<Entity> flats, Func<FormMetadata, string> getData, out decimal minPrice)
        {
            minPrice = 0;
            if (flats.Count() == 0) return new Dictionary<Guid, string>();
            var query = $@"<fetch no-lock='true'>
                           <entity name='sl_listing'>
                             <attribute name='sl_long_rent' />
                             <attribute name='sl_sale_starting_price' />
                             <attribute name='sl_unitid' />
                             <filter type='and'>
                               <condition attribute='sl_marketcode' operator='eq' value='{inputData.Market.Value}' />
                               <condition attribute='sl_clients_promotion_typecode' operator='eq' value='{inputData.PromotionType.Value}' />
                               <condition attribute='sl_unitid' operator='in'>
                                 {string.Join(string.Empty, flats.Select(x => $"<value>{x.Id}</value>"))}
                               </condition>
                             </filter>
                           </entity>
                         </fetch>";

            var listings = CrmService.RetrieveMultiple(new FetchExpression(query)).Entities;
            var isPurchaseofproperty = inputData.PromotionType.Value == (int)PromotionType.PurchaseProperty;
            var isLtRent = inputData.PromotionType.Value == (int)PromotionType.LtRent;
            var priceSold = getData(FormMetadata.Sold);
            var priceRented = getData(FormMetadata.Rented);
            decimal? fromPrice = decimal.MaxValue;
          
            var prices = flats
                .ToDictionary(y => y.Id,
                x =>
                {
                    var saleStatus = x.GetAttributeValue<OptionSetValue>("sl_sale_statuscode")?.Value;
                    var slLtrStatus = x.GetAttributeValue<OptionSetValue>("sl_ltr_statusccode")?.Value;
                    if (isPurchaseofproperty)
                    {
                        if (saleStatus == (int)UnitSaleStatusCode.Reserved || saleStatus == (int)UnitSaleStatusCode.Sold)
                            return priceSold;
                        if (saleStatus == ((int)UnitSaleStatusCode.ForSale))
                        {
                            var listing = listings.FirstOrDefault(l => l.GetAttributeValue<EntityReference>("sl_unitid")?.Id == x.Id);
                            var startingPrice = listing?.GetAttributeValue<Money>("sl_sale_starting_price").Value;
                            if (fromPrice > startingPrice)
                                fromPrice = startingPrice;
                            return startingPrice?.ToString("N", CultureInfo);
                        }
                    }
                    else if (isLtRent)
                    {
                        if (slLtrStatus == (int)UnitLTRStatusCode.Rented)
                            return priceRented;
                        if (slLtrStatus == (int)UnitLTRStatusCode.ForRent)
                        {
                            var listing = listings.FirstOrDefault(l => l.GetAttributeValue<EntityReference>("sl_unitid")?.Id == x.Id);
                            var startingPrice = listing?.GetAttributeValue<Money>("sl_long_rent").Value;
                            if (fromPrice > startingPrice)
                                fromPrice = startingPrice;
                            return startingPrice?.ToString("N", CultureInfo);
                        }
                    }
                    return default;
                });
            minPrice = fromPrice == decimal.MaxValue ? default : fromPrice.Value;
            return prices;
        }


        public virtual Entity RetriveListing(InputPrintFormData inputData, Guid propertyId)
        {            
            var query = $@"<fetch no-lock='true'>
                           <entity name='sl_listing'>
                             <attribute name='sl_long_rent' />
                             <attribute name='sl_short_rent' />
                             <attribute name='sl_sale_starting_price' />
                             <filter type='and'>
                               <condition attribute='sl_marketcode' operator='eq' value='{inputData.Market.Value}' />
                               <condition attribute='sl_clients_promotion_typecode' operator='eq' value='{inputData.PromotionType.Value}' />
                               <condition attribute='sl_unitid' operator='in'>
                                 <value>{propertyId}</value>
                               </condition>
                             </filter>
                           </entity>
                         </fetch>";
            return CrmService.RetrieveMultiple(new FetchExpression(query)).Entities.FirstOrDefault();
        }


        public virtual DataCollection<Entity> RetriveWordWithTranslation(Guid languageid)
        {
            var query = $@"<fetch no-lock='true'>
                           <entity name='sl_word_for_report'>
                             <attribute name='sl_word' />
                             <attribute name='sl_option_set' />
                             <attribute name='sl_value' />
                             <link-entity name='sl_word_translation' from='sl_wordid' to='sl_word_for_reportid' link-type='inner' alias='translation'>
                               <attribute name='sl_name' />
                               <filter type='and'>
                                 <condition attribute='sl_languageid' operator='eq' value='{languageid}' />
                               </filter>
                             </link-entity>
                           </entity>
                         </fetch>";
            return CrmService.RetrieveMultiple(new FetchExpression(query)).Entities;
        }

        public virtual DataCollection<Entity> RetrivePropertyTypeTranslation(Guid languageid)
        {
            var query = $@"<fetch no-lock='true'>
                           <entity name='sl_property_type_translation'>
                             <attribute name='sl_name' />
                             <attribute name='sl_property_typeid' />
                             <filter type='and'>
                               <condition attribute='sl_languageid' operator='eq' value='{languageid}' />
                             </filter>
                           </entity>
                         </fetch>";
            return CrmService.RetrieveMultiple(new FetchExpression(query)).Entities;
        }

        private Dictionary<Guid, Entity> RetriveCityTranslate(Guid languageid)
        {
            return RetriveTranslete(languageid, "sl_city_translation", "sl_cityid");
        }

        private Dictionary<Guid, Entity> RetriveAreaTranslate(Guid languageid)
        {
            return RetriveTranslete(languageid, "sl_area_translation", "sl_areaid");
        }

        public string GetCityTranslate(EntityReference cityRef, Guid languageid)
        {
            if (cityRef is null)
                return string.Empty;
            if (_cityTranslate is null)
            {
                _cityTranslate = RetriveCityTranslate(languageid);
            }
            return _cityTranslate.TryGetValue(cityRef.Id, out Entity translateCity)
                ? translateCity.GetAttributeValue<string>("sl_name")
                : cityRef.Name;
        }

        public string GetAreaTranslate(EntityReference areaRef, Guid languageid)
        {
            if (areaRef is null)
                return string.Empty;
            if (_areaTranslate is null)
            {
                _areaTranslate = RetriveAreaTranslate(languageid);
            }
            return _areaTranslate.TryGetValue(areaRef.Id, out Entity translateArea)
                ? translateArea.GetAttributeValue<string>("sl_name")
                : areaRef.Name;
        }

        private Dictionary<Guid, Entity> RetriveTranslete(Guid languageid, string entityLogicalName, string attr)
        {
            var query = $@"<fetch no-lock='true'>
                          <entity name='{entityLogicalName}'>
                            <attribute name='sl_name' />
                            <attribute name='{attr}' />
                            <filter type='and'>
                              <condition attribute='sl_languageid' operator='eq' value='{languageid}' />
                              <condition attribute='{attr}' operator='not-null' />
                            </filter>                           
                          </entity>
                        </fetch>";
            return CrmService.RetrieveMultiple(new FetchExpression(query))
                .Entities
                .GroupBy(x => x.GetAttributeValue<EntityReference>(attr).Id)
                .ToDictionary(x => x.Key, y => y.FirstOrDefault());
        }

        public string GetMetadataTranslate(FormMetadata formMetadata, Guid languageid)
        {
            if (_metadata is null)
                _metadata = RetriveMetadata(languageid);
            if (_metadata.TryGetValue(formMetadata.StringValue(), out Entity[] values))
            {
                var data = values.FirstOrDefault();
                return data.GetValue<string>("translation.sl_name")
                   ?? data.GetAttributeValue<string>("sl_word");
            }
            return string.Empty;
        }
    }
}
