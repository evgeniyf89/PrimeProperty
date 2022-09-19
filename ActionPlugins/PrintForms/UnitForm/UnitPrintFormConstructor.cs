using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SoftLine.ActionPlugins.Extensions;
using SoftLine.ActionPlugins.OptionSets;
using SoftLine.ActionPlugins.PrintForms.Metadata;
using SoftLine.ActionPlugins.PrintForms.UnitForm.Model;
using SoftLine.ActionPlugins.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.PrintForms.UnitForm
{
    public class UnitPrintFormConstructor : FormConstructor<UnitPrintFrom>
    {

        public UnitPrintFormConstructor(IOrganizationService crmService, ISharePointClient sharePointClient) : base(crmService, sharePointClient)
        {

        }
        public override UnitPrintFrom GetForms(InputPrintFormData inputData)
        {
            var languageid = inputData.Language.Id;
            var property = CrmService.Retrieve("sl_unit", inputData.TargetEntityIds[0], new ColumnSet(true));
            var wordWithTranslation = RetriveWordWithTranslation(languageid);
            var propertyTypeTranslations = RetrivePropertyTypeTranslation(languageid);
            var cityRef = property.GetAttributeValue<EntityReference>("sl_cityid");
            var areaRef = property.GetAttributeValue<EntityReference>("sl_areaid");
            var isExclusive = property.GetAttributeValue<bool>("sl_exclusivebit");

            var metadataByPromotionType = GetMetadataByPromotionType(inputData, property, wordWithTranslation);
            var obj = new Model.Object()
            {
                Area = $"{GetCityTranslate(cityRef, languageid)}.{GetAreaTranslate(areaRef, languageid)}",
                IsExclusive = isExclusive,
                Description = GetDescriptions(inputData, property),
                Price = metadataByPromotionType.Price
            };


            if (inputData.IsWithId == true)
            {
                obj.ObjectId = new ObjectId()
                {
                    Value = property.GetAttributeValue<string>("sl_id"),
                    Label = GetMetadataTranslate(FormMetadata.Id, languageid),
                };
            }

            
            return new UnitPrintFrom()
            {
                Metadata = new Model.Metadata()
                {
                    WithLogo = inputData.IsWithLogo,
                    ReportName = metadataByPromotionType.ReportName,
                    ObjectLabel = GetObjectLabel(property, propertyTypeTranslations),
                    AreaLabel = GetMetadataTranslate(FormMetadata.RegionName, languageid),
                    ExclusiveLabel = isExclusive ? "Exclusive" : string.Empty,
                    CharacteristicsLabel = GetMetadataTranslate(FormMetadata.CharacteristicsLabel, languageid),
                    DescriptionLabel = GetMetadataTranslate(FormMetadata.DescriptionLabel, languageid),
                    PriceLabel = metadataByPromotionType.PriceLabel,                    
                },
                Objects = new List<Model.Object>()
                {
                    obj
                }
            };
        }

        private List<Description> GetDescriptions(InputPrintFormData inputData, Entity property)
        {
            var descriptions = new List<Description>();
            var propertyOwnerRef = property.GetAttributeValue<EntityReference>("sl_property_ownerid");
            var projectRef = property.GetAttributeValue<EntityReference>("sl_projectid");
            var project = projectRef is null
                ? default
                : CrmService.Retrieve(projectRef.LogicalName, projectRef.Id, new ColumnSet("sl_id"));
            string getData(FormMetadata metadata)
            {
                return GetMetadataTranslate(metadata, inputData.Language.Id);
            }
            if (inputData.IsWithLogo)
            {
                var description = new Description()
                {
                    Label = getData(FormMetadata.ProjectName),
                    Value = $@"({getData(FormMetadata.PidName)}:{project?.GetAttributeValue<string>("sl_id")}) {projectRef?.Name}"
                };
                descriptions.Add(description);
            }
            var descriptionOwner = new Description()
            {
                Label = getData(FormMetadata.OwnerName),
                Value = $@"({getData(FormMetadata.UidName)}{project?.GetAttributeValue<string>("sl_id")}) {propertyOwnerRef?.Name}"
            };
            descriptions.Add(descriptionOwner);
            var descriptionCommision = new Description()
            {
                Label = getData(FormMetadata.CommisionName),
                Value = inputData.PromotionType.Value == ((int)PromotionType.PurchaseProperty)
                    ? $"{property.GetAttributeValue<decimal>("sl_commision"):0.0#} %"
                    : inputData.PromotionType.Value == ((int)PromotionType.StRent)
                        ? $"{property.GetAttributeValue<decimal>("sl_commission_for_sr"):0.0#} %" : string.Empty
            };
            descriptions.Add(descriptionCommision);

            var descriptionNotesName = new Description()
            {
                Label = getData(FormMetadata.NotesName),
                Value = property.GetAttributeValue<string>("sl_notes_generaal")
            };
            descriptions.Add(descriptionNotesName);

            var propertyOwner = propertyOwnerRef is null
                ? default
                : CrmService.Retrieve(propertyOwnerRef.LogicalName, propertyOwnerRef.Id, new ColumnSet("emailaddress1"));
            var descriptionOwnerName = new Description()
            {
                Label = getData(FormMetadata.OwnerEmail),
                Value = propertyOwner?.GetAttributeValue<string>("emailaddress1")
            };
            descriptions.Add(descriptionOwnerName);
            var descriptionGoogleName = new Description()
            {
                Label = getData(FormMetadata.GoogleName),
                Value = $"{property.GetAttributeValue<decimal>("sl_google_map_lat").ToString("0.######", System.Globalization.CultureInfo.InvariantCulture)}, {property.GetAttributeValue<decimal>("sl_google_map_lng").ToString("0.######", System.Globalization.CultureInfo.InvariantCulture)}"
            };
            descriptions.Add(descriptionGoogleName);
            return descriptions;
        }

        private (string ReportName, string PriceLabel, string Price) GetMetadataByPromotionType(InputPrintFormData inputData, Entity property, IEnumerable<Entity> wordWithTranslation)
        {
            string getTranslate(string attrName)
            {
                var optionSet = property.GetAttributeValue<OptionSetValue>(attrName);
                var optionSetString = property.FormattedValues.Contains(attrName)
                    ? property.FormattedValues[attrName]
                    : string.Empty;
                var translate = wordWithTranslation
                        .FirstOrDefault(x => x.GetAttributeValue<string>("sl_option_set") == attrName
                        && x.GetAttributeValue<int>("sl_value") == optionSet?.Value
                        && optionSetString.Equals(x.GetAttributeValue<string>("sl_word"), StringComparison.OrdinalIgnoreCase));
                return translate?.GetValue<string>("translation.sl_name") ?? optionSetString;
            }
            string reportName = default;
            string priceLabel = default;
            string price = default;

            var listing = new Lazy<Entity>(() => RetriveListing(inputData, property.Id));
            
            switch (inputData.PromotionType.Value)
            {
                case (int)PromotionType.PurchaseProperty:
                    reportName = getTranslate("sl_sale_statuscode");
                    priceLabel = GetMetadataTranslate(FormMetadata.PriceLabelPromotion, inputData.Language.Id);
                    var saleStatusCode = property.GetAttributeValue<OptionSetValue>("sl_sale_statuscode");
                    price = saleStatusCode?.Value != 588610001 
                        ? reportName
                        : "€ " + string.Format(CultureInfo, "{0:0,##}", listing.Value?.GetAttributeValue<Money>("sl_sale_starting_price").Value); 
                    return (reportName, priceLabel, price);
                case (int)PromotionType.LtRent:
                    reportName = getTranslate("sl_ltr_statusccode");
                    priceLabel = GetMetadataTranslate(FormMetadata.PriceLabelLtRent, inputData.Language.Id);
                    var ltrStatuscCode = property.GetAttributeValue<OptionSetValue>("sl_ltr_statusccode");
                    price = ltrStatuscCode?.Value != 588610001
                        ? reportName
                        : "€ / month " + string.Format(CultureInfo, "{0:0,##}", listing.Value?.GetAttributeValue<Money>("sl_long_rent").Value);  
                    return (reportName, priceLabel, price);
                case ((int)PromotionType.StRent):
                    reportName = getTranslate("sl_str_statuscode");
                    priceLabel = GetMetadataTranslate(FormMetadata.PriceLabelStRent, inputData.Language.Id);
                    var strStatuscCode = property.GetAttributeValue<OptionSetValue>("sl_str_statuscode");
                    price = strStatuscCode?.Value != 588610001
                        ? reportName
                        : "€ / day " + string.Format(CultureInfo, "{0:0,##}", listing.Value?.GetAttributeValue<Money>("sl_short_rent").Value);  
                    return (reportName, priceLabel, price);

                default:
                    var empty = string.Empty;
                    return (empty, empty, empty);
            }
        }

        private string GetObjectLabel(Entity property, IEnumerable<Entity> propertyTypeTranslations)
        {
            var propertyTypeRef = property.GetAttributeValue<EntityReference>("sl_property_typeid");
            var propertyTypeTranslation = propertyTypeTranslations
                .FirstOrDefault(x => x.GetAttributeValue<EntityReference>("sl_property_typeid")?.Id == propertyTypeRef?.Id);
            var name = propertyTypeTranslation?.GetAttributeValue<string>("sl_name");
            return string.IsNullOrEmpty(name)
                ? propertyTypeRef?.Name
                : name;
        }
    }
}
