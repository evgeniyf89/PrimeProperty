using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SoftLine.ActionPlugins.Extensions;
using SoftLine.ActionPlugins.OptionSets;
using SoftLine.ActionPlugins.PrintForms.Metadata;
using SoftLine.ActionPlugins.PrintForms.ProjectPriceForm.Model;
using SoftLine.ActionPlugins.SharePoint;

namespace SoftLine.ActionPlugins.PrintForms
{
    public class PrintFormConstructor : IPrintFormConstructor<ProjectPrintForm>
    {
        private readonly IOrganizationService _crmService;
        private readonly ISharePointClient _sharePointClient;
        public PrintFormConstructor(IOrganizationService crmService, ISharePointClient sharePointClient)
        {
            _crmService = crmService;
            _sharePointClient = sharePointClient;
        }
        public ProjectPrintForm GetForm(InputPrintFormData inputData)
        {
            var project = RetriveLinkProjectData(inputData.TargetEntityRef.Id);
            var pictureUrl = project.GetValue<string>("picture.sl_url");
            var pictureByte = string.IsNullOrEmpty(pictureUrl)
                ? default
                : _sharePointClient.GetFileByAbsoluteUrl(pictureUrl);

            var pictureBase64 = pictureByte is null
                ? default
                : $"data:{MimeType.Jpeg};base64,{Convert.ToBase64String(pictureByte)}";

            var metadata = RettiveMetadata(inputData.Language.Id);
            string getData(ProjectPriceMetadata priceMetadata)
            {
                var data = metadata[priceMetadata.StringValue()].FirstOrDefault();
                return data.GetValue<string>("translation.sl_name")
                    ?? data.GetAttributeValue<string>("sl_word");
            }


            var floorLabel = getData(ProjectPriceMetadata.FloorTitle);
            var typeLabel = getData(ProjectPriceMetadata.TypeHeader);
            var bedLabel = getData(ProjectPriceMetadata.BedroomsTitle);
            var innerTitle = getData(ProjectPriceMetadata.InnerTitle);
            var coveredVerandaTitle = getData(ProjectPriceMetadata.CoveredVerandaTitle);
            var openVerandaTitle = getData(ProjectPriceMetadata.OpenVerandaTitle);
            var terasseTitle = getData(ProjectPriceMetadata.TerasseTitle);
            var plotTitle = getData(ProjectPriceMetadata.PlotTitle);
            var generalUsageTitle = getData(ProjectPriceMetadata.GeneralUsageTitle);
            var storageTitle = getData(ProjectPriceMetadata.StorageTitle);

            var totalAreaHeader = getData(ProjectPriceMetadata.TotalAreaHeader);
            var poolTitle = getData(ProjectPriceMetadata.PoolTitle);
            var parkinTitle = getData(ProjectPriceMetadata.ParkinTitle);
            var priceHeader = getData(ProjectPriceMetadata.PriceHeader);
            var vat = getData(ProjectPriceMetadata.VATHeader);
            var isLtRent = inputData.PromotionType?.Value == ((int)PromotionType.LtRent);


            var columns = new List<MainTableColumn>()
            {
                new MainTableColumn("№", false),
                new MainTableColumn("IBP", isLtRent),
                new MainTableColumn("ID", !project.GetAttributeValue<bool>("sl_idbit")),
                new MainTableColumn(floorLabel, project.GetAttributeValue<bool>("sl_foorbit")),
                new MainTableColumn(typeLabel, project.GetAttributeValue<bool>("sl_typebit")),
                new MainTableColumn(bedLabel, project.GetAttributeValue<bool>("sl_bedbit")),
                new MainTableColumn(innerTitle, project.GetAttributeValue<bool>("sl_indoorbit")),
                new MainTableColumn(coveredVerandaTitle, project.GetAttributeValue<bool>("sl_covverbit")),
                new MainTableColumn(openVerandaTitle, project.GetAttributeValue<bool>("sl_uncovverbit")),
                new MainTableColumn(terasseTitle, project.GetAttributeValue<bool>("sl_roof_terracebit")),
                new MainTableColumn(plotTitle, project.GetAttributeValue<bool>("sl_plotbit")),
                new MainTableColumn(generalUsageTitle, project.GetAttributeValue<bool>("sl_commonbit")),
                new MainTableColumn(storageTitle, project.GetAttributeValue<bool>("sl_storagebit")),
                new MainTableColumn(totalAreaHeader, project.GetAttributeValue<bool>("sl_total_areabit")),
                new MainTableColumn(poolTitle, project.GetAttributeValue<bool>("sl_poolbit")),
                new MainTableColumn(parkinTitle, project.GetAttributeValue<bool>("sl_parkingbit")),
                new MainTableColumn(priceHeader, project.GetAttributeValue<bool>("sl_pricebit")),
                new MainTableColumn(vat,isLtRent)
            };

            var printForm = new ProjectPrintForm()
            {
                ObjectName = project.GetAttributeValue<string>("sl_name"),
                MajorBenefits = project.GetAttributeValue<string>("sl_major_benefit"),
                Details = new Details()
                {
                    DistanceToSea = $"{project.GetAttributeValue<decimal>("sl_distance_sea"):0:0.0#} m",
                    City = project.GetValue<string>("city.sl_name"),
                    Area = project.GetValue<string>("area.sl_name"),
                    Completion = project.FormattedValues.Contains("sl_stagecode") ? project.FormattedValues["sl_stagecode"] : default,
                    CompletionDate = project.GetAttributeValue<DateTime?>("sl_completion_date")?.ToString("dd.MM.yyyy"),
                    PricesFrom = "100"
                },
                Preview = pictureBase64,
                Metadata = new ProjectPriceForm.Model.Metadata()
                {
                    WithLogo = inputData.IsWithLogo,
                    MainTableColumns = columns
                },
                DetailsLabel = getData(ProjectPriceMetadata.DetailsLabel),
                MajorBenefitsLabel = getData(ProjectPriceMetadata.MainBenefitsName),
                PhoneLabel = getData(ProjectPriceMetadata.PhoneLabel),
                FaxLabel = getData(ProjectPriceMetadata.FaxLabel),
                EmailLabel = getData(ProjectPriceMetadata.EmailLabel),
                Tagline = getData(ProjectPriceMetadata.Tagline),
                MainTable = new List<MainTable>()
                {

                }
            };
            var account = new Lazy<Entity>(RetriveAccount);
            if (inputData.IsWithLogo && account.Value != default)
            {
                printForm.Email = account.Value.GetAttributeValue<string>("emailaddress1");
                printForm.Fax = account.Value.GetAttributeValue<string>("fax");
                printForm.WebsiteUrl = account.Value.GetAttributeValue<string>("websiteurl");
                printForm.Phone = account.Value.GetAttributeValue<string>("telephone1");
                printForm.Address = account.Value.GetAttributeValue<string>("sl_address");
            }

            return printForm;
        }

        private Entity RetriveLinkProjectData(Guid projectid)
        {
            var query = $@"<fetch no-lock='true' top='1' >
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
                             <filter type='and' >
                               <condition attribute='sl_projectid' operator='eq' value='{projectid}' />
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
                           </entity>
                         </fetch>";
            return _crmService.RetrieveMultiple(new FetchExpression(query))
                .Entities
                .FirstOrDefault();
        }

        private Entity RetriveAccount()
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
            return _crmService.RetrieveMultiple(new FetchExpression(query))
                .Entities
                .FirstOrDefault();
        }

        private List<string> RetriveFlats(InputPrintFormData inputData)
        {
            var attrs = new[]
            {
                "sl_number", "", "sl_floor", "sl_property_typeid", "sl_number_of_bedrooms",
                "sl_indoor","sl_covered_veranda","sl_open_veranda","sl_roof_terrace",
                "sl_common_area","sl_total_area", "sl_type_of_poolcode",
            };
            var query = $@"<fetch no-lock='true' >
                   <entity name='sl_unit' >    
                      <attribute name='' />
                      <attribute name='sl_id' />
                      <attribute name='' />
                      <attribute name='' />
                      <attribute name='' />
                      <attribute name='' />
                      <attribute name='' />
                      <attribute name='sl_plot_size' />
                      <attribute name='' />
                      <attribute name='sl_storage_size' />
                      <attribute name='' />
                      <attribute name='sl_ibpbit' />
                      <filter type='and' >
                        <condition attribute='sl_sale_statuscode' operator='ne' value='588610000' />
                        <condition attribute='sl_projectid' operator='eq' value='{inputData.TargetEntityRef.Id}' />
                      </filter>
                    </entity>
                  </fetch>";
            return null;
        }

        private Dictionary<string, Entity[]> RettiveMetadata(Guid languageid)
        {
            var keys = Enum.GetValues(typeof(ProjectPriceMetadata))
                .Cast<ProjectPriceMetadata>()
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
            return _crmService
                .RetrieveMultiple(new FetchExpression(query))
                .Entities
                .GroupBy(x => x.GetAttributeValue<string>("sl_id"))
                .ToDictionary(x => x.Key, y => y.ToArray());


            //return translations
            //    .Select(x =>
            //    {
            //        var firstTranslate = x.First();
            //        var traslate = firstTranslate.GetAttributeValue<AliasedValue>("translation.sl_name")?.Value as string
            //            ?? firstTranslate.GetAttributeValue<string>("sl_word");
            //        return new WordParameters()
            //        {
            //            WordDocVariable = documentPatemeterRecodrIdMap[x.Key],
            //            Value = traslate
            //        };
            //    })
            //    .ToArray();
        }
    }
}
