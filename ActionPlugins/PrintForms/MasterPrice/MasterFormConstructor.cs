using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SoftLine.ActionPlugins.Extensions;
using SoftLine.ActionPlugins.Languages;
using SoftLine.ActionPlugins.OptionSets;
using SoftLine.ActionPlugins.PrintForms.MasterPrice.Model;
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

namespace SoftLine.ActionPlugins.PrintForms.MasterPrice
{
    public class MasterFormConstructor : FormConstructor<MasterPriceModel>
    {
        public MasterFormConstructor(IOrganizationService crmService, ISharePointClient sharePointClient) : base(crmService, sharePointClient)
        {
        }

        public override MasterPriceModel GetForms(InputPrintFormData inputData)
        {
            var metadata = RetriveMetadata(inputData.Language.Id);
            var projects = base.RetriveLinkProjectData(inputData);
            string getData(FormMetadata priceMetadata)
            {
                var data = metadata[priceMetadata.StringValue()].FirstOrDefault();
                return data.GetValue<string>("translation.sl_name")
                    ?? data.GetAttributeValue<string>("sl_word");
            }
            var account = RetriveAccount();
            var isEnglish = Language.English.Equals(inputData.Language);
            var printForm = new MasterPriceModel()
            {
                Metadata = new Model.Metadata()
                {
                    Address = account.GetAttributeValue<string>("sl_address"),
                    Email = account.GetAttributeValue<string>("emailaddress1"),
                    Fax = account.GetAttributeValue<string>("fax"),
                    WebsiteUrl = account.GetAttributeValue<string>("websiteurl"),
                    Phone = account.GetAttributeValue<string>("telephone1"),
                    EmailLabel = getData(FormMetadata.EmailLabel),
                    PhoneLabel = getData(FormMetadata.PhoneLabel),
                    FaxLabel = getData(FormMetadata.FaxLabel),
                    Tagline = getData(FormMetadata.Tagline),
                    Notification = getData(FormMetadata.SellerNotification),
                    WithLogo = inputData.IsWithLogo,
                    MainTableColumns = new List<Model.MainTableColumn>()
                },
                ObjectGroups = new List<ObjectGroup>()
            };
            var groupProjectsByCategory = projects
                .GroupBy(x => new
                {
                    Category = x.FormattedValues.Contains("sl_project_categorycode")
                        ? x.FormattedValues["sl_project_categorycode"]
                        : default
                });
            foreach (var groupByCategory in groupProjectsByCategory)
            {
                var objectGroup = new ObjectGroup()
                {
                    GroupName = groupByCategory.Key.Category,
                    Objects = new List<Model.Object>()
                };
                var groupsById = groupByCategory.GroupBy(x => x.Id);
                foreach (var byId in groupsById)
                {

                    var project = byId.First();
                    var pictureUrl = project.GetValue<string>("picture.sl_url");
                    var pictureByte = string.IsNullOrEmpty(pictureUrl)
                        ? default
                        : SharePointClient.GetFileByAbsoluteUrl(pictureUrl);

                    var pictureBase64 = pictureByte is null
                        ? default
                        : $"data:{MimeType.Jpeg};base64,{Convert.ToBase64String(pictureByte)}";

                    var floorLabel = getData(FormMetadata.FloorTitle);
                    var typeLabel = getData(FormMetadata.TypeHeader);
                    var bedLabel = getData(FormMetadata.BedroomsTitle);
                    var innerTitle = getData(FormMetadata.InnerTitle);
                    var coveredVerandaTitle = getData(FormMetadata.CoveredVerandaTitle);
                    var openVerandaTitle = getData(FormMetadata.OpenVerandaTitle);
                    var terasseTitle = getData(FormMetadata.TerasseTitle);
                    var plotTitle = getData(FormMetadata.PlotTitle);
                    var generalUsageTitle = getData(FormMetadata.GeneralUsageTitle);
                    var storageTitle = getData(FormMetadata.StorageTitle);
                    var totalAreaHeader = getData(FormMetadata.TotalAreaHeader);
                    var poolTitle = getData(FormMetadata.PoolTitle);
                    var parkinTitle = getData(FormMetadata.ParkinTitle);
                    var priceHeader = getData(FormMetadata.PriceHeader);
                    var m2Title = getData(FormMetadata.M2Title);
                    var vat = getData(FormMetadata.VATHeader);
                    var cityLabel = getData(FormMetadata.CityName);
                    var regionName = getData(FormMetadata.RegionName);
                    var stageLabel = getData(FormMetadata.StageName);
                    var toSeaName = getData(FormMetadata.ToSeaName);
                    var priceFromName = getData(FormMetadata.PriceFromName);
                    var completedConstructionName = getData(FormMetadata.CompletedConstructionName);
                    var yesWord = getData(FormMetadata.IBP);
                    var isLtRent = inputData.PromotionType?.Value == ((int)PromotionType.LtRent);

                    var _empty = string.Empty;
                    var flats = RetriveFlats(inputData.PromotionType?.Value, project.Id);
                    var flatsType = new Lazy<Dictionary<Guid, string>>(() =>
                    {
                        var typeGuid = flats
                        .Select(x => x.GetAttributeValue<EntityReference>("sl_property_typeid"))
                        .Where(x => x != null);
                        var flatFilter = string.Join(_empty, typeGuid.Select(x => $"<value>{x.Id}</value>"));
                        var query = $@"<fetch no-lock='true'>
                               <entity name='sl_property_type'>
                                 <attribute name='sl_name' />
                                 <filter type='and'>
                                   <condition attribute='sl_property_typeid' operator='in' >
                                   {flatFilter}
                                   </condition>
                                 </filter>
                                 <link-entity name='sl_property_type_translation' from='sl_property_typeid' to='sl_property_typeid' link-type='outer' alias='ad'>
                                   <attribute name='sl_name' />
                                   <filter>
                                     <condition attribute='sl_languageid' operator='eq' value='{inputData.Language.Id}' />
                                   </filter>
                                 </link-entity>
                               </entity>
                             </fetch>";
                        return CrmService.RetrieveMultiple(new FetchExpression(query))
                        .Entities
                        .GroupBy(x => x.Id)
                        .ToDictionary(x => x.Key, y =>
                        {
                            var value = y.FirstOrDefault();
                            return value.GetValue<string>("ad.sl_name") ?? value.GetValue<string>("sl_name");
                        });
                    });

                    var prices = RetriveListings(inputData, flats, getData, out decimal minPrice);
                    var isWithId = project.GetAttributeValue<bool>("sl_idbit");
                    var isFloor = project.GetAttributeValue<bool>("sl_foorbit");
                    var isType = project.GetAttributeValue<bool>("sl_typebit");
                    var isBed = project.GetAttributeValue<bool>("sl_bedbit");
                    var isIndoor = project.GetAttributeValue<bool>("sl_indoorbit");
                    var isCovver = project.GetAttributeValue<bool>("sl_covverbit");
                    var isUncovver = project.GetAttributeValue<bool>("sl_uncovverbit");
                    var isRoofTerrace = project.GetAttributeValue<bool>("sl_roof_terracebit");
                    var isPlot = project.GetAttributeValue<bool>("sl_plotbit");
                    var isCommon = project.GetAttributeValue<bool>("sl_commonbit");
                    var isStorage = project.GetAttributeValue<bool>("sl_storagebit");
                    var isTotal = project.GetAttributeValue<bool>("sl_total_areabit");
                    var isPoolb = project.GetAttributeValue<bool>("sl_poolbit");
                    var isParking = project.GetAttributeValue<bool>("sl_parkingbit");
                    var isPrice = project.GetAttributeValue<bool>("sl_pricebit");

                    var poolDictinary = new Lazy<Dictionary<Guid, string>>(() => GetPoolDataByFlatId(inputData, flats));

                    var flatsList = flats
                        .Select(flat =>
                        {
                            var typeRef = flat.GetAttributeValue<EntityReference>("sl_property_typeid");
                            var type = isEnglish
                                ? typeRef?.Name
                                : typeRef is null ? default : flatsType.Value[typeRef.Id];
                            return new Model.Flat()
                            {
                                Bedrooms = flat.FormattedValue("sl_number_of_bedrooms"),
                                CommonArea = flat.FormattedValue("sl_common_area"),
                                CovereddVeranda = flat.FormattedValue("sl_covered_veranda"),
                                FlatNumber = flat.FormattedValue("sl_number"),
                                Floor = flat.FormattedValue("sl_floor"),
                                Ibp = flat.GetAttributeValue<bool>("sl_ibpbit") ? yesWord : string.Empty,
                                IndoorArea = flat.FormattedValue("sl_indoor"),
                                Pool = isPoolb ? poolDictinary.Value[flat.Id] : default,
                                Price = prices[flat.Id],
                                TotalArea = flat.FormattedValue("sl_total_area"),
                                RoofTerrace = flat.FormattedValue("sl_roof_terrace"),
                                Type = type,
                                UncoveredVeranda = flat.FormattedValue("sl_open_veranda"),
                                Vat = isLtRent ? yesWord : string.Empty,
                            };
                        })
                        .ToList();
                    var block = new Block()
                    {
                        FlatGroups = new List<FlatGroup>()
                    };
                    var flatGroup = new FlatGroup()
                    {
                        Flats = flatsList
                    };
                    block.FlatGroups.Add(flatGroup);
                    var emptyList = new List<string>();

                    var flatColumns = new (string Header, string Dimension, bool IsVisible)[]
                    {
                       ("№", _empty, true),
                       ("IBP",_empty,  isLtRent),
                       ("ID",_empty,  isWithId),
                       (floorLabel,_empty,  isFloor),
                       (typeLabel,_empty, isType),
                       (bedLabel,_empty,  isBed),
                       (innerTitle, m2Title, isIndoor),
                       (coveredVerandaTitle, m2Title,  isCovver),
                       (openVerandaTitle,m2Title,  isUncovver),
                       (terasseTitle,m2Title,  isRoofTerrace ),
                       (plotTitle,m2Title, isPlot),
                       (generalUsageTitle,m2Title, isCommon),
                       (storageTitle, m2Title, isStorage),
                       (totalAreaHeader, m2Title, isTotal),
                       (poolTitle,_empty, isPoolb),
                       (parkinTitle,_empty,  isParking),
                       (priceHeader,_empty,  isPrice),
                       (vat, _empty,  isLtRent)
                    };

                    var minPriceStr = minPrice == default ? string.Empty : minPrice.ToString("N", new CultureInfo("en-US"));
                    var majorBenefits = project.GetValue<string>("projectlanguage.sl_major_benefit")
                        ?? project.GetAttributeValue<string>("sl_major_benefit");
                    printForm.Metadata.MainTableColumns = flatColumns
                            .Where(x => x.IsVisible)
                            .Select(x => new Model.MainTableColumn(x.Header, x.Dimension))
                            .ToList();
                    var objectMaster = new Model.Object()
                    {
                        Preview = pictureBase64,
                        DetailsLabel = getData(FormMetadata.DetailsLabel),
                        MajorBenefitsLabel = getData(FormMetadata.MainBenefitsName),
                        MajorBenefits = FindFourParagraph(majorBenefits),
                        ObjectName = project.GetAttributeValue<string>("sl_name"),
                        Details = new List<Model.Detail>()
                        {
                            new Model.Detail(cityLabel, project.GetValue<string>("city.sl_name")),
                            new Model.Detail(regionName, project.GetValue<string>("area.sl_name")),
                            new Model.Detail(stageLabel, project.FormattedValues.Contains("sl_stagecode") ? project.FormattedValues["sl_stagecode"] : default),
                            new Model.Detail(toSeaName, $"{project.GetAttributeValue<decimal>("sl_distance_sea"):0.##} m"),
                            new Model.Detail(priceFromName, minPriceStr),
                            new Model.Detail(completedConstructionName, project.GetAttributeValue<DateTime?>("sl_completion_date")?.ToString("dd.MM.yyyy")),
                        },
                        Blocks = new List<Model.Block>() { block }
                    };
                    objectGroup.Objects.Add(objectMaster);
                }
                printForm.ObjectGroups.Add(objectGroup);
            }
            return printForm;
        }
    }
}
