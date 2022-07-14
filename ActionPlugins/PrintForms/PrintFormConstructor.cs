using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SoftLine.ActionPlugins.Extensions;
using SoftLine.ActionPlugins.Languages;
using SoftLine.ActionPlugins.OptionSets;
using SoftLine.ActionPlugins.PrintForms.Metadata;
using SoftLine.ActionPlugins.PrintForms.Project_Price_Form.Model;
using SoftLine.ActionPlugins.PrintForms.ProjectPriceForm.Model;
using SoftLine.ActionPlugins.SharePoint;

namespace SoftLine.ActionPlugins.PrintForms
{
    public class PrintFormConstructor : IPrintFormConstructor<List<ProjectPrintForm>>
    {
        private readonly IOrganizationService _crmService;
        private readonly ISharePointClient _sharePointClient;
        public PrintFormConstructor(IOrganizationService crmService, ISharePointClient sharePointClient)
        {
            _crmService = crmService;
            _sharePointClient = sharePointClient;
        }
        public List<ProjectPrintForm> GetForms(InputPrintFormData inputData)
        {
            var projects = RetriveLinkProjectData(inputData);
            var printForms = new List<ProjectPrintForm>();
            return projects
                .GroupBy(x => x.Id)
                .Select(groupProjects =>
                {
                    var project = groupProjects.First();
                    var pictureUrl = project.GetValue<string>("picture.sl_url");
                    var pictureByte = string.IsNullOrEmpty(pictureUrl)
                        ? default
                        : _sharePointClient.GetFileByAbsoluteUrl(pictureUrl);

                    var pictureBase64 = pictureByte is null
                        ? default
                        : $"data:{MimeType.Jpeg};base64,{Convert.ToBase64String(pictureByte)}";

                    var metadata = RetriveMetadata(inputData.Language.Id);
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
                    var m2Title = getData(ProjectPriceMetadata.M2Title);
                    var vat = getData(ProjectPriceMetadata.VATHeader);
                    var cityLabel = getData(ProjectPriceMetadata.CityName);
                    var regionName = getData(ProjectPriceMetadata.RegionName);
                    var stageLabel = getData(ProjectPriceMetadata.StageName);
                    var toSeaName = getData(ProjectPriceMetadata.ToSeaName);
                    var priceFromName = getData(ProjectPriceMetadata.PriceFromName);
                    var completedConstructionName = getData(ProjectPriceMetadata.CompletedConstructionName);
                    var yesWord = getData(ProjectPriceMetadata.IBP);
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
                         return _crmService.RetrieveMultiple(new FetchExpression(query))
                         .Entities
                         .GroupBy(x => x.Id)
                         .ToDictionary(x => x.Key, y =>
                         {
                             var value = y.FirstOrDefault();
                             return value.GetValue<string>("ad.sl_name") ?? value.GetValue<string>("sl_name");
                         });
                     });

                    var emptyList = new List<string>();
                    List<string> flatValues(string attr, bool isNeeded = true)
                    {
                        if (!isNeeded) return emptyList;
                        return flats
                           .Select(x =>
                           {
                               return x.FormattedValues.Contains(attr)
                               ? x.FormattedValues[attr]
                               : x.Attributes.ContainsKey(attr) ? x[attr].ToString() : default;
                           })
                           .ToList();
                    }

                    var ibp = new Lazy<List<string>>(() =>
                    flats.Select(x => x.GetAttributeValue<bool>("sl_ibpbit") ? yesWord : string.Empty)
                    .ToList());

                    var vatsValue = isLtRent
                        ? Enumerable.Repeat(yesWord, flats.Count).ToList()
                        : emptyList;
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

                    var types = isType ? emptyList : flats.Select(x =>
                    {
                        var typeRef = x.GetAttributeValue<EntityReference>("sl_property_typeid");
                        if (Language.English.Equals(inputData.Language))
                            return typeRef?.Name;
                        return typeRef is null ? default : flatsType.Value[typeRef.Id];
                    });
                    var poolData = isPoolb ? GetPoolData(inputData, flats) : emptyList;
                    var minPrice = decimal.Zero;
                    var prices = isPrice ? RetriveListings(inputData, flats, getData, out minPrice) : emptyList;

                    var flatsWithPrice = flats
                        .Where(x => x.GetAttributeValue<OptionSetValue>("sl_sale_statuscode")?.Value == (int)ShortRentSTRentStatus.ForRent);

                    var flatColumns = new[]
                    {
                new FlatTableColumn("№", _empty, flatValues("sl_number")),
                new FlatTableColumn("IBP",_empty, isLtRent ? ibp.Value : emptyList, isLtRent),
                new FlatTableColumn("ID",_empty, flatValues("sl_id", isWithId), isWithId),
                new FlatTableColumn(floorLabel,_empty, flatValues("sl_floor", isFloor), isFloor),
                new FlatTableColumn(typeLabel,_empty, types , isType),
                new FlatTableColumn(bedLabel,_empty, flatValues("sl_number_of_bedrooms", isBed), isBed),
                new FlatTableColumn(innerTitle, m2Title,flatValues("sl_indoor", isIndoor), isIndoor),
                new FlatTableColumn(coveredVerandaTitle, m2Title, flatValues("sl_covered_veranda", isCovver), isCovver),
                new FlatTableColumn(openVerandaTitle,m2Title, flatValues("sl_open_veranda ", isUncovver), isUncovver),
                new FlatTableColumn(terasseTitle,m2Title, flatValues("sl_roof_terrace", isRoofTerrace), isRoofTerrace ),
                new FlatTableColumn(plotTitle,m2Title, flatValues("sl_plot_size ", isPlot), isPlot),
                new FlatTableColumn(generalUsageTitle,m2Title, flatValues("sl_common_area", isCommon), isCommon),
                new FlatTableColumn(storageTitle, m2Title, flatValues("sl_storage_size", isStorage),isStorage),
                new FlatTableColumn(totalAreaHeader, m2Title, flatValues("sl_total_area", isTotal), isTotal),
                new FlatTableColumn(poolTitle,_empty, poolData, isPoolb),
                new FlatTableColumn(parkinTitle,_empty, flatValues("sl_parking", isParking), isParking),
                new FlatTableColumn(priceHeader,_empty, prices, isPrice),
                new FlatTableColumn(vat, _empty, vatsValue , isLtRent)
             };

                    var columns = flatColumns
                        .Where(x => x.IsVisibility && x.Value.Length > 0)
                        .ToArray();

                    var rowsLenght = columns.Length == default ? default : columns.Max(x => x.Value.Length);
                    var cels = new string[columns.Length, rowsLenght];
                    for (var i = 0; i < columns.Length; i++)
                    {
                        var values = columns[i].Value;
                        for (var j = 0; j < rowsLenght; j++)
                        {
                            cels[i, j] = values[j];
                        }
                    }

                    var table = new List<List<string>>(columns.Length);
                    for (var i = 0; i < rowsLenght; i++)
                    {
                        var rows = new List<string>(rowsLenght);
                        for (var j = 0; j < columns.Length; j++)
                        {
                            rows.Add(cels[j, i]);
                        }
                        table.Add(rows);
                    }
                    var mainTable = new MainTable()
                    {
                        Flats = new List<List<List<string>>>()
                    };
                    mainTable.Flats.Add(table);
                    var minPriceStr = minPrice == default ? string.Empty : minPrice.ToString("0.##");
                    var printForm = new ProjectPrintForm()
                    {
                        ObjectName = project.GetAttributeValue<string>("sl_name"),
                        MajorBenefits = project.GetValue<string>("projectlanguage.sl_major_benefit") ?? project.GetAttributeValue<string>("sl_major_benefit"),
                        Details = new List<Detail>()
                 {
                    new Detail(cityLabel, project.GetValue<string>("city.sl_name")),
                    new Detail(regionName, project.GetValue<string>("area.sl_name")),
                    new Detail(stageLabel, project.FormattedValues.Contains("sl_stagecode") ? project.FormattedValues["sl_stagecode"] : default),
                    new Detail(toSeaName, $"{project.GetAttributeValue<decimal>("sl_distance_sea"):0.##} m"),
                    new Detail(priceFromName, minPriceStr),
                    new Detail(completedConstructionName, project.GetAttributeValue<DateTime?>("sl_completion_date")?.ToString("dd.MM.yyyy")),
                 },
                        Preview = pictureBase64,
                        Metadata = new ProjectPriceForm.Model.Metadata()
                        {
                            WithLogo = inputData.IsWithLogo,
                            MainTableColumns = columns
                            .Select(x => new MainTableColumn(x.Header, x.Dimension))
                            .ToList()
                        },
                        DetailsLabel = getData(ProjectPriceMetadata.DetailsLabel),
                        MajorBenefitsLabel = getData(ProjectPriceMetadata.MainBenefitsName),
                        PhoneLabel = getData(ProjectPriceMetadata.PhoneLabel),
                        FaxLabel = getData(ProjectPriceMetadata.FaxLabel),
                        EmailLabel = getData(ProjectPriceMetadata.EmailLabel),
                        Tagline = getData(ProjectPriceMetadata.Tagline),
                        MainTable = new List<MainTable>()
                 {
                    mainTable
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
                }).ToList();
        }

        private DataCollection<Entity> RetriveLinkProjectData(InputPrintFormData inputData)
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
            return _crmService.RetrieveMultiple(new FetchExpression(query))
                .Entities;
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

        private DataCollection<Entity> RetriveFlats(int? promotionType, Guid projectid)
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
            return _crmService.RetrieveMultiple(new FetchExpression(query)).Entities;
        }

        private Dictionary<string, Entity[]> RetriveMetadata(Guid languageid)
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
        }

        private List<string> GetPoolData(InputPrintFormData inputData, IEnumerable<Entity> flats)
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
            var translations = _crmService.RetrieveMultiple(new FetchExpression(query)).Entities;
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

        private List<string> RetriveListings(InputPrintFormData inputData, IEnumerable<Entity> flats, Func<ProjectPriceMetadata, string> getData, out decimal minPrice)
        {
            minPrice = 0;
            if (flats.Count() == 0) return new List<string>();
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

            var listings = _crmService.RetrieveMultiple(new FetchExpression(query)).Entities;
            var isPurchaseofproperty = inputData.PromotionType.Value == (int)PromotionType.PurchaseProperty;
            var isLtRent = inputData.PromotionType.Value == (int)PromotionType.LtRent;
            var priceSold = getData(ProjectPriceMetadata.Sold);
            var priceRented = getData(ProjectPriceMetadata.Rented);
            decimal? fromPrice = decimal.MaxValue;
            var prices = flats
                .Select(x =>
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
                            return startingPrice?.ToString("0.##");
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
                            return startingPrice?.ToString("0.##");
                        }
                    }
                    return default;
                }).ToList();
            minPrice = fromPrice == decimal.MaxValue ? default : fromPrice.Value;
            return prices;
        }
    }
}
