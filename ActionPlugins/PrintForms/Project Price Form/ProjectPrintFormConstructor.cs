using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SoftLine.ActionPlugins.Extensions;
using SoftLine.ActionPlugins.Languages;
using SoftLine.ActionPlugins.OptionSets;
using SoftLine.ActionPlugins.PrintForms;
using SoftLine.ActionPlugins.PrintForms.Metadata;
using SoftLine.ActionPlugins.PrintForms.Project_Price_Form.Model;
using SoftLine.ActionPlugins.PrintForms.ProjectPriceForm.Model;
using SoftLine.ActionPlugins.SharePoint;

namespace SoftLine.ActionPlugins.ProjectPriceForm.PrintForms
{
    public class ProjectPrintFormConstructor : FormConstructor<List<ProjectPrintForm>>
    {
        public ProjectPrintFormConstructor(IOrganizationService crmService, ISharePointClient sharePointClient) : base(crmService, sharePointClient)
        {
        }

        public override List<ProjectPrintForm> GetForms(InputPrintFormData inputData)
        {
            var projects = RetriveLinkProjectData(inputData);
            var metadata = RetriveMetadata(inputData.Language.Id);
            return projects
                .GroupBy(x => x.Id)
                .Select(groupProjects =>
                {
                    var project = groupProjects.First();
                    return GetForm(inputData, project, metadata);
                })
                .ToList();
        }

        private ProjectPrintForm GetForm(InputPrintFormData inputData, Entity project, Dictionary<string, Entity[]> metadata)
        {
            string getData(FormMetadata priceMetadata)
            {
                var data = metadata[priceMetadata.StringValue()].FirstOrDefault();
                return data.GetValue<string>("translation.sl_name")
                    ?? data.GetAttributeValue<string>("sl_word");
            }
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

            var emptyList = new List<string>();
            List<string> flatValues(string attr, bool isNeeded = true)
            {
                if (!isNeeded) return emptyList;
                return flats
                   .Select(x =>
                   {
                       return x.FormattedValues.Contains(attr)
                       ? x.FormattedValues[attr]
                       : x.Contains(attr) ? x[attr].ToString() : default;
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
            var prices = isPrice ? RetriveListings(inputData, flats, getData, out minPrice).Values.ToList() : emptyList;

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
            var minPriceStr = minPrice == default ? string.Empty : minPrice.ToString("N", new CultureInfo("en-US"));
            var majorBenefits = project.GetValue<string>("projectlanguage.sl_major_benefit") ?? project.GetAttributeValue<string>("sl_major_benefit");
            var printForm = new ProjectPrintForm()
            {
                ObjectName = project.GetAttributeValue<string>("sl_name"),
                MajorBenefits = FindFourParagraph(majorBenefits),
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
                Metadata = new Metadata
                {
                    WithLogo = inputData.IsWithLogo,
                    MainTableColumns = columns
                    .Select(x => new MainTableColumn(x.Header, x.Dimension))
                    .ToList()
                },
                DetailsLabel = getData(FormMetadata.DetailsLabel),
                MajorBenefitsLabel = getData(FormMetadata.MainBenefitsName),
                PhoneLabel = getData(FormMetadata.PhoneLabel),
                FaxLabel = getData(FormMetadata.FaxLabel),
                EmailLabel = getData(FormMetadata.EmailLabel),
                Tagline = getData(FormMetadata.Tagline),
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
        }
    }
}
