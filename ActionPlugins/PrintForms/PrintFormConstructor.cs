using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SoftLine.ActionPlugins.Extensions;
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

            var printForm = new ProjectPrintForm()
            {
                ObjectName = project.GetAttributeValue<string>("sl_name"),
                MajorBenefits = project.GetAttributeValue<string>("sl_major_benefit"),
                Details = new Details()
                {
                    DistanceToSea = $"{project.GetAttributeValue<decimal>("sl_distance_sea")} m",
                    City = project.GetValue<string>("city.sl_name"),
                    Area = project.GetValue<string>("area.sl_name"),
                    Completion = project.FormattedValues.Contains("sl_stagecode") ? project.FormattedValues["sl_stagecode"] : default,
                    CompletionDate = project.GetAttributeValue<DateTime?>("sl_completion_date")?.ToString("dd.MM.yyyy"),
                    PricesFrom = @"(11)	ЦенаОт – одинакова для всех языков
(a)	Minimum from sl_listing. sl_sale_starting_price, для @Promotion type = Purchase of property (102 690 000) Сущность Listing
(i)	sl_listing. sl_clients_promotion_typecode=@Promotion type
(ii)	sl_listing. sl_marketcode=@Market
(iii)	sl_listing. sl_unitid  (sl_unit.sl_projectid == выбранный проект из представления)
(b)	Minimum from sl_listing. sl_long_rent, для @Promotion type = LT Rent (102 690 001) Сущность Listing
(i)	sl_listing. sl_clients_promotion_typecode=@Promotion type
(ii)	sl_listing. sl_marketcode=@Market
(iii)	sl_listing. sl_unitid  (sl_unit.sl_projectid == выбранный проект из представления)
"
                },
                Preview = pictureBase64,
                Metadata = new Metadata()
                {
                    WithLogo = inputData.IsWithLogo
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
    }
}
