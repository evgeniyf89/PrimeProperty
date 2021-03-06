using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using SofLine.Picture;
using SoftLine.ActionPlugins;
using SoftLine.ActionPlugins.Dto;
//using SoftLine.ActionPlugins.Dto;
using SoftLine.ActionPlugins.OptionSets;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace SoftLine.ActionPlugins.Tests
{
    [TestClass()]
    public class ImageLoaderTests
    {
        private readonly IOrganizationService _service;
        public ImageLoaderTests()
        {
            var crmSoap = "https://ppcrm1.api.crm4.dynamics.com/XRMServices/2011/Organization.svc";
            var uri = new Uri(crmSoap);
            var credentials = new ClientCredentials();
            credentials.UserName.Password = "Nhgb432!";
            credentials.UserName.UserName = "test.crm@prime-property.com";
            ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            _service = new OrganizationServiceProxy(uri, null, credentials, null);

        }
        [TestMethod()]
        public void ExecuteTest()
        {
            var uriFolder = RetriveAbsoluteUriFolder();
            var image = ImageToBase64();
            var stream = Base64ToImage(image);
            using (var ctx = new ClientContext($@"{uriFolder.GetLeftPart(UriPartial.Authority)}"))
            {
                var web = ctx.Web;
                var passWord = new SecureString();
                var passwordString = "Prime_2021";
                Array.ForEach(passwordString.ToCharArray(), ch => passWord.AppendChar(ch));
                ctx.Credentials = new SharePointOnlineCredentials("test.crm@prime-property.com", passWord);
                var folder = web.GetFolderByServerRelativeUrl($"{uriFolder.AbsolutePath}");
                var folders = folder.Folders;
                //ctx.Load(folder, f => f.Name);
                ctx.Load(folders);
                ctx.ExecuteQuery();
                folders.Add("Images");
                ctx.ExecuteQuery();
                Microsoft.SharePoint.Client.File.SaveBinaryDirect(ctx, $"{uriFolder.AbsolutePath}/AAA.jpg", stream, false);
                //var i = files.Count;
            }
        }

        [TestMethod()]
        public void FromJson()
        {
            var json = "[{\"Name\":\"5j5hIf24eM.jpg\",\"FolderName\":\"Folder\",\"Base64\":\"asdh\"}]";


            var serializer = new JavaScriptSerializer();
            var images64 = serializer.Deserialize<ImageBase64[]>(json);

            var t = "[102690000]";

            var json2 = serializer.Deserialize<int[]>(t);
            //var folder = Enum.GetName(typeof(Formats), json2[0]);
        }

        [TestMethod()]
        public void RetriveFormatsMetadata()
        {
            var attributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = "sl_upload_format",
                LogicalName = "sl_type",
                RetrieveAsIfPublished = true
            };

            var attributeResponse = (RetrieveAttributeResponse)_service.Execute(attributeRequest);
            var attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

            var data = attributeMetadata
                .OptionSet
                .Options
                .ToDictionary(x => x.Value, y => y.Label.UserLocalizedLabel.Label);

        }

        [TestMethod()]
        public void RetrivePictures()
        {

            var tt = new ImageRecipient();
            var obj = new EntityReference("sl_project", new Guid("{30824c04-2cdb-eb11-bacb-000d3a2c32dc}"));
            var formatRef = new Entity("sl_upload_format", new Guid("{5A616E8C-36E6-EB11-BACB-000D3A470D6F}"));
            formatRef["sl_type"] = new OptionSetValue(102690000);
            var pictures = tt.RetrivePictures(obj, formatRef, _service);

            var groupUrls = pictures
                    .Where(x => x.GetAttributeValue<EntityReference>("sl_upload_formatid") is null)
                    .GroupBy(x => x.FormattedValues["sl_typecode"])
                    .ToArray();
            var images = tt.GetFileAndDeleteFolderInSp(groupUrls, formatRef, _service);
            foreach (var pic in pictures.Where(x => x.GetAttributeValue<EntityReference>("sl_upload_formatid") != null))
            {
                _service.Delete(pic.LogicalName, pic.Id);
            }
            var json = JsonConvert.SerializeObject(images);
        }

        public static Stream Base64ToImage(string base64String)
        {
            var imageBytes = Convert.FromBase64String(base64String);
            var ms = new MemoryStream(imageBytes, 0, imageBytes.Length);
            return ms;
        }

        public static string ImageToBase64()
        {
            var desktop = @"C:\Users\Fedotoveni\Desktop";
            var path = Path.Combine(desktop, "Вакцинация.jpg");
            using (Image image = Image.FromFile(path))
            {
                using (MemoryStream m = new MemoryStream())
                {
                    image.Save(m, image.RawFormat);
                    byte[] imageBytes = m.ToArray();
                    string base64String = Convert.ToBase64String(imageBytes);
                    return base64String;
                }
            }
        }

        private Uri RetriveAbsoluteUriFolder()
        {
            var fetch = $@"<fetch top='1' no-lock='true' >
                           <entity name='sharepointdocumentlocation' >
                             <attribute name='parentsiteorlocation' />
                             <attribute name='relativeurl' />
                             <filter type='and' >
                               <condition attribute='locationtype' operator='eq' value='0' />
                               <condition attribute='servicetype' operator='eq' value='0' />
                             </filter>
                             <link-entity name='sl_project' from='sl_projectid' to='regardingobjectid' link-type='inner' alias='ac' >
                               <filter type='and' >
                                 <condition attribute='sl_projectid' operator='eq' value='30824C04-2CDB-EB11-BACB-000D3A2C32DC' />
                               </filter>
                             </link-entity>
                           </entity>
                         </fetch>";
            var location = _service.RetrieveMultiple(new FetchExpression(fetch)).Entities.FirstOrDefault();
            var stack = new Stack<string>();
            var i = 0;
            while (i < 5)
            {
                var relativeurl = location.GetAttributeValue<string>("relativeurl");
                stack.Push(relativeurl);
                var parentsiteorlocationRef = location.GetAttributeValue<EntityReference>("parentsiteorlocation");
                if (parentsiteorlocationRef is null)
                    break;
                else if (parentsiteorlocationRef.LogicalName == "sharepointsite")
                {
                    var sharepointsite = _service.Retrieve(parentsiteorlocationRef.LogicalName, parentsiteorlocationRef.Id, new ColumnSet("absoluteurl"));
                    var sharepointsiteUrl = sharepointsite.GetAttributeValue<string>("absoluteurl");
                    stack.Push(sharepointsiteUrl);
                    break;
                }
                else
                {
                    location = _service.Retrieve(parentsiteorlocationRef.LogicalName, parentsiteorlocationRef.Id, new ColumnSet("relativeurl", "parentsiteorlocation"));
                }
                i++;
            }
            var url = string.Join(@"/", stack);
            return new Uri(url);
        }

        [TestMethod()]
        public void UploadDocumentTest()
        {

            var loader = new ImageLoader();

            var image64 = ImageToBase64();
            var images = new Images()
            {
                BaseFolder = "Images",
                Base = new[] { new ImageBase64()
                {
                    Base64 = image64,
                    FullName = "7F1OM2dU0a.jpg",
                    FolderName = "",
                    Height = 820,
                    Width = 1000,
                },
                },
                Resize = new[] { new ImageBase64()
                {
                    Base64 = image64,
                    FullName = "7F1OM2dU0aGG.jpg",
                    FolderName = "768x576_60",
                    Height = 768,
                    Width = 576,
                    Formatid = new Guid("7cb6c24d-36e6-eb11-bacb-000d3a470d6f")
                },new ImageBase64()
                {
                    Base64 = image64,
                    FullName = "7F1OM2dU0aKKKK.jpg",
                    FolderName = "150x90_60",
                    Height = 150,
                    Width = 90,
                    Formatid = new Guid("1808f71a-36e6-eb11-bacb-000d3a470d6f")
                },
                },
            };
            var refer = new EntityReference("sl_unit", new Guid("{DD8C3E5A-8D09-EC11-B6E6-00224882AB04}"));
            var uriFolder = loader.RetriveAbsoluteUriFolder(refer, _service);
            var list = new List<ImageBase64>();

            loader.UploadDocument(uriFolder, images, refer, _service);
           
        }


        [TestMethod()]
        public void Calendar()
        {
            var calendar = new CalendarLogic();
            var prop = new EntityReference("sl_property_for_opportunity", new Guid("3e132c12-5f2b-ec11-b6e5-000d3a268060"));
            var start = new DateTime(2021, 12, 5);
            var end = new DateTime(2021, 12, 15);
            var rents = calendar.RetriveRent(prop, start, end, _service);
            var obj = calendar.Map(rents);
            var t = JsonConvert.SerializeObject(obj);
        }

        [TestMethod()]
        public void RentPriceLogic()
        {
            var logic = new RentPriceLogic();
            var startDate = new DateTime(2022, 01, 11);
            var endDate = new DateTime(2022, 01, 20);
            var refer = new EntityReference("sl_listing", new Guid("{4B0C2DCA-482B-EC11-B6E5-000D3A268060}"));
            var tt = logic.RetriveRentPrice(refer, startDate, endDate, _service);
            var ranges = logic.Map(startDate, endDate, tt);
            var kk = JsonConvert.SerializeObject(ranges);
        }

        [TestMethod()]
        public void ReserveOrRentProperty()
        {
            var data = new EntityReference("sl_property_for_opportunity", new Guid("{5430D489-7752-EC11-8C62-000D3ABFC59F}"));
            var calendar = new CalendarLogic();
            var startDate = new DateTime(2022, 1, 23);
            var endDate = new DateTime(2022, 1, 31);
            var status = ShortRentSTRentSstatus.Reserved;

            var rents = calendar.RetriveRent(data, startDate, endDate, _service);
            var obj = calendar.Map(rents);
            if (obj.RentedByOpportunity.Any() || obj.ReservedByOpportunity.Any())
            {
                 JsonConvert.SerializeObject(new { IsError = true, Message = $"The number of rental days is too small." });
                return;
            };
            var rangeDays = (endDate - startDate).Days;
            if (rangeDays < obj.MinDays)
            {
                JsonConvert.SerializeObject(new { IsError = true, Message = $"Dates are not available for selection." });
                return;
            };
            var yy = new ReserveOrRentProperty();
            var propertyOpportunity = yy.RetrivePropertyOpportunity(data, _service);
            var listing = propertyOpportunity.GetAttributeValue<EntityReference>("sl_listingid");
            var opportunityRef = propertyOpportunity.GetAttributeValue<EntityReference>("sl_opportunityid");
            if (listing is null)
            {
                JsonConvert.SerializeObject(new { IsError = true, Message = $"listing is empty" });
                return;
            }
            var priceLogic = new RentPriceLogic();
            var from = startDate.AddDays(1);
            var to = endDate.AddDays(1);
            var rentPrices = priceLogic.RetriveRentPrice(listing, from, to, _service);
            var rentPricesObj = priceLogic.Map(from, to, rentPrices);
            var rentalFree = rentPricesObj.Sum(x => x.Price);

            var rented = obj.Rented.FirstOrDefault();
            var reserved = obj.Reserved.FirstOrDefault();

            if (rented != default || reserved != default)
            {
                var rentavailableid = rented?.Id ?? reserved.Id;
                yy.UpdateRentavailable(status, startDate, endDate, rentavailableid, _service);
                yy.UpdateOpportynity(status, startDate, endDate, opportunityRef, _service);
                yy.UpdatePropertyOpportunity(status, startDate, endDate, rentalFree, data, _service);
                return;
            }
            else
            {
                yy.CreateRentavailable(status, startDate, endDate, propertyOpportunity, _service);
            }
        }

    }
}