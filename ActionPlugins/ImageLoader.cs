using Microsoft.SharePoint.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using SoftLine.ActionPlugins.Dto;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using SoftLine.ActionPlugins.OptionSets;

namespace SoftLine.ActionPlugins
{
    public class ImageLoader : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            try
            {              
                var input = context.InputParameters;
                var images = input["images"] as string;
                var regardingobjectRef = input["regardingobject"] as EntityReference;

                var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                var service = serviceFactory.CreateOrganizationService(null);
                var documentsFolderUri = RetriveAbsoluteUriFolder(regardingobjectRef, service);

                var images64 = JsonConvert.DeserializeObject<Images>(images);
                UploadDocument(documentsFolderUri, images64, regardingobjectRef, service);
                context.OutputParameters["responce"] = JsonConvert.SerializeObject(new { IsError = false, Message = $"Files uploaded." });
            }
            catch (Exception ex)
            {
                context.OutputParameters["responce"] = JsonConvert.SerializeObject(new
                {
                    IsError = true,
                    Message = $"{ex.Message}\n{ex.InnerException?.Message}\n{ex.StackTrace}"
                });
            }

        }

        public void UploadDocument(Uri folderUri, Images images, EntityReference regardingobjectRef, IOrganizationService service)
        {
            var basePath = folderUri.GetLeftPart(UriPartial.Authority);
            var spData = Helper.GetInputDataForSp(service);

            byte[] getByteBase64(string base64)
            {
                var splitBase64 = base64.Split(',');
                var base64WithouotType = splitBase64.Length > 1 ? splitBase64[1] : splitBase64[0];
                return Convert.FromBase64String(base64WithouotType);
            }
            using (var ctx = new ClientContext(basePath))
            {
                ctx.Credentials = spData.Credentials;
                var web = ctx.Web;
                var folder = web.GetFolderByServerRelativeUrl($"{folderUri.AbsolutePath}");
                ctx.Load(folder);
                ctx.ExecuteQuery();
                var parentFolder = CreateFolder(images.BaseFolder, folder, web);
                var metadataTypeCode = RetriveMetadata("sl_picture", "sl_typecode", service);
                var typecode = GetValueByLabel(metadataTypeCode, images.BaseFolder);
                if (images.Base != null)
                {
                    foreach (var image64 in images.Base)
                    {
                        var imageBytes = getByteBase64(image64.Base64);
                        using (var stream = new MemoryStream(imageBytes))
                        {
                            var existPicture = RetriveExistPicture(image64.FullName, regardingobjectRef, image64.Formatid, service);
                            if (existPicture is null)
                            {
                                var path = $"{parentFolder.ServerRelativeUrl}/{image64.FullName}";
                                Microsoft.SharePoint.Client.File.SaveBinaryDirect(ctx, path, stream, true);
                                var fullPath = $@"{basePath}{path}";

                                var picture = FormPicture(image64, fullPath, imageBytes, regardingobjectRef);
                                picture["sl_typecode"] = typecode;
                                if (typecode?.Value == (int)UploadFormatTyprCode.Construction_progress)
                                {
                                    picture["sl_consrtuction_date"] = image64.LastModifiedDate;
                                }
                                service.Create(picture);
                            }
                        }
                    }
                }

                if (images.Resize?.Count() == 0) return;

                var groupByFolder = images
                    .Resize
                    .GroupBy(x => x.FolderName);

                var primaryImageFolderUrl = $@"{basePath}/{parentFolder.ServerRelativeUrl}";
                foreach (var g in groupByFolder)
                {
                    var folderName = g.Key;
                    var subFolder = CreateFolder(folderName, parentFolder, web);
                    foreach (var resizeImage in g)
                    {
                        var imageBytes = getByteBase64(resizeImage.Base64);
                        using (var stream = new MemoryStream(imageBytes))
                        {
                            var existPicture = RetriveExistPicture(resizeImage.FullName, regardingobjectRef, resizeImage.Formatid, service);
                            if (existPicture is null)
                            {
                                var path = $"{subFolder.ServerRelativeUrl}/{resizeImage.NameForUrl}";
                                Microsoft.SharePoint.Client.File.SaveBinaryDirect(ctx, path, stream, true);
                                var fullPath = $@"{basePath}{path}";

                                var picture = FormPicture(resizeImage, fullPath, imageBytes, regardingobjectRef);
                                picture["sl_typecode"] = typecode;
                                picture["sl_primary_image"] = $"{primaryImageFolderUrl}/{resizeImage.FullName}";
                                service.Create(picture);
                            }                            
                        }
                    }
                }
            }
        }

        private Entity RetriveExistPicture(string name, EntityReference regardingobjectRef, Guid? formatid, IOrganizationService service)
        {
            var field = GetReqardingField(regardingobjectRef.LogicalName);
            var formatCondition = !formatid.HasValue
                ? "<condition attribute='sl_upload_formatid' operator='null' />"
                : $"<condition attribute='sl_upload_formatid' operator='eq' value='{formatid}' />";
            var query = $@"<fetch top='1' no-lock='true' >
                           <entity name='sl_picture' >
                             <attribute name='sl_pictureid' />
                             <filter type='and' >
                               {formatCondition}
                               <condition attribute='sl_name' operator='eq' value='{name}' />
                               <condition attribute='{field}' operator='eq'  value='{regardingobjectRef.Id}' />
                             </filter>
                           </entity>
                         </fetch>";

            return service.RetrieveMultiple(new FetchExpression(query)).Entities.FirstOrDefault();
        }

        private Entity FormPicture(ImageBase64 image, string url, byte[] imageByte, EntityReference regardingobjectRef)
        {
            var size = Convert.ToDecimal(imageByte.Length / 1024);
            var picture = new Entity("sl_picture")
            {
                ["sl_name"] = image.FullName,
                ["sl_height"] = image.Height == default ? null : (decimal?)image.Height,
                ["sl_width"] = image.Width == default ? null : (decimal?)image.Width,
                ["sl_upload_formatid"] = image.Formatid.HasValue ? new EntityReference("sl_upload_format", image.Formatid.Value) : null,
                ["sl_image"] = image.MimeType == MimeType.Jpeg ? imageByte : null,
                ["sl_url"] = url,
                ["sl_size"] = size
            };
            var fieldName = GetReqardingField(regardingobjectRef.LogicalName);
            if (!string.IsNullOrEmpty(fieldName))
            {
                picture[fieldName] = regardingobjectRef;
            }
            return picture;
        }

        private Folder CreateFolder(string fullFolderUrl, Folder parentFolder, Web web)
        {
            var folderUrls = fullFolderUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string folderUrl = folderUrls[0];
            var curFolder = parentFolder.Folders.Add(folderUrl);
            web.Context.Load(curFolder);
            web.Context.ExecuteQuery();
            if (folderUrls.Length > 1)
            {
                var subFolderUrl = string.Join("/", folderUrls, 1, folderUrls.Length - 1);
                return CreateFolder(subFolderUrl, curFolder, web);
            }
            return curFolder;
        }

        public Uri RetriveAbsoluteUriFolder(EntityReference regardingobjectRef, IOrganizationService service)
        {
            var logicalName = regardingobjectRef.LogicalName;
            var fetch = $@"<fetch top='1' no-lock='true' >
                           <entity name='sharepointdocumentlocation' >
                             <attribute name='parentsiteorlocation' />
                             <attribute name='relativeurl' />
                             <filter type='and' >
                               <condition attribute='locationtype' operator='eq' value='0' />
                               <condition attribute='servicetype' operator='eq' value='0' />
                             </filter>
                             <link-entity name='{logicalName}' from='{logicalName}id' to='regardingobjectid' link-type='inner' alias='ac' >
                               <filter type='and' >
                                 <condition attribute='{logicalName}id' operator='eq' value='{regardingobjectRef.Id}' />
                               </filter>
                             </link-entity>
                           </entity>
                         </fetch>";
            var location = service.RetrieveMultiple(new FetchExpression(fetch)).Entities.FirstOrDefault();
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
                    var sharepointsite = service.Retrieve(parentsiteorlocationRef.LogicalName, parentsiteorlocationRef.Id, new ColumnSet("absoluteurl"));
                    var sharepointsiteUrl = sharepointsite.GetAttributeValue<string>("absoluteurl");
                    stack.Push(sharepointsiteUrl);
                    break;
                }
                else
                {
                    location = service.Retrieve(parentsiteorlocationRef.LogicalName, parentsiteorlocationRef.Id, new ColumnSet("relativeurl", "parentsiteorlocation"));
                }
                i++;
            }
            var url = string.Join(@"/", stack);
            if (string.IsNullOrEmpty(url))
            {
                throw new InvalidPluginExecutionException($"У {regardingobjectRef.Name} не создана папка расположения документов");
            }
            return new Uri(url);
        }

        private EnumAttributeMetadata RetriveMetadata(string entityLogicalName, string label, IOrganizationService service)
        {
            var attributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityLogicalName,
                LogicalName = label,
                RetrieveAsIfPublished = true,

            };

            var attributeResponse = (RetrieveAttributeResponse)service.Execute(attributeRequest);
            return (EnumAttributeMetadata)attributeResponse.AttributeMetadata;
        }

        private OptionSetValue GetValueByLabel(EnumAttributeMetadata attributeMetadata, string label)
        {
            var value = attributeMetadata.OptionSet.Options
               .Where(x => x.Label.UserLocalizedLabel.Label == label)
               .Select(x => x.Value)
               .FirstOrDefault();

            return value is null ? null : new OptionSetValue(value.Value);
        }

        private string GetReqardingField(string regardingLogicalName)
        {
            return regardingLogicalName == "sl_unit"
               ? "sl_propertyid"
               : regardingLogicalName == "sl_project" ? "sl_projectid" : default;
        }
    }
}
