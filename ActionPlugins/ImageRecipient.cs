using Microsoft.SharePoint.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using SoftLine.ActionPlugins.Dto;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins
{
    public class ImageRecipient : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            try
            {
                var input = context.InputParameters;
                var regardingobjectRef = input["regardingobject"] as EntityReference;
                var format = input.Contains("format") ? input["format"] as Entity : default;
                var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                var service = serviceFactory.CreateOrganizationService(null);
                var pictures = RetrivePictures(regardingobjectRef, format, service);
                var groupUrls = pictures
                     .Where(x => x.GetAttributeValue<EntityReference>("sl_upload_formatid") is null)
                     .GroupBy(x => x.FormattedValues["sl_typecode"]);
                var images = GetFileAndDeleteFolderInSp(groupUrls, format, service);
                foreach (var pic in pictures.Where(x => x.GetAttributeValue<EntityReference>("sl_upload_formatid") != null))
                {
                    service.Delete(pic.LogicalName, pic.Id);
                }
                context.OutputParameters["responce"] = JsonConvert.SerializeObject(new { IsError = false, Images = images });
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

        public DataCollection<Entity> RetrivePictures(EntityReference entityReference, Entity format, IOrganizationService service)
        {
            var fieldName = GetReqardingField(entityReference.LogicalName);
            if (string.IsNullOrEmpty(fieldName))
            {
                throw new InvalidPluginExecutionException("Error in ImageRecipient");
            }
            int[] typecodesFilter;
            string condition = string.Empty;
            if (format is null)
            {
                typecodesFilter = new int[] { 102690000, 102690001 };
            }
            else
            {
                var typecode = format.GetAttributeValue<OptionSetValue>("sl_type")?.Value ?? 0;
                typecodesFilter = new int[] { typecode };
                condition = $@"<filter type='or' >
                      <condition attribute='sl_upload_formatid' operator='eq' value='{format.Id}' />
                      <condition attribute='sl_upload_formatid' operator='null' />
                    </filter>";
            }
            var query = $@"<fetch no-lock='true' >
                           <entity name='sl_picture' >
                             <attribute name='sl_pictureid' />
                             <attribute name='sl_primary_image' />
                             <attribute name='sl_url' />
                             <attribute name='sl_typecode' />
                             <attribute name='sl_upload_formatid' />
                             <filter>
                               <condition attribute='sl_typecode' operator='in' >
                                {string.Join("", typecodesFilter.Select(x => $"<value>{x}</value>"))}
                               </condition>
                             </filter>
                             <filter type='and' >
                               <condition attribute='{fieldName}' operator='eq' value='{entityReference.Id}' />
                               {condition}
                             </filter>
                           </entity>
                         </fetch>";
            return service.RetrieveMultiple(new FetchExpression(query)).Entities;
        }

        private string GetReqardingField(string regardingLogicalName)
        {
            return regardingLogicalName == "sl_unit"
               ? "sl_propertyid"
               : regardingLogicalName == "sl_project" ? "sl_projectid" : default;
        }

        public List<Images> GetFileAndDeleteFolderInSp(IEnumerable<IGrouping<string, Entity>> groupUrls, Entity format, IOrganizationService service)
        {
            var firstUrl = groupUrls.FirstOrDefault();
            if (firstUrl is null) return null;

            var uri = new Uri(firstUrl.First()["sl_url"] as string);
            var basePath = uri.GetLeftPart(UriPartial.Authority);
            var images = new List<Images>();
            using (var ctx = new ClientContext(basePath))
            {
                ctx.Credentials = Helper.GetСredentialsForSp(service);
                var web = ctx.Web;
                var folderForDelete = new List<string>();
                foreach (var gr in groupUrls)
                {
                    var image = new Images()
                    {
                        BaseFolder = gr.Key
                    };
                    var images64 = new List<ImageBase64>();
                    var isAddFolder = false;
                    foreach (var entity in gr)
                    {
                        var url = entity["sl_url"] as string;
                        var file = web.GetFileByUrl(url);
                        ctx.Load(file,
                            i => i.Exists,
                            u => u.Name,
                            u => u.ListItemAllFields["FileDirRef"]);

                        ctx.ExecuteQuery();
                        if (!file.Exists) continue;
                        if (!isAddFolder)
                        {
                            var fileDirRef = file.ListItemAllFields["FileDirRef"] as string;
                            folderForDelete.Add(fileDirRef);
                            isAddFolder = true;
                        }
                        var stream = file.OpenBinaryStream();
                        ctx.ExecuteQuery();
                        using (var mStream = new MemoryStream())
                        {
                            stream.Value.CopyTo(mStream);
                            var imageArray = mStream.ToArray();
                            var image64 = new ImageBase64()
                            {
                                Base64 = $"data:{MimeType.Jpeg};base64," + Convert.ToBase64String(imageArray),
                                MimeType = MimeType.Jpeg,
                                FullName = file.Name
                            };
                            images64.Add(image64);
                        }
                    }
                    image.Base = images64;
                    images.Add(image);
                }
                DeleteFolderInSp(ctx, format, folderForDelete);
            }
            return images;
        }

        private void DeleteFolderInSp(ClientContext ctx, Entity format, IEnumerable<string> folderUrls)
        {
            foreach (var url in folderUrls)
            {
                var parentFolder = ctx.Web.GetFolderByServerRelativeUrl(url);
                var folders = parentFolder.Folders;
                if (format is null)
                {
                    ctx.Load(folders);
                }
                else
                {
                    var folderName = format.GetAttributeValue<string>("sl_name");
                    ctx.Load(folders,
                            x => x.Where(y => y.Name == folderName));
                }
                ctx.ExecuteQuery();
                foreach (var f in folders.ToArray())
                {
                    f.DeleteObject();
                }
                ctx.ExecuteQuery();
            }
        }
    }
}
