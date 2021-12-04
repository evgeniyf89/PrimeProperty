using Microsoft.SharePoint.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SofLine.Picture
{
    /// <summary>
    ///  Удаление Picture
    /// </summary>
    /// <remarks>
    /// <para>
    /// <code title="Настройка">
    /// Step
    /// Message:                Delete
    /// Primary Entity:         pl_picture

    /// Name:                   SofLine.Picture.DeletePicture
    /// Run in User's Context:  Calling User
    /// Execution Order:        1
    /// Description:            
    /// Stage:                  Post-operation
    /// Execution Mode:         Async
    /// PreImage :              sl_upload_formatid, sl_url, sl_name, sl_projectid, sl_propertyid
    /// Unsecure Configuration:     
    /// </code>
    /// </para>
    /// </remarks>
    public class DeletePicture : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(null);
       
            var preImage = context.PreEntityImages["Image"];
            var uploadFormatid = preImage.GetAttributeValue<EntityReference>("sl_upload_formatid");
            if (uploadFormatid != default)
            {
                var url = preImage.GetAttributeValue<string>("sl_url");
                if (string.IsNullOrEmpty(url)) return;
                DeleteFileFromSp(service, url);
                return;
            }           
            var basePictureUrl = preImage.GetAttributeValue<string>("sl_url");            
            var resizeImage = RetriveResizeImage(preImage, service);
            DeleteResizePicture(resizeImage, service);
            var allUrls = resizeImage
                .Select(x => x.GetAttributeValue<string>("sl_url"))
                .Concat(new[] { basePictureUrl })
                .ToArray();
            DeleteFileFromSp(service, allUrls);           
        }

        private void DeleteResizePicture(IEnumerable<Entity> resizeImage, IOrganizationService service)
        {            
            foreach(var image in resizeImage)
            {
                service.Delete(image.LogicalName, image.Id);
            }
        }

        private DataCollection<Entity> RetriveResizeImage(Entity preImage, IOrganizationService service)
        {
            var baseName = preImage.GetAttributeValue<string>("sl_name");
            var regarding = preImage.GetAttributeValue<EntityReference>("sl_projectid")
                ?? preImage.GetAttributeValue<EntityReference>("sl_propertyid");
            var fieldName = GetReqardingField(regarding.LogicalName);

            var query = $@"<fetch no-lock='true' >
                           <entity name='sl_picture' >
                             <attribute name='sl_url' />
                             <filter type='and' >
                               <condition attribute='{fieldName}' operator='eq' value='{regarding.Id}' />
                               <condition attribute='sl_typecode' operator='in' >
                                 <value>102690000</value>
                                 <value>102690001</value>
                               </condition>
                               <condition attribute='sl_name' operator='eq' value='{baseName}' />
                               <condition attribute='sl_primary_image' operator='not-null' />
                               <condition attribute='sl_upload_formatid' operator='not-null' />
                               <condition attribute='sl_url' operator='not-null' />
                             </filter>
                           </entity>
                         </fetch>";
            return service.RetrieveMultiple(new FetchExpression(query)).Entities;
        }

        public void DeleteFileFromSp(IOrganizationService service, params string[] urls)
        {
            if (urls is null || urls.Length == 0) return;
            var uri = new Uri(urls[0]);
            var basePath = uri.GetLeftPart(UriPartial.Authority);
            using (var ctx = new ClientContext(basePath))
            {
                ctx.Credentials = GetСredentialsForSp(service);
                foreach (var url in urls)
                {
                    var file = ctx.Web.GetFileByUrl(url);
                    ctx.Load(file, i => i.Exists);
                    ctx.ExecuteQuery();
                    if (!file.Exists) continue;
                    file.DeleteObject();
                }
                ctx.ExecuteQuery();
            }
        }

        private ICredentials GetСredentialsForSp(IOrganizationService service)
        {
            var query = $@"<fetch no-lock='true' >
                           <entity name='sl_setting' >
                             <attribute name='sl_settingid' />
                             <attribute name='sl_key' />
                             <attribute name='sl_value' />
                             <order attribute='sl_key' descending='false' />
                             <filter type='and' >
                               <condition attribute='sl_key' operator='in' >
                                 <value>SpPassword</value>
                                 <value>SpUser</value>
                               </condition>
                             </filter>
                           </entity>
                         </fetch>";
            var settings = service.RetrieveMultiple(new FetchExpression(query)).Entities;
            var passwordString = settings
                .First(x => x.GetAttributeValue<string>("sl_key") == "SpPassword")
                .GetAttributeValue<string>("sl_value");

            var userName = settings
                .First(x => x.GetAttributeValue<string>("sl_key") == "SpUser")
                .GetAttributeValue<string>("sl_value");

            var passWord = new SecureString();
            Array.ForEach(passwordString.ToCharArray(), ch => passWord.AppendChar(ch));
            return new SharePointOnlineCredentials(userName, passWord);
        }

        private string GetReqardingField(string regardingLogicalName)
        {
            return regardingLogicalName == "sl_unit"
               ? "sl_propertyid"
               : regardingLogicalName == "sl_project" ? "sl_projectid" : default;
        }
    }
}

