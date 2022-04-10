using System;
using Microsoft.Xrm.Sdk;
using SoftLine.ActionPlugins.WordTemplateConstructor;
using SoftLine.ActionPlugins.SharePoint;
using Newtonsoft.Json;

namespace SoftLine.ActionPlugins
{
    public class CreatePrintedForm : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(null);
            var input = context.InputParameters;
            try
            {
                var entityRef = input["data"] as EntityReference;
                var templateid = (int)input["templateid"];
                var printForm = FormPrintForm(templateid, entityRef.Id, entityRef.Id, service);
                context.OutputParameters["responce"] = JsonConvert.SerializeObject(printForm);
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

        public object FormPrintForm(int templateid, Guid entityid, Guid languageId, IOrganizationService service)
        {
            var spData = Helper.GetInputDataForSp(service);
            using (var spClient = new SharePointClient(spData.Url, spData.Credentials))
            {
                var parametersContstructor = new ParametersContstructor(service, spClient);
                var settings = parametersContstructor.GetPrintFormSettings(templateid, entityid, languageId);

                var url = $"{settings.TemplateFolderInSharePoint}/{settings.TemplateNameInSharePoint}";

                var templateWord = spClient.GetFileByAbsoluteUrl(url);
                var wordFormater = new WordDocumentFormater();
                var printForm = wordFormater.Form(templateWord, settings);
                return new { IsError = false, FileName = settings.TemplateNameInSharePoint, File = Convert.ToBase64String(printForm) };
            }
        }
    }
}
