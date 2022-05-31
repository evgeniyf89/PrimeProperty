using System;
using Microsoft.Xrm.Sdk;
using SoftLine.ActionPlugins.SharePoint;
using Newtonsoft.Json;
using SoftLine.ActionPlugins.PrintForms;
using SoftLine.ActionPlugins.PrintForms.ProjectPriceForm.Model;

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
                var inputData = new InputPrintFormData()
                {
                    IsWithLogo = (bool)input["isWithLogo"],
                    TargetEntityRef = input["data"] as EntityReference,
                    PromotionType = input["promotionType"] as OptionSetValue,
                    Language = input["language"] as EntityReference,
                    Market = input["market"] as OptionSetValue,
                };

                var printForm = FormPrintForm(inputData, service);
               
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

        public ProjectPrintForm FormPrintForm(InputPrintFormData inputData, IOrganizationService service)
        {
            var spData = Helper.GetInputDataForSp(service);
            using (var spClient = new SharePointClient(spData.Url, spData.Credentials))
            {
                var constructor = new PrintFormConstructor(service, spClient);
                return constructor.GetForm(inputData);
            }
        }
    }
}
