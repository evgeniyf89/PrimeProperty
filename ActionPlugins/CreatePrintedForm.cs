using System;
using Microsoft.Xrm.Sdk;
using SoftLine.ActionPlugins.SharePoint;
using Newtonsoft.Json;
using SoftLine.ActionPlugins.PrintForms;
using SoftLine.ActionPlugins.PrintForms.ProjectPriceForm.Model;
using System.Collections.Generic;
using SoftLine.ActionPlugins.ProjectPriceForm.PrintForms;
using SoftLine.ActionPlugins.PrintForms.Metadata;
using SoftLine.ActionPlugins.PrintForms.MasterPrice;
using SoftLine.ActionPlugins.PrintForms.UnitForm.Model;

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
                var inputData = JsonConvert.DeserializeObject<InputPrintFormData>(input["inputData"] as string);
                var printForm = СreatePrintForm(inputData, service);

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

        public object СreatePrintForm(InputPrintFormData inputData, IOrganizationService service)
        {
            var spData = Helper.GetInputDataForSp(service);
            using (var spClient = new SharePointClient(spData.Url, spData.Credentials))
            {
                switch (inputData.PrintFormId)
                {
                    case (int)PrintFormId.Price:
                        var constructor = new ProjectPrintFormConstructor(service, spClient);
                        return constructor.GetForms(inputData);
                    case (int)PrintFormId.MasterPrice:
                        var masterContstrucor = new MasterFormConstructor(service, spClient);
                        return masterContstrucor.GetForms(inputData);
                    case (int)PrintFormId.Unit:
                        return new UnitPrintFrom();
                    default:
                        return default;
                }
            }
        }
    }
}
