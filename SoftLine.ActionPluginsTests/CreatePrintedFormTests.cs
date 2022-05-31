using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Newtonsoft.Json;
using SoftLine.ActionPlugins.PrintForms;
using SoftLine.ActionPlugins.PrintForms.Metadata;
using SoftLine.ActionPlugins.SharePoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.Tests
{
    [TestClass()]
    public class CreatePrintedFormTests
    {
        private readonly CreatePrintedForm _printedForm;
        private readonly IOrganizationService _service;
        public CreatePrintedFormTests()
        {
            _printedForm = new CreatePrintedForm();
            var crmSoap = "https://ppcrm1.api.crm4.dynamics.com/XRMServices/2011/Organization.svc";
            var uri = new Uri(crmSoap);
            var credentials = new ClientCredentials();
            credentials.UserName.Password = "SuperMuper!!";
            credentials.UserName.UserName = "test.crm@prime-property.com";
            ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            _service = new OrganizationServiceProxy(uri, null, credentials, null);
        }

        [TestMethod()]
        public void RetrivePrintedFormTest()
        {
            var spData = Helper.GetInputDataForSp(_service);
            var spService = new SharePointClient(spData.Url, spData.Credentials);
            var printFormConstructor = new PrintFormConstructor(_service, spService);
            var inputData = new InputPrintFormData()
            {
                IsWithLogo = true,
                Language = new EntityReference("sl_language", new Guid("{A0523A24-20DB-EB11-BACB-000D3A2C3636}")),
                Market = new OptionSetValue(486160002),
                PromotionType = new OptionSetValue(102690001),
                TargetEntityRef = new EntityReference("sl_project", new Guid("30824c04-2cdb-eb11-bacb-000d3a2c32dc"))
            };
            var tt = printFormConstructor.GetForm(inputData);
            var json = JsonConvert.SerializeObject(tt);
        }

        [TestMethod()]
        public void GetPrintFormSettingsTest()
        {

        }
        
    }
}