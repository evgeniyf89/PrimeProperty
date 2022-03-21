using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using SoftLine.ActionPlugins;
using System;
using System.Collections.Generic;
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

            _printedForm.RetrivePrintedForm();
        }

        [TestMethod()]
        public void RenderHtmlToPdfTest()
        {
            _printedForm.RenderHtmlToPdf();
        }

        [TestMethod()]
        public void GetPrintFormSettingsTest()
        {
            var settings = _printedForm.GetPrintFormSettings(_service);
            var isValid = settings.ValidData(out string message);
        }

        [TestMethod()]
        public void GetFileByAbsoluteUrlTest()
        {
            var settings = _printedForm.GetPrintFormSettings(_service);
            var url = $"{settings.TemplateFolderInSharePoint}/{settings.TemplateNameInSharePoint}";
            var bytes = _printedForm.GetFileByAbsoluteUrl(url, _service);
        }
    }
}