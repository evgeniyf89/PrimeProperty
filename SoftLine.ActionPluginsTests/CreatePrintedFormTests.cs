﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
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

        }

        [TestMethod()]
        public void GetPrintFormSettingsTest()
        {

        }

        [TestMethod()]
        public void GetFileByAbsoluteUrlTest()
        {
            var settings = _printedForm.FormPrintForm(1, new Guid("8c0768db-b773-ec11-8941-002248818536"), Guid.NewGuid(), _service);           
            var savePath2 = @"E:Project price bbf c параметрами.docx";
            File.WriteAllBytes(savePath2, settings);
        }
    }
}