﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using SoftLine.ActionPlugins.PrintForms;
using SoftLine.ActionPlugins.PrintForms.MasterPrice;
using SoftLine.ActionPlugins.PrintForms.Metadata;
using SoftLine.ActionPlugins.PrintForms.UnitForm;
using SoftLine.ActionPlugins.ProjectPriceForm.PrintForms;
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
            var str = @"AuthType=Office365;Url=https://ppcrm1.crm4.dynamics.com;Username=subbotin.a@prime-property.com;Password=Port-0712";
            _service = new CrmServiceClient(str);


            //_printedForm = new CreatePrintedForm();
            //var crmSoap = "https://ppcrm1.api.crm4.dynamics.com/XRMServices/2011/Organization.svc";
            //var uri = new Uri(crmSoap);
            //var credentials = new ClientCredentials();
            //credentials.UserName.Password = "Port-0712";
            //credentials.UserName.UserName = "subbotin.a@prime-property.com";
            //ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            //_service = new OrganizationServiceProxy(uri, null, credentials, null);
        }

        [TestMethod()]
        public void RetrivePrintedFormTest()
        {

            var str = "<li>距著名的达索迪桉树公园的沙滩700米；</li><li>利马索尔最受欢迎的地区之一；</li>" +
                "<li>室外游泳池和带顶停车场；</li><li>顶级标准装修：拼花地板、高天花板（3.2米）、大理石卫生间、安全大门、" +
                "热铝窗框、顶级标准嵌入式家具和卫生洁具</li><li>标准的水地板采暖、VRV空调；</li><li>顶层套房享受私人屋顶露台和游泳池；</li></ul>";
            var ss = FindFourParagraph(str);



            var spData = Helper.GetInputDataForSp(_service);
            var spService = new SharePointClient(spData.Url, spData.Credentials);
            var printFormConstructor = new ProjectPrintFormConstructor(_service, spService);
            var inputData = new InputPrintFormData()
            {
                IsWithLogo = true,
                Language = new EntityReference("sl_language", new Guid("{a4523a24-20db-eb11-bacb-000d3a2c3636}")),
                Market = new OptionSetValue(486160000),
                PromotionType = new OptionSetValue(102690000),
                TargetEntityIds = new[] {
                    "b10d68db-b773-ec11-8941-002248818536", "af0d68db-b773-ec11-8941-002248818536",
                    "7f0d68db-b773-ec11-8941-002248818536", "530d68db-b773-ec11-8941-002248818536",
                    "e808fa32-cb0f-ec11-b6e6-0022488425d2" }.Select(x => new Guid(x)).ToArray()
            };
            var tt = printFormConstructor.GetForms(inputData);
            var json = JsonConvert.SerializeObject(tt);
        }

        private string FindFourParagraph(string text)
        {
            if (string.IsNullOrEmpty(text)) return default;
            var i = 0;
            var paragraphIndex = 0;
            var findWord = "</li>";
            while (i < 4)
            {
                paragraphIndex = text.IndexOf(findWord, paragraphIndex);
                if (paragraphIndex == -1)
                    return text;
                i++;
                paragraphIndex += findWord.Length;
            }
            return text.Substring(0, paragraphIndex);
        }

        [TestMethod()]
        public void PGetMasterPintFormTest()
        {

            var inputData = new InputPrintFormData()
            {
                IsWithLogo = true,
                Language = new EntityReference("sl_language", new Guid("a0523a24-20db-eb11-bacb-000d3a2c3636")),
                Market = new OptionSetValue(486160002),
                PromotionType = new OptionSetValue(588610000),
                TargetEntityIds = new[] {
                    "db0468db-b773-ec11-8941-002248818536", "f80468db-b773-ec11-8941-002248818536",
                    "770568db-b773-ec11-8941-002248818536"}.Select(x => new Guid(x)).ToArray(),
                PrintFormId = (int)PrintFormId.Price
            };

            var ss = JsonConvert.SerializeObject(inputData);

            var js = "{'TargetEntityIds':['{5C993064-8974-EC11-8941-002248818089}'],'PrintFormId':'3','Language':{'LogicalName':'sl_language','Id':'a0523a24-20db-eb11-bacb-000d3a2c3636'},'Market':{'Value':'486160000'},'PromotionType':{'Value':'102690000'},'IsWithLogo':false,'Unit':'1','IsWithPrice':false,'IsWithId':false,'IsOnePage':false}";
            var dd = JsonConvert.DeserializeObject<InputPrintFormData>(js);

            var spData = Helper.GetInputDataForSp(_service);
            var spService = new SharePointClient(spData.Url, spData.Credentials);
            var printForm = new ProjectPrintFormConstructor(_service, spService);
            var tt = printForm.GetForms(inputData);
            var json = JsonConvert.SerializeObject(tt);
        }

        [TestMethod()]
        public void GetUnitPintFormTest()
        {
            var inputDataJson = "{'TargetEntityIds':['{cfe22b82-c51a-ec11-b6e6-6045bd89207e}'],'PrintFormId':'3'," +
                "'Language':{'LogicalName':'sl_language','Id':'A2523A24-20DB-EB11-BACB-000D3A2C3636'}," +
                "'Market':{'Value':'486160000'},'PromotionType':{'Value':'102690000'}," +
                "'IsWithLogo':true,'Unit':'1','IsWithPrice':true,'IsWithId':true,'IsOnePage':true}";
            var inputData = JsonConvert.DeserializeObject<InputPrintFormData>(inputDataJson);

            var spData = Helper.GetInputDataForSp(_service);
            var spService = new SharePointClient(spData.Url, spData.Credentials);
            var printForm = new UnitPrintFormConstructor(_service, spService);
            var tt = printForm.GetForms(inputData);
            var json = JsonConvert.SerializeObject(tt);
        }

        /*
         
         */
    }
}