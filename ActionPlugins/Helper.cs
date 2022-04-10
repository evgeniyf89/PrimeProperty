using Microsoft.SharePoint.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SoftLine.ActionPlugins.SharePoint.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins
{
    class Helper
    {
        public static SharePointData GetInputDataForSp(IOrganizationService service)
        {
            var query = $@"<fetch no-lock='true' >
                           <entity name='sl_setting' >                             
                             <attribute name='sl_key' />
                             <attribute name='sl_value' />                           
                             <filter type='and' >
                               <condition attribute='sl_key' operator='in' >
                                 <value>SpPassword</value>
                                 <value>SpUser</value>
                                 <value>SpURL</value>
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

            var url = settings
                .First(x => x.GetAttributeValue<string>("sl_key") == "SpURL")
                .GetAttributeValue<string>("sl_value");

            var passWord = new SecureString();
            Array.ForEach(passwordString.ToCharArray(), ch => passWord.AppendChar(ch));
            var credentials = new SharePointOnlineCredentials(userName, passWord);
            return new SharePointData(credentials, url);
        }
    }
}
