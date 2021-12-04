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

namespace SoftLine.ActionPlugins
{
    class Helper
    {
        public static ICredentials GetСredentialsForSp(IOrganizationService service)
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
    }
}
