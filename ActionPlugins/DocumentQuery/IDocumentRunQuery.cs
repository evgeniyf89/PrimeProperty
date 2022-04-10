using Microsoft.Xrm.Sdk;
using SoftLine.ActionPlugins.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.DocumentQuery
{
   public interface IDocumentRunQuery
    {
        void Init(IOrganizationService crmService);

        WordParameters[] RunQuery(Guid entityId);
    }
}
