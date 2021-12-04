using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Softline.ShortRent
{
    public class CheckShortRentPrice : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(null);

            var target = (Entity)context.InputParameters["Target"];
            var listingid = target.GetAttributeValue<EntityReference>("sl_llstingid");

            if (listingid is null)
            {
                throw new InvalidPluginExecutionException("Listing is Empty");
            }
            var dateFrom = target.GetAttributeValue<DateTime?>("sl_date_from");
            if (!dateFrom.HasValue || dateFrom.Value.Date < DateTime.Now.Date)
            {
                throw new InvalidPluginExecutionException("Date from < Now Date");
            };
            var dateTo = target.GetAttributeValue<DateTime?>("sl_date_to");
            if (!dateTo.HasValue || dateTo <= dateFrom)
            {
                throw new InvalidPluginExecutionException("dateTo <= dateFrom");
            }
            var existingRentPrice = RetriveExistingRentPrice(listingid.Id, dateFrom.Value, dateTo.Value, service);
            if (existingRentPrice is null) return;
            throw new InvalidPluginExecutionException("There is already a price for the dates you enter.");
        }

        private Entity RetriveExistingRentPrice(Guid listingid, DateTime startDate, DateTime endDate, IOrganizationService service)
        {
            var query = $@"<fetch top='1' no-lock='true' >
                           <entity name='sl_listing' >
                             <filter type='and' >
                               <condition attribute='sl_listingid' operator='eq' value='{listingid}' />
                             </filter>
                             <link-entity name='sl_short_rent_price' from='sl_llstingid' to='sl_listingid' link-type='inner' alias='price' >
                               <attribute name='sl_short_rent_priceid' />
                               <filter type='and' >
                                 <condition attribute='sl_date_to' operator='on-or-after' value='{startDate:yyyy-MM-dd}' />
                                 <condition attribute='sl_date_from' operator='on-or-before' value='{endDate:yyyy-MM-dd}' />
                               </filter>
                             </link-entity>
                           </entity>
                         </fetch>";
            return service.RetrieveMultiple(new FetchExpression(query)).Entities.FirstOrDefault();
        }
    }
}
