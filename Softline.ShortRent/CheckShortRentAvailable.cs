using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Softline.ShortRent
{
    public class CheckShortRentAvailable : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(null);

            var target = (Entity)context.InputParameters["Target"];
            var propertyRef = target.GetAttributeValue<EntityReference>("sl_property");

            if (propertyRef is null)
            {
                throw new InvalidPluginExecutionException("Property is empty");
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
            var existingRentAvailable = RetriveExistingRentAvailable(propertyRef.Id, dateFrom.Value, dateTo.Value, service);
            if (existingRentAvailable is null) return;
            throw new InvalidPluginExecutionException("There is already a reservation or rental for the dates you enter.");
        }

        private object RetriveExistingRentAvailable(Guid propertyid, DateTime startDate, DateTime endDate, IOrganizationService service)
        {
            var query = $@"<fetch top='1' no-lock='true' >
                          <entity name='sl_unit' >
                            <filter type='and' >
                              <condition attribute='sl_unitid' operator='eq' value='{propertyid}' />
                            </filter>
                            <link-entity name='sl_short_rent_available' from='sl_property' to='sl_unitid' link-type='inner' alias='ab' >
                              <attribute name='sl_short_rent_availableid' />                           
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
