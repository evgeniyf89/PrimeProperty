using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using SoftLine.ActionPlugins.Dto;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins
{
    public class RentPriceLogic : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            try
            {
                var input = context.InputParameters;
                var data = input["data"] as EntityReference;
                var startDateStr = input["startDate"] as string;
                var endDateStr = input["endDate"] as string;

                var startDate = DateTime.Parse(startDateStr);
                var endDate = DateTime.Parse(endDateStr);
                if (data is null)
                {
                    context.OutputParameters["responce"] = JsonConvert.SerializeObject(new { IsError = true, Message = $"input data is empty" });
                    return;
                }
                var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                var service = serviceFactory.CreateOrganizationService(null);
                var rentPrices = RetriveRentPrice(data, startDate, endDate, service);
                var ranges = Map(startDate, endDate, rentPrices);
                context.OutputParameters["responce"] = JsonConvert.SerializeObject(ranges);
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

        public List<RentPrice> Map(DateTime startDate, DateTime endDate, IEnumerable<Entity> rentPrices)
        {
            var diff = (endDate - startDate).Days;

            var ranges = new List<RentPrice>();
            var listingPrice = rentPrices.First().GetAttributeValue<Money>("sl_short_rent");
            for (var i = 0; i < diff; i++)
            {
                var date = startDate.AddDays(i);
                var shortRentPrice = rentPrices
                  .FirstOrDefault(x =>
                  {
                      var dateFrom = x.GetAttributeValue<AliasedValue>("price.sl_date_from")?.Value as DateTime?;
                      var dateTo = x.GetAttributeValue<AliasedValue>("price.sl_date_to")?.Value as DateTime?;
                      return date >= dateFrom && date <= dateTo;
                  });

                var price = shortRentPrice?.GetAttributeValue<AliasedValue>("price.sl_price")?.Value as Money ?? listingPrice;
                var rentPrice = new RentPrice(date, price.Value)
                {
                    Owner = (shortRentPrice.GetAttributeValue<AliasedValue>("price.ownerid")?.Value as EntityReference)?.Name,
                    Opportunity = (shortRentPrice.GetAttributeValue<AliasedValue>("price.sl_opportunityid")?.Value as EntityReference)?.Name,
                };
                ranges.Add(rentPrice);
            }
            return ranges;
        }

        public DataCollection<Entity> RetriveRentPrice(EntityReference listingRef, DateTime startDate, DateTime endDate, IOrganizationService service)
        {         
            var query = $@"<fetch no-lock='true' >
                           <entity name='sl_listing' >
                             <attribute name='sl_short_rent' />
                             <filter type='and' >
                               <condition attribute='sl_listingid' operator='eq' value='{listingRef.Id}' />
                             </filter>
                             <link-entity name='sl_short_rent_price' from='sl_llstingid' to='sl_listingid' link-type='outer' alias='price' >
                               <attribute name='sl_price' />
                               <attribute name='sl_date_from' />
                               <attribute name='sl_date_to' />
                               <filter type='and' >
                                 <condition attribute='sl_date_to' operator='on-or-after' value='{startDate:yyyy-MM-dd}' />
                                 <condition attribute='sl_date_from' operator='on-or-before' value='{endDate:yyyy-MM-dd}' />
                               </filter>
                             </link-entity>
                             <link-entity name='sl_unit' from='sl_unitid' to='sl_unitid' link-type='outer' alias='ae' >
                               <link-entity name='sl_short_rent_available' from='sl_property' to='sl_unitid' link-type='outer' alias='shortRent' >
                                 <attribute name='sl_opportunityid' />
                                 <attribute name='ownerid' />
                                 <filter type='and' >
                                   <condition attribute='sl_date_to' operator='on-or-after' value='{startDate:yyyy-MM-dd}' />
                                   <condition attribute='sl_date_from' operator='on-or-before' value='{endDate:yyyy-MM-dd}' />
                                 </filter>
                               </link-entity>
                             </link-entity>
                           </entity>
                         </fetch>"
            return service.RetrieveMultiple(new FetchExpression(query)).Entities;
        }
    }
}
