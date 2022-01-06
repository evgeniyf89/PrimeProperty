using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using SoftLine.ActionPlugins.OptionSets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins
{
    public class ReserveOrRentProperty : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            var input = context.InputParameters;
            var propertyForOpportunityRef = input["data"] as EntityReference;
            var startDateStr = input["startDate"] as string;
            var endDateStr = input["endDate"] as string;
            var statusStr = input["status"] as string;

            var startDate = DateTime.Parse(startDateStr);
            var endDate = DateTime.Parse(endDateStr);
            var status = (ShortRentSTRentSstatus)int.Parse(statusStr);

            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var systemService = serviceFactory.CreateOrganizationService(null);
            var userService = serviceFactory.CreateOrganizationService(context.UserId);


            var propertyOpportunity = RetrivePropertyOpportunity(propertyForOpportunityRef, systemService);
            var listing = propertyOpportunity.GetAttributeValue<EntityReference>("sl_listingid");
            var opportunityRef = propertyOpportunity.GetAttributeValue<EntityReference>("sl_opportunityid");
            var propertyRef = propertyOpportunity.GetAttributeValue<EntityReference>("sl_propertyid");

            var calendar = new CalendarLogic();
            var rents = calendar.RetriveRent(propertyRef, startDate, endDate, systemService);
            var obj = calendar.Map(rents, startDate, endDate, opportunityRef);
            if (obj?.RentedByOpportunity.Any() ?? false)
            {
                context.OutputParameters["responce"] = JsonConvert.SerializeObject(new { IsError = true, Message = $"Dates are not available for selection." });
                return;
            };
            var rangeDays = (endDate - startDate).Days;
            if (rangeDays < obj.MinDays)
            {
                context.OutputParameters["responce"] = JsonConvert.SerializeObject(new { IsError = true, Message = $"The number of rental days is too small." });
                return;
            };

            if (listing is null)
            {
                context.OutputParameters["responce"] = JsonConvert.SerializeObject(new { IsError = true, Message = $"listing is empty" });
                return;
            }
            var priceLogic = new RentPriceLogic();
            var from = startDate.AddDays(1);
            var to = endDate.AddDays(1);
            var rentPrices = priceLogic.RetriveRentPrice(listing, from, to, systemService);
            var rentPricesObj = priceLogic.Map(from, to, rentPrices);
            var rentalFree = rentPricesObj.Sum(x => x.Price);
            var shortRentAvailable = RetriveShortRentAvailable(propertyOpportunity, systemService);
            string message;
            if (shortRentAvailable is null)
            {
                CreateRentavailable(status, startDate, endDate, propertyOpportunity, userService);
                message = "created";
            }
            else
            {
                UpdateRentavailable(status, startDate, endDate, shortRentAvailable.Id, userService);
                message = "updated";
            }          
            UpdateOpportynity(status, startDate, endDate, opportunityRef, userService);
            UpdatePropertyOpportunity(status, startDate, endDate, rentalFree, propertyForOpportunityRef, userService);
            context.OutputParameters["responce"] = JsonConvert.SerializeObject(new { IsError = false, Message = $"Short rent available {message}" });
        }

        public Entity CreateRentavailable(ShortRentSTRentSstatus status, DateTime startDate, DateTime endDate, Entity propertyOpportunity, IOrganizationService userService)
        {
            var opportunityRef = propertyOpportunity.GetAttributeValue<EntityReference>("sl_opportunityid");
            var propertyRef = propertyOpportunity.GetAttributeValue<EntityReference>("sl_propertyid");
            var rentavailable = new Entity("sl_short_rent_available")
            {
                ["sl_opportunityid"] = opportunityRef,
                ["sl_property"] = propertyRef,
                ["sl_date_from"] = startDate,
                ["sl_date_to"] = endDate,
                ["sl_st_rent_statuscode"] = new OptionSetValue((int)status),
                ["sl_property_for_opportunityid"] = propertyOpportunity.ToEntityReference()
            };
            rentavailable.Id = userService.Create(rentavailable);
            return rentavailable;
        }

        public void UpdatePropertyOpportunity(ShortRentSTRentSstatus status, DateTime startDate, DateTime endDate, decimal rentalFree, EntityReference propertyOpportunityRef, IOrganizationService userService)
        {
            var propertyOpportunity = new Entity(propertyOpportunityRef.LogicalName, propertyOpportunityRef.Id)
            {
                ["sl_date_from"] = startDate,
                ["sl_date_to"] = endDate,
                ["sl_st_rent_statuscode"] = new OptionSetValue((int)status),
                ["sl_deal_amount"] = new Money(rentalFree)
            };
            userService.Update(propertyOpportunity);
        }

        public void UpdateOpportynity(ShortRentSTRentSstatus status, DateTime startDate, DateTime endDate, EntityReference opportunityRef, IOrganizationService userService)
        {
            if (opportunityRef is null) return;
            var opportunity = new Entity(opportunityRef.LogicalName, opportunityRef.Id)
            {
                ["sl_date_from"] = startDate,
                ["sl_date_to"] = endDate,
                ["sl_st_rent_statuscode"] = new OptionSetValue((int)status)
            };
            userService.Update(opportunity);
        }

        public void UpdateRentavailable(ShortRentSTRentSstatus status, DateTime start, DateTime end, Guid rentAvailable, IOrganizationService userService)
        {
            var rentavailable = new Entity("sl_short_rent_available", rentAvailable)
            {
                ["sl_date_from"] = start,
                ["sl_date_to"] = end,
                ["sl_st_rent_statuscode"] = new OptionSetValue((int)status)
            };
            userService.Update(rentavailable);
        }

        public Entity RetrivePropertyOpportunity(EntityReference propertyForOpportunityRef, IOrganizationService systemService)
        {            
            return systemService.Retrieve(propertyForOpportunityRef.LogicalName, propertyForOpportunityRef.Id, new ColumnSet("sl_opportunityid", "sl_listingid", "sl_propertyid"));
        }

        public Entity RetriveShortRentAvailable(Entity propertyForOpportunity, IOrganizationService systemService)
        {
            var opportunityRef = propertyForOpportunity.GetAttributeValue<EntityReference>("sl_opportunityid");
            var propertyRef = propertyForOpportunity.GetAttributeValue<EntityReference>("sl_propertyid");
            if (opportunityRef is null || propertyRef is null) return default;
            var query = $@"<fetch top='1' no-lock='true' >
                          <entity name='sl_short_rent_available' >
                            <attribute name='sl_short_rent_availableid' />
                            <filter type='and' >
                              <condition attribute='sl_opportunityid' operator='eq' value='{opportunityRef.Id}' />
                              <condition attribute='sl_property' operator='eq' value='{propertyRef.Id}' />
                              <condition attribute='sl_property_for_opportunityid' operator='eq' value='{propertyForOpportunity.Id}' />
                            </filter>
                          </entity>
                        </fetch>";
            return systemService.RetrieveMultiple(new FetchExpression(query)).Entities.FirstOrDefault();
        }
    }
}
