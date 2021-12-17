using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using SoftLine.ActionPlugins.Dto;
using SoftLine.ActionPlugins.OptionSets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins
{
    public class CalendarLogic : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            try
            {
                var input = context.InputParameters;
                var projectRef = input["data"] as EntityReference;
                var opportunityRef = input.Contains("opportunityRef") ? input["opportunityRef"] as EntityReference : default;
                var startDateStr = input["startDate"] as string;
                var endDateStr = input["endDate"] as string;

                var startDate = DateTime.Parse(startDateStr);
                var endDate = DateTime.Parse(endDateStr);
                var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                var service = serviceFactory.CreateOrganizationService(null);
                var rents = RetriveRent(projectRef, startDate, endDate, service);
                var rentObj = Map(rents, startDate, endDate, opportunityRef);
                context.OutputParameters["responce"] = JsonConvert.SerializeObject(rentObj);
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

        public DataCollection<Entity> RetriveRent(EntityReference propertyRef, DateTime start, DateTime end, IOrganizationService service)
        {
            var query = $@"<fetch no-lock='true' >
                              <entity name='sl_unit' >
                                <attribute name='sl_min_days_short_rent' />
                                <filter type='and' >
                                  <condition attribute='sl_unitid' operator='eq' value='{propertyRef?.Id}' />
                                </filter>
                                <link-entity name='sl_short_rent_available' from='sl_property' to='sl_unitid' link-type='outer' alias='rent' >
                                  <attribute name='sl_date_from' />
                                  <attribute name='sl_date_to' />
                                  <attribute name='sl_short_rent_availableid' />
                                  <attribute name='sl_opportunityid' />
                                  <attribute name='sl_st_rent_statuscode' />
                                  <filter type='and' >
                                    <condition attribute='sl_date_to' operator='not-null' />
                                    <condition attribute='sl_date_from' operator='not-null' />
                                    <condition attribute='sl_st_rent_statuscode' operator='in' >
                                      <value>588610003</value>
                                      <value>588610002</value>
                                    </condition>
                                  <filter type='or' >
                                      <filter>
                                        <condition attribute='sl_date_to' operator='gt' value='{end:yyyy-MM-dd}' />
                                        <condition attribute='sl_date_from' operator='lt' value='{end:yyyy-MM-dd}' />
                                      </filter>
                                      <filter>
                                        <condition attribute='sl_date_to' operator='gt' value='{start:yyyy-MM-dd}' />
                                        <condition attribute='sl_date_from' operator='lt' value='{start:yyyy-MM-dd}' />
                                      </filter>
                                      <filter>
                                        <condition attribute='sl_date_to' operator='on-or-before' value='{end:yyyy-MM-dd}' />
                                        <condition attribute='sl_date_from' operator='on-or-after' value='{start:yyyy-MM-dd}' />
                                      </filter>
                                    </filter>
                                  </filter>
                                  <order attribute='sl_date_from' />                                
                              </link-entity>
                            </entity>
                          </fetch>";
            return service.RetrieveMultiple(new FetchExpression(query)).Entities;
        }

        public ShortRent Map(IEnumerable<Entity> data, DateTime startDate, DateTime endDate, EntityReference baseOpportunityRef)
        {
            var baseDate = data?.FirstOrDefault();
            if (baseDate is null)
                return new ShortRent();
            var rended = new List<Range>();
            var reserved = new List<Range>();
            var rendedByOtherOpportunity = new List<Range>();
            var reservedByOtherOpportunity = new List<Range>();
            foreach (var entity in data)
            {
                var opportunityRef = entity.GetAttributeValue<AliasedValue>("rent.sl_opportunityid")?.Value as EntityReference;
                var status = entity.GetAttributeValue<AliasedValue>("rent.sl_st_rent_statuscode")?.Value as OptionSetValue;
                var isOppEqual = opportunityRef?.Id == baseOpportunityRef?.Id;
                var from = entity.GetAttributeValue<AliasedValue>("rent.sl_date_from")?.Value as DateTime?;
                var to = entity.GetAttributeValue<AliasedValue>("rent.sl_date_to")?.Value as DateTime?;
                if (!from.HasValue || !to.HasValue) continue;
                var range = new Range(from.Value, to.Value)
                {
                    Id = (Guid)entity.GetAttributeValue<AliasedValue>("rent.sl_short_rent_availableid").Value
                };
                switch (status?.Value)
                {
                    case (int)ShortRentSTRentSstatus.Rented:
                        if (isOppEqual)
                        {
                            rended.Add(range);
                        }
                        else
                        {
                            rendedByOtherOpportunity.Add(range);
                        }
                        break;
                    case (int)ShortRentSTRentSstatus.Reserved:
                        if (isOppEqual)
                        {
                            reserved.Add(range);
                        }
                        else
                        {
                            reservedByOtherOpportunity.Add(range);
                        }
                        break;
                }
            }
            var minDays = baseDate.GetAttributeValue<decimal?>("sl_min_days_short_rent");
            return new ShortRent()
            {
                MinDays = (int?)minDays,
                Rented = rended,
                RentedByOpportunity = rendedByOtherOpportunity,
                Reserved = reserved,
                ReservedByOpportunity = reservedByOtherOpportunity,
                FromDate = startDate,
                ToDate = endDate
            };
        }
    }
}
