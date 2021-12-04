using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace SoftLine.ActionPlugins.Dto
{
    public class ShortRent
    {
        [JsonProperty("minDays")]
        public int? MinDays { get; set; }

        [JsonProperty("fromDate")]
        public DateTime? FromDate { get; set; }

        [JsonProperty("toDate")]
        public DateTime? ToDate { get; set; }

        [JsonProperty("rented")]
        public IEnumerable<Range> Rented;

        [JsonProperty("reserved")]
        public IEnumerable<Range> Reserved;

        [JsonProperty("rentedByOpportunity")]
        public IEnumerable<Range> RentedByOpportunity;

        [JsonProperty("reservedByOpportunity")]
        public IEnumerable<Range> ReservedByOpportunity;

    }

    public class Range
    {
        [JsonProperty("dateFrom")]
        public DateTime DateFrom { get; set; }

        [JsonProperty("dateTo")]
        public DateTime DateTimeTo { get; set; }

        [JsonProperty("id")]
        public Guid Id { get; set; }

        public Range(DateTime from, DateTime to)
        {
            DateTimeTo = to;
            DateFrom = from;
        }
       
    }
}
