using Newtonsoft.Json;
using System;

namespace SoftLine.ActionPlugins.Dto
{
    public class RentPrice
    {
        [JsonProperty("date")]
        public DateTime Date { get; set; }

        [JsonProperty("price")]
        public decimal Price { get; set; }

        [JsonProperty("owner")]
        public string Owner { get; set; }

        [JsonProperty("opportunity")]
        public string Opportunity { get; set; }

        public RentPrice(DateTime date, decimal price)
        {
            Date = date;
            Price = price;
        }
    }
}
