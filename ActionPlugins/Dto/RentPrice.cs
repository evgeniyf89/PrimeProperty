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

        [JsonProperty("employee")]
        public string Employee { get; set; }

        [JsonProperty("existing_reservation")]
        public string ExistingReservation { get; set; }

        public RentPrice(DateTime date, decimal price)
        {
            Date = date;
            Price = price;
        }
    }
}
