using Newtonsoft.Json;
using System.Collections.Generic;

namespace SoftLine.ActionPlugins.PrintForms.ProjectPriceForm.Model
{  
    public class Details
    {
        [JsonProperty("City")]
        public string City { get; set; }

        [JsonProperty("Area")]
        public string Area { get; set; }

        [JsonProperty("Completion")]
        public string Completion { get; set; }

        [JsonProperty("Distance to sea")]
        public string DistanceToSea { get; set; }

        [JsonProperty("Prices from")]
        public string PricesFrom { get; set; }

        [JsonProperty("Completion date")]
        public string CompletionDate { get; set; }
    }

    public class MainTable
    {
        [JsonProperty("blockName")]
        public string BlockName { get; set; }

        [JsonProperty("flats")]
        public List<List<List<string>>> Flats { get; set; }
    }

    public class MainTableColumn
    {
        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("unit")]
        public bool Unit { get; set; }
    }

    public class Metadata
    {
        [JsonProperty("withLogo")]
        public bool WithLogo { get; set; }

        [JsonProperty("mainTableColumns")]
        public List<MainTableColumn> MainTableColumns { get; set; }
    }

    public class ProjectPrintForm
    {
        [JsonProperty("objectName")]
        public string ObjectName { get; set; }

        [JsonProperty("preview")]
        public string Preview { get; set; }

        [JsonProperty("detailsLabel")]
        public string DetailsLabel { get; set; }

        [JsonProperty("details")]
        public Details Details { get; set; }

        [JsonProperty("majorBenefitsLabel")]
        public string MajorBenefitsLabel { get; set; }

        [JsonProperty("majorBenefits")]
        public string MajorBenefits { get; set; }

        [JsonProperty("websiteUrl")]
        public string WebsiteUrl { get; set; }

        [JsonProperty("phoneLabel")]
        public string PhoneLabel { get; set; }

        [JsonProperty("phone")]
        public string Phone { get; set; }

        [JsonProperty("faxLabel")]
        public string FaxLabel { get; set; }

        [JsonProperty("fax")]
        public string Fax { get; set; }

        [JsonProperty("emailLabel")]
        public string EmailLabel { get; set; }

        [JsonProperty("email")]
        public string Email { get; set; }

        [JsonProperty("address")]
        public string Address { get; set; }

        [JsonProperty("tagline")]
        public string Tagline { get; set; }

        [JsonProperty("metadata")]
        public Metadata Metadata { get; set; }

        [JsonProperty("mainTable")]
        public List<MainTable> MainTable { get; set; }
    }


}
