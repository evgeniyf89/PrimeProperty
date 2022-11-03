using Newtonsoft.Json;
using SoftLine.ActionPlugins.PrintForms.ProjectPriceForm.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.PrintForms.MasterPrice.Model
{
    
    public class Block
    {
        [JsonProperty("blockName")]
        public string BlockName { get; set; }

        [JsonProperty("flatGroups")]
        public List<FlatGroup> FlatGroups { get; set; }
    }
   

    public class Flat
    {
        [JsonProperty("flatNumber")]
        public string FlatNumber { get; set; }

        [JsonProperty("ibp")]
        public string Ibp { get; set; }

        [JsonProperty("floor")]
        public string Floor { get; set; }

        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("bedrooms")]
        public string Bedrooms { get; set; }

        [JsonProperty("indoorArea")]
        public string IndoorArea { get; set; }

        [JsonProperty("covereddVeranda")]
        public string CovereddVeranda { get; set; }

        [JsonProperty("uncoveredVeranda")]
        public string UncoveredVeranda { get; set; }

        [JsonProperty("roofTerrace")]
        public string RoofTerrace { get; set; }

        [JsonProperty("commonArea")]
        public string CommonArea { get; set; }

        [JsonProperty("totalArea")]
        public string TotalArea { get; set; }

        [JsonProperty("pool")]
        public string Pool { get; set; }

        [JsonProperty("price")]
        public string Price { get; set; }

        [JsonProperty("vat")]
        public string Vat { get; set; }
    }

    public class FlatGroup
    {
        [JsonProperty("flats")]
        public List<Flat> Flats { get; set; }
    }

    public class Metadata
    {
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

        [JsonProperty("withLogo")]
        public bool WithLogo { get; set; }

        [JsonProperty("notification")]
        public string Notification { get; set; }

        [JsonProperty("mainTableColumns")]
        public List<MainTableColumn> MainTableColumns { get; set; }
    }

    public class Object
    {
        [JsonProperty("objectName")]
        public string ObjectName { get; set; }

        [JsonProperty("preview")]
        public string Preview { get; set; }

        [JsonProperty("detailsLabel")]
        public string DetailsLabel { get; set; }

        [JsonProperty("details")]
        public List<Detail> Details { get; set; }

        [JsonProperty("majorBenefitsLabel")]
        public string MajorBenefitsLabel { get; set; }

        [JsonProperty("majorBenefits")]
        public string MajorBenefits { get; set; }

        [JsonProperty("blocks")]
        public List<Block> Blocks { get; set; }
    }

    public class ObjectGroup
    {
        [JsonProperty("groupName")]
        public string GroupName { get; set; }

        [JsonProperty("objects")]
        public List<Object> Objects { get; set; }
    }

    public class Root
    {
        [JsonProperty("metadata")]
        public Metadata Metadata { get; set; }

        [JsonProperty("objectGroups")]
        public List<ObjectGroup> ObjectGroups { get; set; }
    }


}
