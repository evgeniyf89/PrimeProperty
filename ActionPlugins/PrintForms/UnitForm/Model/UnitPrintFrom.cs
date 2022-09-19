using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.PrintForms.UnitForm.Model
{
    public class Characteristic
    {
        [JsonProperty("label")]
        public string Label { get; set; }

        [JsonProperty("value")]
        public string Value { get; set; }
    }

    public class Description
    {
        [JsonProperty("label")]
        public string Label { get; set; }

        [JsonProperty("value")]
        public string Value { get; set; }
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

        [JsonProperty("objectLabel")]
        public string ObjectLabel { get; set; }

        [JsonProperty("areaLabel")]
        public string AreaLabel { get; set; }

        [JsonProperty("priceLabel")]
        public string PriceLabel { get; set; }

        [JsonProperty("exclusiveLabel")]
        public string ExclusiveLabel { get; set; }

        [JsonProperty("descriptionLabel")]
        public string DescriptionLabel { get; set; }

        [JsonProperty("characteristicsLabel")]
        public string CharacteristicsLabel { get; set; }

        [JsonProperty("vatLabel")]
        public string VatLabel { get; set; }

        [JsonProperty("reportName")]
        public string ReportName { get; set; }

        [JsonProperty("withLogo")]
        public bool WithLogo { get; set; }
    }

    public class Object
    {
        [JsonProperty("objectId")]
        public ObjectId ObjectId { get; set; }

        [JsonProperty("price")]
        public string Price { get; set; }

        [JsonProperty("vatStatus")]
        public bool VatStatus { get; set; }

        [JsonProperty("area")]
        public string Area { get; set; }

        [JsonProperty("isExclusive")]
        public bool IsExclusive { get; set; }

        [JsonProperty("previewUrl")]
        public string PreviewUrl { get; set; }

        [JsonProperty("map")]
        public string Map { get; set; }

        [JsonProperty("characteristics")]
        public List<Characteristic> Characteristics { get; set; }

        [JsonProperty("description")]
        public List<Description> Description { get; set; }

        [JsonProperty("pictureGallery")]
        public List<string> PictureGallery { get; set; }
    }

    public class ObjectId
    {
        [JsonProperty("label")]
        public string Label { get; set; }

        [JsonProperty("value")]
        public string Value { get; set; }
    }

    public class UnitPrintFrom
    {
        [JsonProperty("metadata")]
        public Metadata Metadata { get; set; }

        [JsonProperty("objects")]
        public List<Object> Objects { get; set; }
    }


}
