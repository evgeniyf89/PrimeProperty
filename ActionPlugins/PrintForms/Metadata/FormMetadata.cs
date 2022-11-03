using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.PrintForms.Metadata
{
    public enum FormMetadata
    {
        DetailsLabel = 01001,
        ReportName = 01003,
        CityName = 01004,
        RegionName = 01005,
        StageName = 01006,
        ToSeaName = 01007,
        PriceFromName = 01008,
        CompletedConstructionName = 01009,
        MainBenefitsName = 01010,
        FloorTitle = 01014,
        TypeHeader = 01015,
        BedroomsTitle = 01016,
        InnerTitle = 01017,
        CoveredVerandaTitle = 01018,
        OpenVerandaTitle = 01019,
        TerasseTitle = 01020,
        GeneralUsageTitle = 01021,
        TotalAreaHeader = 01022,
        PoolTitle = 01023,
        PriceHeader = 01024,
        VATHeader = 01025,
        PlotTitle = 01026,
        StorageTitle = 01027,
        ParkinTitle = 01028,
        M2Title = 01029,
        IBP = 01030,        
        Sold = 01035,
        Rented =01036,
        PhoneLabel = 01037,
        FaxLabel = 01038,
        EmailLabel = 01039,
        Tagline = 01040,
        FormName = 01059,
        SellerNotification = 01060,
        Id = 01080,
        CharacteristicsLabel = 01127,
        DescriptionLabel = 01057,
        ProjectName = 01068,
        PidName = 01081,
        OwnerName = 01082,
        UidName = 01083,
        TypeName = 01084,
        CommisionName = 01085,
        OwnerEmail = 01086,
        NotesName = 01087,
        GoogleName = 01088,
        PriceLabelPromotion = 01111,
        PriceLabelLtRent = 01100,
        PriceLabelStRent = 01101
    }
    public static class ProjectPriceMetadataExtensions
    {
        public static string StringValue(this FormMetadata projectPriceMetadata)
        {
            return $"0{(int)projectPriceMetadata}";
        }
    }

}
