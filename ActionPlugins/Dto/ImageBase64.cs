using System;
using System.Collections.Generic;

namespace SoftLine.ActionPlugins.Dto
{
    public class ImageBase64
    {
        public string FullName { get; set; }

        public string NameForUrl { get; set; }

        public string Base64 { get; set; }

        public string MimeType { get; set; }

        public string FolderName { get; set; }

        public int Height { get; set; }
        public int Width { get; set; }

        public decimal? Weight { get; set; }

        public Guid? Formatid { get; set; }

        public DateTime? LastModifiedDate { get; set; }
    }


    public class Images
    {
        public string BaseFolder { get; set; }

        public IEnumerable<ImageBase64> Base;

        public IEnumerable<ImageBase64> Resize { get; set; }
    }
}
