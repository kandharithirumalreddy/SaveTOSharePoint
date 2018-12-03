using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Newtonsoft.Json;

namespace SaveToSharepointWeb.Models
{
    public class DocumentLibraryInfo
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string WebUrl { get; set; }
        public string DriveType { get; set; }
        public string Description { get; set; }
    }
}