﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaveToSharepointWeb.Models
{
    public class SaveAttachmentRequest
    {
        public string[] attachmentIds { get; set; }
        public string messageId { get; set; }
        public string driveId { get; set; }
        public string folderName { get; set; }
    }
}