﻿using System;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class PAM_OtherSupportingDocumentAttachment : BasePAMAttachment, IAttachmentType3
    {
        public string Description { get; set; }
        public int MWorkflowStatusIdWhenAdded { get; set; }
        public string SNWhenAdded { get; set; }
    }
}