﻿using System;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class PAM_RiskRatingAttachment : BasePAMAttachment, IAttachmentType1
    {
        public int OrderNumber { get; set; }
    }
}