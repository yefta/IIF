using System;

namespace IIF.PAM.WebServices.Models
{
    public enum AttachmentTypeConstants
    {
        PAM_MergedDocument = 1,
        PAM_LegalDueDiligence = 2,
        PAM_SAndEDueDiligence = 3,
        PAM_OtherReports = 4,
        PAM_OtherSupportingDocument = 5,
        PAM_DocumentRelatedToBoDDecision = 6,
        PAM_DocumentRelatedToBoCDecision = 7,
        CM_MergedDocument = 8,
        CM_OtherSupportingDocument = 9,
        CM_DocumentRelatedToBoDDecision = 10,
        CM_DocumentRelatedToBoCDecision = 11,
    }

    public static class AttachmentTypeConstants_Extensions
    {
        public static string ToDisplayName(this AttachmentTypeConstants value)
        {
            string result = string.Empty;

            switch (value)
            {
                case AttachmentTypeConstants.PAM_MergedDocument:
                    result = "Project Proposal";
                    break;
                case AttachmentTypeConstants.PAM_LegalDueDiligence:
                    result = "Project Proposal";
                    break;
                case AttachmentTypeConstants.PAM_SAndEDueDiligence:
                    result = "Project Proposal";
                    break;
                case AttachmentTypeConstants.PAM_OtherReports:
                    result = "Other Reports";
                    break;
                case AttachmentTypeConstants.PAM_OtherSupportingDocument:
                    result = "Project Proposal";
                    break;
                case AttachmentTypeConstants.PAM_DocumentRelatedToBoDDecision:
                    result = "MoM";
                    break;
                case AttachmentTypeConstants.PAM_DocumentRelatedToBoCDecision:
                    result = "MoM";
                    break;
                case AttachmentTypeConstants.CM_MergedDocument:
                    result = "Credit Memo Memorandum";
                    break;
                case AttachmentTypeConstants.CM_OtherSupportingDocument:
                    result = "Project Proposal";
                    break;
                case AttachmentTypeConstants.CM_DocumentRelatedToBoDDecision:
                    result = "MoM";
                    break;
                case AttachmentTypeConstants.CM_DocumentRelatedToBoCDecision:
                    result = "MoM";
                    break;
                default:
                    throw new NotImplementedException(value.ToString() + " is not implemented in ToDisplayName method.");
            }
            return result;
        }

        public static string ToDMSMetadataDisplayName(this AttachmentTypeConstants value)
        {
            string result = string.Empty;

            switch (value)
            {
                case AttachmentTypeConstants.PAM_MergedDocument:
                    result = "Project Proposal";
                    break;
                case AttachmentTypeConstants.PAM_LegalDueDiligence:
                    result = "Project Proposal";
                    break;
                case AttachmentTypeConstants.PAM_SAndEDueDiligence:
                    result = "Project Proposal";
                    break;
                case AttachmentTypeConstants.PAM_OtherReports:
                    result = "Other Reports";
                    break;
                case AttachmentTypeConstants.PAM_OtherSupportingDocument:
                    result = "Project Proposal";
                    break;
                case AttachmentTypeConstants.PAM_DocumentRelatedToBoDDecision:
                    result = "MoM";
                    break;
                case AttachmentTypeConstants.PAM_DocumentRelatedToBoCDecision:
                    result = "MoM";
                    break;
                case AttachmentTypeConstants.CM_MergedDocument:
                    result = "Credit Memo/Memorandum";
                    break;
                case AttachmentTypeConstants.CM_OtherSupportingDocument:
                    result = "Project Proposal";
                    break;
                case AttachmentTypeConstants.CM_DocumentRelatedToBoDDecision:
                    result = "MoM";
                    break;
                case AttachmentTypeConstants.CM_DocumentRelatedToBoCDecision:
                    result = "MoM";
                    break;
                default:
                    throw new NotImplementedException(value.ToString() + " is not implemented in ToDMSMetadataDisplayName method.");
            }
            return result;
        }
    }
}