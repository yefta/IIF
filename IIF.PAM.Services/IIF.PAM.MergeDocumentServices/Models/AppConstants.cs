using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IIF.PAM.MergeDocumentServices.Models
{
    public static class AppConstants
    {
        public static class TableName
        {
            #region PAM
            public static string PAM_ProjectAnalysis { get { return "[dbo].[PAM_ProjectAnalysis]"; } }
            public static string PAM_HistoricalFinancial { get { return "[dbo].[PAM_HistoricalFinancialAndFinancialProjection]"; } }
            public static string PAM_Supplemental { get { return "[dbo].[PAM_SupplementalProcurementAndInsurance]"; } }
            public static string PAM_Social { get { return "[dbo].[PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciples]"; } }
            public static string PAM_GroupStructure { get { return "[dbo].[PAM_GroupStructure]"; } }
            public static string PAM_TermSheet { get { return "[dbo].[PAM_TermSheet]"; } }
            public static string PAM_RiskRating { get { return "[dbo].[PAM_RiskRating]"; } }
            public static string PAM_KYCChecklists { get { return "[dbo].[PAM_KYCChecklists]"; } }
            public static string PAM_OtherBanksFacilities { get { return "[dbo].[PAM_OtherBanksFacilitiesOrSummaryOfPefindoReport]"; } }
            public static string PAM_IndustryAnalysis { get { return "[dbo].[PAM_IndustryAnalysis]"; } }
            public static string PAM_LegalDueDiligenceReport { get { return "[dbo].[PAM_LegalDueDiligenceReport]"; } }
            public static string PAM_SAndEDueDiligence { get { return "[dbo].[PAM_SAndEDueDiligence]"; } }
            public static string PAM_ShareValuationReport { get { return "[dbo].[PAM_ShareValuationReport]"; } }
            public static string PAM_OtherReports { get { return "[dbo].[PAM_OtherReports]"; } }

			public static string PAM_ProjectData { get { return "[dbo].[PAM_ProjectData]"; } }
			public static string PAM_BorrowerOrTargetCompanyData { get { return "[dbo].[PAM_BorrowerOrTargetCompanyData]"; } }
			public static string PAM_ProposalData { get { return "[dbo].[PAM_ProposalData]"; } }
			public static string PAM_RecommendationData { get { return "[dbo].[PAM_RecommendationData]"; } }
			

			#endregion

			#region CM
			public static string CM_PeriodicReview { get { return "[dbo].[CM_PeriodicReview]"; } }
            public static string CM_CreditMemorandum { get { return "[dbo].[CM_CreditMemorandum]"; } }
            public static string CM_PreviousApprovals { get { return "[dbo].[CM_PreviousApprovals]"; } }
            public static string CM_RiskRating{ get { return "[dbo].[CM_RiskRating]"; } }
			//public static string CM_KYCChecklists { get { return "[dbo].[PAM_SupplementalProcurementAndInsurance]"; } }
			public static string CM_KYCChecklists { get { return "[dbo].[CM_KYCChecklists]"; } }
			public static string CM_SAndEReview { get { return "[dbo].[CM_SAndEReview]"; } }
            public static string CM_OtherBanksFacilities { get { return "[dbo].[CM_OtherBanksFacilitiesOrSummaryofPefindoReport]"; } }
            public static string CM_ValuationReport { get { return "[dbo].[CM_ValuationReport]"; } }
            public static string CM_OtherAttachment { get { return "[dbo].[CM_OtherAttachmentOrReports]"; } }
			public static string CM_ProjectData { get { return "[dbo].[CM_ProjectData]"; } }
			public static string CM_BorrowerOrInvesteeCompanyData { get { return "[dbo].[CM_BorrowerOrInvesteeCompanyData]"; } }
			public static string CM_ProposalOrFacilityData { get { return "[dbo].[CM_ProposalOrFacilityData]"; } }
			public static string CM_RecommendationData { get { return "[dbo].[CM_RecommendationData]"; } }
			
			#endregion
		}
    }
}
