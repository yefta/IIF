using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IIF.PAM.MergeDocumentServices.Models
{
	public class CMData
	{
		public Int64 Id { get; set; }
		public string ReviewMemo { get; set; }
		public string ProductType { get; set; }
		public string CompanyName { get; set; }
		public string CMNumber { get; set; }
		public string ProjectName { get; set; }
		public string ProjectCode { get; set; }
		public DateTime CMDate { get; set; }
		public string ProjectDescription { get; set; }
		public string SubSector { get; set; }
		public string SubSectorDesc { get; set; }
		public string Sector { get; set; }
		public string SectorDesc { get; set; }
		public string ProjectCosCUrr { get; set; }
		public string ProjectCostAmount { get; set; }
		public string ProjectScope { get; set; }
		public string ProjectStructure { get; set; }
		public string FundingNeeds { get; set; }
		public string DealStrategy { get; set; }
		public string UltimateBeneficialOwner { get; set; }
		public string IIFRate { get; set; }
		public DateTime? IIFRatingDate { get; set; }
		public string SAndPRate { get; set; }
		public string MoodysRate { get; set; }
		public string FitchRate { get; set; }
		public string PefindoRate { get; set; }
		public string SAndECategoryRate { get; set; }
		public string SAndECategoryType { get; set; }
		public string LQCOrBICheckingRate { get; set; }
		public string BusinessActivities { get; set; }
		public string OtherInformation { get; set; }
		public string Purpose { get; set; }
		public string ApprovalAuhority { get; set; }
		public string GroupExposureCurr { get; set; }
		public string GroupExposureAmount { get; set; }
		public string Remarks { get; set; }
		public string FacilityOrInvestmentRemarks { get; set; }
		public int? ExpectedHoldingPeriodYear { get; set; }
		public int? ExpectedHoldingPeriodMonth { get; set; }
		public int TenorMonth { get; set; }
		public int TenorYear { get; set; }
		public int AverageLoanLifeMonth { get; set; }
		public int AverageLoanLifeYear { get; set; }
		public string PricingInterestRate { get; set; }
		public string PricingCommitmentFee { get; set; }
		public string PricingUpfrontFacilityFee { get; set; }
		public string PricingStructuringFee { get; set; }
		public string PricingArrangerFee { get; set; }
		public string PricingCollateral { get; set; }
		public string PricingOtherConditions { get; set; }
		public string PricingExceptionToIIFPolicy { get; set; }
		public string ProposalReviewPeriod { get; set; }
		public Int64 FacilityLimitComplianceCurrencyId { get; set; }
		public string LimitComplianceCurrency { get; set; }
		public int FacilityLimitComplianceMonth { get; set; }
		public int FacilityLimitComplianceYear { get; set; }
		public string FacilityLimitComplianceSingleProjectExposureMaxLimit { get; set; }
		public string FacilityLimitComplianceSingleProjectExposureProposed { get; set; }
		public string SingleProjectExposureRemarks { get; set; }
		//public string FacilityLimitComplianceProductItemId { get; set; }
		public string FacilityLimitComplianceProductMaxLimit { get; set; }
		public string FacilityLimitComplianceProductProposed { get; set; }
		public string ProductRemarks { get; set; }
		public int FacilityLimitComplianceRiskRatingId { get; set; }
		public string FacilityLimitComplianceIIFRate { get; set; }
		public string FacilityLimitComplianceRiskRatingMaxLimit { get; set; }
		public string FacilityLimitComplianceRiskRatingProposed { get; set; }
		public string RiskRatingRemarks { get; set; }
		public string FacilityLimitComplianceGrupExposureMaxLimit { get; set; }
		public string FacilityLimitComplianceGrupExposureProposed { get; set; }
		public string GrupExposureRemarks { get; set; }		
		public string FacilityLimitComplianceSectorExposureMaxLimit { get; set; }
		public string FacilityLimitComplianceSectorExposureProposed { get; set; }
		public string SectorExposureRemarks { get; set; }
		public string notes { get; set; }
		public string KeyInvestmentRecommendation { get; set; }
		public string Recommendation { get; set; }
		public string AccountResponsibleCIOName { get; set; }

	}
}
