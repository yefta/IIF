using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace IIF.PAM.MergeDocumentServices.Models
{
    public class PAMData
    {
        public Int64 Id { get; set; }
        public string ProductType { get; set; }
        public string ProjectCompanyName { get; set; }
        public string ProjectName { get; set; }
        public string ProjectCode { get; set; }
        public DateTime PAMDate { get; set; }
        public string ProjectCostCurr { get; set; }
        public decimal ProjectCostAmount { get; set; }
        public string ProjectDescription { get; set; }
		public string SubSector { get; set; }
		public string SubSectorDesc { get; set; }
		public string Sector { get; set; }
		public string SectorDesc { get; set; }
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
        public string LQCOrBICheckingRate { get; set; }
        public string BusinessActivities { get; set; }
        public string OtherInformation { get; set; }
        public string Purpose { get; set; }
        public string ApprovalAuthority { get; set; }
        public string GroupExposureCurr { get; set; }
        public decimal GroupExposureAmount { get; set; }
        public string Remarks { get; set; }
        public int? tenorMonth { get; set; }
        public int? tenorYear { get; set; }
        public int? averageLoanLifeMonth { get; set; }
        public int? averageLoanLifeYear { get; set; }
        public string pricingInterestRate { get; set; }
        public string pricingCommitmentFee { get; set; }
        public string pricingUpfrontFacilityFee { get; set; }
        public string pricingStructuringFee { get; set; }
        public string pricingArrangerFee { get; set; }
        public string pricingCollateral { get; set; }
        public string pricingOtherConditions { get; set; }
        public string pricingExceptionToIIFPolicy { get; set; }
        public string reviewPeriod { get; set; }
        public string KeyInvestmentRecommendation { get; set; }
        public string Recommendation { get; set; }
        public string AccountResponsibleCIOName { get; set; }
    }
}
