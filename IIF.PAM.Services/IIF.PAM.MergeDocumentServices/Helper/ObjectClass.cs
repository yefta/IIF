using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IIF.PAM.MergeDocumentServices.Helper
{
    public class ObjectClass
    {
        public class borrowerList
        {
            public string BorrowerOrTargetCompany { get; set; }
            public string ProjectSponsorOrShareHolder { get; set; }
            public decimal PercentOwnership { get; set; }
        }

        public class FacilityList
        {
            public string Type { get; set; }
            public string CurrFacility { get; set; }
            public decimal Amount { get; set; }
        }

        public class CM_Data
        {
            public Int64 Id { get; set; }
            public string ProductType { get; set; }
            public string ReviewMemo  { get; set; }
            public string CompanyName { get; set; }
            public string ProjectName { get; set; }
            public string ProjectCode { get; set; }
            public DateTime? CMDate { get; set; }
            public string ProjectDescription { get; set; }
            public decimal SubSector { get; set; }
            public string Sector { get; set; }
            public string ProjectCosCUrr { get; set; }
            public string ProjectCostAmount { get; set; }
            public string ProjectScope { get; set; }
            public string ProjectStructure { get; set; }
            public string DealStrategy { get; set; }
            public string UltimateBeneficialOwner { get; set; }
            public string OtherInformation { get; set; }
            public DateTime? Purpose { get; set; }
            public string ApprovalAuhority { get; set; }
            public string GroupExposureCurr { get; set; }
            public string GroupExposureAmount { get; set; }
            public string Remarks { get; set; }
            public string TenorMonth { get; set; }
            public string TenorYear { get; set; }
            public string AverageLoanLifeMonth { get; set; }
            public string AverageLoanLifeYear { get; set; }
            public string PricingInterestRate { get; set; }
            public string PricingCommitmentFee { get; set; }
            public decimal PricingUpfrontFacilityFee { get; set; }
            public string PricingStructuringFee { get; set; }
            public int? PricingArrangerFee { get; set; }
            public int? PricingCollateral { get; set; }
            public int? PricingOtherConditions { get; set; }
            public int? PricingExceptionToIIFPolicy { get; set; }
            public string ProposalReviewPeriod { get; set; }
            public string KeyInvestmentRecommendation { get; set; }
            public string Recommendation { get; set; }
        }
    }
}
