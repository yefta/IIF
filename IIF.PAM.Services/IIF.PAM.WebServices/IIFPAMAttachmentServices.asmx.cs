using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Web.Services;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;
using IIF.PAM.WebServices.Services;

namespace IIF.PAM.WebServices
{
    /// <summary>
    /// Summary description for AttachmentServices
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class IIFPAMAttachmentServices : BaseWebService
    {
        [WebMethod]
        public List<PAM_ProjectAnalysisAttachment> List_PAM_ProjectAnalysisAttachment(long pamId)
        {
            PAM_ProjectAnalysis_Services svc = new PAM_ProjectAnalysis_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_ProjectAnalysisAttachment_Content(long id)
        {
            PAM_ProjectAnalysis_Services svc = new PAM_ProjectAnalysis_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_HistoricalFinancialAndFinancialProjectionAttachment> List_PAM_HistoricalFinancialAndFinancialProjectionAttachment(long pamId)
        {
            PAM_HistoricalFinancialAndFinancialProjection_Services svc = new PAM_HistoricalFinancialAndFinancialProjection_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_HistoricalFinancialAndFinancialProjectionAttachment_Content(long id)
        {
            PAM_HistoricalFinancialAndFinancialProjection_Services svc = new PAM_HistoricalFinancialAndFinancialProjection_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_SupplementalProcurementAndInsuranceAttachment> List_PAM_SupplementalProcurementAndInsuranceAttachment(long pamId)
        {
            PAM_SupplementalProcurementAndInsurance_Services svc = new PAM_SupplementalProcurementAndInsurance_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_SupplementalProcurementAndInsuranceAttachment_Content(long id)
        {
            PAM_SupplementalProcurementAndInsurance_Services svc = new PAM_SupplementalProcurementAndInsurance_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciplesAttachment> List_PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciplesAttachment(long pamId)
        {
            PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciples_Services svc = new PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciples_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciplesAttachment_Content(long id)
        {
            PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciples_Services svc = new PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciples_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_GroupStructureAttachment> List_PAM_GroupStructureAttachment(long pamId)
        {
            PAM_GroupStructure_Services svc = new PAM_GroupStructure_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_GroupStructureAttachment_Content(long id)
        {
            PAM_GroupStructure_Services svc = new PAM_GroupStructure_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_TermSheetAttachment> List_PAM_TermSheetAttachment(long pamId)
        {
            PAM_TermSheet_Services svc = new PAM_TermSheet_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_TermSheetAttachment_Content(long id)
        {
            PAM_TermSheet_Services svc = new PAM_TermSheet_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_RiskRatingAttachment> List_PAM_RiskRatingAttachment(long pamId)
        {
            PAM_RiskRating_Services svc = new PAM_RiskRating_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_RiskRatingAttachment_Content(long id)
        {
            PAM_RiskRating_Services svc = new PAM_RiskRating_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_KYCChecklistsAttachment> List_PAM_KYCChecklistsAttachment(long pamId)
        {
            PAM_KYCChecklists_Services svc = new PAM_KYCChecklists_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_KYCChecklistsAttachment_Content(long id)
        {
            PAM_KYCChecklists_Services svc = new PAM_KYCChecklists_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_OtherBanksFacilitiesOrSummaryOfPefindoReportAttachment> List_PAM_OtherBanksFacilitiesOrSummaryOfPefindoReportAttachment(long pamId)
        {
            PAM_OtherBanksFacilitiesOrSummaryOfPefindoReport_Services svc = new PAM_OtherBanksFacilitiesOrSummaryOfPefindoReport_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_OtherBanksFacilitiesOrSummaryOfPefindoReportAttachment_Content(long id)
        {
            PAM_OtherBanksFacilitiesOrSummaryOfPefindoReport_Services svc = new PAM_OtherBanksFacilitiesOrSummaryOfPefindoReport_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_IndustryAnalysisAttachment> List_PAM_IndustryAnalysisAttachment(long pamId)
        {
            PAM_IndustryAnalysis_Services svc = new PAM_IndustryAnalysis_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_IndustryAnalysisAttachment_Content(long id)
        {
            PAM_IndustryAnalysis_Services svc = new PAM_IndustryAnalysis_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_LegalDueDiligenceReportAttachment> List_PAM_LegalDueDiligenceReportAttachment(long pamId)
        {
            PAM_LegalDueDiligenceReport_Services svc = new PAM_LegalDueDiligenceReport_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_LegalDueDiligenceReportAttachment_Content(long id)
        {
            PAM_LegalDueDiligenceReport_Services svc = new PAM_LegalDueDiligenceReport_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_SAndEDueDiligenceAttachment> List_PAM_SAndEDueDiligenceAttachment(long pamId)
        {
            PAM_SAndEDueDiligence_Services svc = new PAM_SAndEDueDiligence_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void GetS_PAM_AndEDueDiligenceAttachment_Content(long id)
        {
            PAM_SAndEDueDiligence_Services svc = new PAM_SAndEDueDiligence_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_ShareValuationReportAttachment> List_PAM_ShareValuationReportAttachment(long pamId)
        {
            PAM_ShareValuationReport_Services svc = new PAM_ShareValuationReport_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_ShareValuationReportAttachment_Content(long id)
        {
            PAM_ShareValuationReport_Services svc = new PAM_ShareValuationReport_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_OtherReportsAttachment> List_PAM_OtherReportsAttachment(long pamId)
        {
            PAM_OtherReports_Services svc = new PAM_OtherReports_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_OtherReportsAttachment_Content(long id)
        {
            PAM_OtherReports_Services svc = new PAM_OtherReports_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }
        
        [WebMethod]
        public List<PAM_WorkingPaperAttachment> List_PAM_WorkingPaperAttachment_Previous(long pamId, string snWhenAdded_NOT)
        {
            PAM_WorkingPaper_Services svc = new PAM_WorkingPaper_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId, null, null, snWhenAdded_NOT);
        }

        [WebMethod]
        public void Get_PAM_WorkingPaperAttachment_Content(long id)
        {
            PAM_WorkingPaper_Services svc = new PAM_WorkingPaper_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_OtherSupportingDocumentAttachment> List_PAM_OtherSupportingDocumentAttachment(long pamId, int roleIdWhenAdded, string snWhenAdded_NOT)
        {
            PAM_OtherSupportingDocument_Services svc = new PAM_OtherSupportingDocument_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId, null, roleIdWhenAdded, snWhenAdded_NOT);
        }

        [WebMethod]
        public void Get_PAM_OtherSupportingDocumentAttachment_Content(long id)
        {
            PAM_OtherSupportingDocument_Services svc = new PAM_OtherSupportingDocument_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_MergedDocumentResultAttachment> List_PAM_MergedDocumentResultAttachment(long pamId)
        {
            PAM_MergedDocumentResult_Services svc = new PAM_MergedDocumentResult_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListMergedDocumentResultAttachment(pamId);
        }

        [WebMethod]
        public void Get_PAM_MergedDocumentResultAttachment_Content(long id)
        {
            PAM_MergedDocumentResult_Services svc = new PAM_MergedDocumentResult_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_DocumentRelatedToBoDDecisionAttachment> List_PAM_DocumentRelatedToBoDDecisionAttachment_Previous(long pamId, string snWhenAdded_NOT)
        {
            PAM_DocumentRelatedToBoDDecision_Services svc = new PAM_DocumentRelatedToBoDDecision_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId, snWhenAdded_NOT);
        }

        [WebMethod]
        public void Get_PAM_DocumentRelatedToBoDDecisionAttachment_Content(long id)
        {
            PAM_DocumentRelatedToBoDDecision_Services svc = new PAM_DocumentRelatedToBoDDecision_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<PAM_DocumentRelatedToBoCDecisionAttachment> List_PAM_DocumentRelatedToBoCDecisionAttachment_Previous(long pamId, string snWhenAdded_NOT)
        {
            PAM_DocumentRelatedToBoCDecision_Services svc = new PAM_DocumentRelatedToBoCDecision_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(pamId, snWhenAdded_NOT);
        }

        [WebMethod]
        public void Get_PAM_DocumentRelatedToBoCDecisionAttachment_Content(long id)
        {
            PAM_DocumentRelatedToBoCDecision_Services svc = new PAM_DocumentRelatedToBoCDecision_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }
        
    }
}