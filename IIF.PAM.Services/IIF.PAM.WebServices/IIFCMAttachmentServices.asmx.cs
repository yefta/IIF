using System.Collections.Generic;
using System.ComponentModel;
using System.Web.Services;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;
using IIF.PAM.WebServices.Services;

namespace IIF.PAM.WebServices
{
    /// <summary>
    /// Summary description for IIFCMAttachmentServices
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class IIFCMAttachmentServices : BaseWebService
    {
        [WebMethod]
        public List<CM_PeriodicReviewAttachment> List_CM_PeriodicReviewAttachment(long cmId)
        {
            CM_PeriodicReview_Services svc = new CM_PeriodicReview_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(cmId);
        }

        [WebMethod]
        public void Get_CM_PeriodicReviewAttachment_Content(long id)
        {
            CM_PeriodicReview_Services svc = new CM_PeriodicReview_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<CM_CreditMemorandumAttachment> List_CM_CreditMemorandumAttachment(long cmId)
        {
            CM_CreditMemorandum_Services svc = new CM_CreditMemorandum_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(cmId);
        }

        [WebMethod]
        public void Get_CM_CreditMemorandumAttachment_Content(long id)
        {
            CM_CreditMemorandum_Services svc = new CM_CreditMemorandum_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<CM_RiskRatingAttachment> List_CM_RiskRatingAttachment(long cmId)
        {
            CM_RiskRating_Services svc = new CM_RiskRating_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(cmId);
        }

        [WebMethod]
        public void Get_CM_RiskRatingAttachment_Content(long id)
        {
            CM_RiskRating_Services svc = new CM_RiskRating_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<CM_KYCChecklistsAttachment> List_CM_KYCChecklistsAttachment(long cmId)
        {
            CM_KYCChecklists_Services svc = new CM_KYCChecklists_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(cmId);
        }

        [WebMethod]
        public void Get_CM_KYCChecklistsAttachment_Content(long id)
        {
            CM_KYCChecklists_Services svc = new CM_KYCChecklists_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<CM_SAndEReviewAttachment> List_CM_SAndEReviewAttachment(long cmId)
        {
            CM_SAndEReview_Services svc = new CM_SAndEReview_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(cmId);
        }

        [WebMethod]
        public void Get_CM_SAndEReviewAttachment_Content(long id)
        {
            CM_SAndEReview_Services svc = new CM_SAndEReview_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<CM_OtherBanksFacilitiesOrSummaryOfPefindoReportAttachment> List_CM_OtherBanksFacilitiesOrSummaryOfPefindoReportAttachment(long cmId)
        {
            CM_OtherBanksFacilitiesOrSummaryOfPefindoReport_Services svc = new CM_OtherBanksFacilitiesOrSummaryOfPefindoReport_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(cmId);
        }

        [WebMethod]
        public void Get_CM_OtherBanksFacilitiesOrSummaryOfPefindoReportAttachment_Content(long id)
        {
            CM_OtherBanksFacilitiesOrSummaryOfPefindoReport_Services svc = new CM_OtherBanksFacilitiesOrSummaryOfPefindoReport_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<CM_ValuationReportAttachment> List_CM_ValuationReportAttachment(long cmId)
        {
            CM_ValuationReport_Services svc = new CM_ValuationReport_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(cmId);
        }

        [WebMethod]
        public void Get_CM_ValuationReportAttachment_Content(long id)
        {
            CM_ValuationReport_Services svc = new CM_ValuationReport_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<CM_OtherAttachmentOrReportsAttachment> List_CM_OtherAttachmentOrReportsAttachment(long cmId)
        {
            CM_OtherAttachmentOrReports_Services svc = new CM_OtherAttachmentOrReports_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(cmId);
        }

        [WebMethod]
        public void Get_CM_OtherAttachmentOrReportsAttachment_Content(long id)
        {
            CM_OtherAttachmentOrReports_Services svc = new CM_OtherAttachmentOrReports_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<CM_WorkingPaperAttachment> List_CM_WorkingPaperAttachment_Previous(long CMId, string snWhenAdded_NOT)
        {
            CM_WorkingPaper_Services svc = new CM_WorkingPaper_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(CMId, null, null, snWhenAdded_NOT);
        }

        [WebMethod]
        public void Get_CM_WorkingPaperAttachment_Content(long id)
        {
            CM_WorkingPaper_Services svc = new CM_WorkingPaper_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<CM_OtherSupportingDocumentAttachment> List_CM_OtherSupportingDocumentAttachment(long CMId, int roleIdWhenAdded, string snWhenAdded_NOT)
        {
            CM_OtherSupportingDocument_Services svc = new CM_OtherSupportingDocument_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(CMId, null, roleIdWhenAdded, snWhenAdded_NOT);
        }

        [WebMethod]
        public void Get_CM_OtherSupportingDocumentAttachment_Content(long id)
        {
            CM_OtherSupportingDocument_Services svc = new CM_OtherSupportingDocument_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<CM_MergedDocumentResultAttachment> List_CM_MergedDocumentResultAttachment(long CMId)
        {
            CM_MergedDocumentResult_Services svc = new CM_MergedDocumentResult_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListMergedDocumentResultAttachment(CMId);
        }

        [WebMethod]
        public void Get_CM_MergedDocumentResultAttachment_Content(long id)
        {
            CM_MergedDocumentResult_Services svc = new CM_MergedDocumentResult_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<CM_DocumentRelatedToBoDDecisionAttachment> List_CM_DocumentRelatedToBoDDecisionAttachment_Previous(long CMId, string snWhenAdded_NOT)
        {
            CM_DocumentRelatedToBoDDecision_Services svc = new CM_DocumentRelatedToBoDDecision_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(CMId, snWhenAdded_NOT);
        }

        [WebMethod]
        public void Get_CM_DocumentRelatedToBoDDecisionAttachment_Content(long id)
        {
            CM_DocumentRelatedToBoDDecision_Services svc = new CM_DocumentRelatedToBoDDecision_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }

        [WebMethod]
        public List<CM_DocumentRelatedToBoCDecisionAttachment> List_CM_DocumentRelatedToBoCDecisionAttachment_Previous(long CMId, string snWhenAdded_NOT)
        {
            CM_DocumentRelatedToBoCDecision_Services svc = new CM_DocumentRelatedToBoCDecision_Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListAttachment(CMId, snWhenAdded_NOT);
        }

        [WebMethod]
        public void Get_CM_DocumentRelatedToBoCDecisionAttachment_Content(long id)
        {
            CM_DocumentRelatedToBoCDecision_Services svc = new CM_DocumentRelatedToBoCDecision_Services();
            svc.AppConfig = this.AppConfig;
            XDocument xDoc = svc.GetAttachmentContent(id);
            this.ReturnXDocumentFile(xDoc);
        }
        
    }
}
