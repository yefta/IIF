using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;

using Microsoft.Office.Interop.Word;

using IIF.PAM.MergeDocumentServices.Helper;
using IIF.PAM.MergeDocumentServices.Models;
using System.Configuration;

namespace IIF.PAM.MergeDocumentServices.Services
{
    public class PAM_ProjectFinanceNEW : BaseServices
    {
        public FileMergeResult MergePAMProjectFinance(SqlConnection con, long pamId, string foldertemplate, string temporaryFolderLocation)
        {
            DBHelper db = new DBHelper();

			//foldertemplate = System.Configuration.ConfigurationManager.AppSettings["PAM_TEMPLATE_FOLDER_LOCATION"];
			//temporaryFolderLocation = System.Configuration.ConfigurationManager.AppSettings["PAM_MERGE_FOLDER_LOCATION"];

			List<PAMData> dataResult = db.ExecToModel<PAMData>(con, "dbo.Generate_Document_PAM_Data_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

            System.Data.DataTable listBorrower = db.ExecToDataTable(con, "Generate_Document_PAM_Borrower_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });
			System.Data.DataTable listBorrowerCover = db.ExecToDataTable(con, "Generate_Document_PAM_Borrower_Cover_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

			System.Data.DataTable listFacility = db.ExecToDataTable(con, "Generate_Document_PAM_ProposalFacility_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

            System.Data.DataTable listDealTeam = db.ExecToDataTable(con, "Generate_Document_PAM_DealTeam_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

            string fileName = "PAM-" + dataResult[0].ProductType + "-" + dataResult[0].ProjectCompanyName + "-" + dataResult[0].ProjectCode + ".docx";
            string fileNamePDF = "PAM-" + dataResult[0].ProductType + "-" + dataResult[0].ProjectCompanyName + "-" + dataResult[0].ProjectCode + ".pdf";
            string fileTemplateName = "PAM Project Finance Template.docx";
            string fileTemplateFullName = foldertemplate.AppendPath("\\", fileTemplateName);

            string getfileName = Path.GetFileName(fileTemplateFullName);
            string destFile = Path.Combine(temporaryFolderLocation, fileName);
            File.Copy(fileTemplateFullName, destFile, true);

            object missing = System.Reflection.Missing.Value;
            object readOnly = (object)false;
            Application app = new Application();

			string currFontFamily = "Roboto Light";
			float currFontSize = 10;
			try
			{
				Document doc = app.Documents.Open(destFile, ref missing, ref readOnly);
				app.Visible = false;
				try
				{
					
					#region Cover     
					//app.ActiveDocument.Bookmarks["CompanyName"].Range.Text = dataResult[0].ProjectCompanyName;

					int countBorrower = 0;
					string prevBorrower = "";
					string currentBorrower = "";					
					foreach (DataRow item in listBorrowerCover.Rows)
					{
						countBorrower++;
						prevBorrower = item[0].ToString();
						if (currentBorrower != prevBorrower)
						{
							app.ActiveDocument.Bookmarks["CompanyName"].Range.Text = item[0].ToString();
							currentBorrower = item[0].ToString();

							if (countBorrower < listBorrowerCover.Rows.Count)
								app.ActiveDocument.Bookmarks["CompanyName"].Range.Text = System.Environment.NewLine;
						}												
					}					

					app.ActiveDocument.Bookmarks["ProjectName"].Range.Text = dataResult[0].ProjectName;

					
					app.ActiveDocument.Bookmarks["ProjectCode"].Range.Text = dataResult[0].ProjectCode;
					app.ActiveDocument.Bookmarks["ProjectDate"].Range.Text = dataResult[0].PAMDate.ToString("dd-MMMM-yyyy");

					app.ActiveDocument.Bookmarks["FooterProjectCode"].Range.Text = dataResult[0].ProjectCode;
					#endregion

					#region PROJECT
					app.ActiveDocument.Bookmarks["AxPROJECTxProjectDescription"].Range.Text = dataResult[0].ProjectDescription;
					app.ActiveDocument.Bookmarks["AxPROJECTxSectorSubsector"].Range.Text = dataResult[0].SectorDesc + " - " + dataResult[0].SubSectorDesc;
					app.ActiveDocument.Bookmarks["AxPROJECTxProjectCost"].Range.Text = dataResult[0].ProjectCostCurr + " " + dataResult[0].ProjectCostAmount;
					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "AxPROJECTxProjectScope", AppConstants.TableName.PAM_ProjectData, pamId, "ProjectScope", "Id");					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "AxPROJECTxProjectStructure", AppConstants.TableName.PAM_ProjectData, pamId, "ProjectStructure", "Id");					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "AxPROJECTxDealStrategy", AppConstants.TableName.PAM_ProjectData, pamId, "DealStrategy", "Id");
					#endregion

					#region BORROWER
					app.ActiveDocument.Bookmarks["BxBORROWERxProjectCompany"].Range.Text = dataResult[0].ProjectCompanyName;

					Table tblProjectSponsors = IIFCommon.createTable(app, "BxBORROWERxProjectSponsors", 3, true);
					//header
					tblProjectSponsors.Cell(1, 1).Range.Text = "Project Company";
					tblProjectSponsors.Cell(1, 2).Range.Text = "Project Sponsors";
					tblProjectSponsors.Cell(1, 3).Range.Text = "% ownership";

					string prevkey = "";
					string cellText = "";
					int rowCounter = 1;
					int rowTemp = 0;
					foreach (DataRow item in listBorrower.Rows)
					{
						tblProjectSponsors.Rows.Add(ref missing);
						rowCounter++;

						prevkey = item[0].ToString();
						if (cellText != prevkey)
						{
							//merge kolom kalo value nya beda, mulai row ke 3
							if (rowCounter > 2 && (rowTemp != (rowCounter - 1)))
							{
								tblProjectSponsors.Cell(rowTemp, 1).Merge(tblProjectSponsors.Cell(rowCounter - 1, 1));
							}

							rowTemp = rowCounter;

							tblProjectSponsors.Cell(rowCounter, 1).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
							tblProjectSponsors.Cell(rowCounter, 1).Range.Text = item[0].ToString();
							cellText = item[0].ToString();
							tblProjectSponsors.Cell(rowCounter, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
						}

						tblProjectSponsors.Cell(rowCounter, 2).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
						tblProjectSponsors.Cell(rowCounter, 2).Range.Text = item[1].ToString();
						tblProjectSponsors.Cell(rowCounter, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

						tblProjectSponsors.Cell(rowCounter, 3).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
						tblProjectSponsors.Cell(rowCounter, 3).Range.Text = item[2].ToString();
						tblProjectSponsors.Cell(rowCounter, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					}

					//merge kolom untuk value trakhir
					if (rowCounter > 2 && (rowTemp != rowCounter))
					{
						tblProjectSponsors.Cell(rowTemp, 1).Merge(tblProjectSponsors.Cell(rowCounter, 1));
					}

					app.ActiveDocument.Bookmarks["BxBORROWERxUltimateBeneficialOwner"].Range.Text = dataResult[0].UltimateBeneficialOwner;
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxRating"].Range.Text = dataResult[0].IIFRate;
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxRatingDate"].Range.Text = Convert.ToDateTime(dataResult[0].IIFRatingDate).ToString("dd MMM yyyy");
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxSP"].Range.Text = dataResult[0].SAndPRate;
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxMoodys"].Range.Text = dataResult[0].MoodysRate;
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxFitch"].Range.Text = dataResult[0].FitchRate;
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxPefindo"].Range.Text = dataResult[0].PefindoRate;
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxSAndECategory"].Range.Text = dataResult[0].SAndECategoryRate + "-" + dataResult[0].SAndECategoryType;
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxLQCBIChecking"].Range.Text = dataResult[0].LQCOrBICheckingRate;
										
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "BxBORROWERxOtherInformation", AppConstants.TableName.PAM_BorrowerOrTargetCompanyData, pamId, "OtherInformation", "Id");
					#endregion

					#region PROPOSAL					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "CxPROPOSALxPurpose", AppConstants.TableName.PAM_ProposalData, pamId, "Purpose", "Id");
					app.ActiveDocument.Bookmarks["CxPROPOSALxApprovalAuthority"].Range.Text = dataResult[0].ApprovalAuthority;

					Table tblFacility = IIFCommon.createTable(app, "CxPROPOSALxFacility", 2, true);
					//header
					tblFacility.Cell(1, 1).Range.Text = "Type";
					tblFacility.Cell(1, 2).Range.Text = "Amount";
					
					rowCounter = 1;

					foreach (DataRow item in listFacility.Rows)
					{
						tblFacility.Rows.Add(ref missing);
						rowCounter++;

						tblFacility.Cell(rowCounter, 1).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
						tblFacility.Cell(rowCounter, 1).Range.Text = item[0].ToString();
						tblFacility.Cell(rowCounter, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

						tblFacility.Cell(rowCounter, 2).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
						tblFacility.Cell(rowCounter, 2).Range.Text = item[1].ToString() + " " + item[2].ToString();
						tblFacility.Cell(rowCounter, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					}

					tblFacility.Rows.Add(ref missing);
					tblFacility.Cell(rowCounter + 1, 1).Range.Text = "Remarks : " + dataResult[0].FacilityOrInvestmentRemarks;
					tblFacility.Cell(rowCounter + 1, 1).Merge(tblFacility.Cell(rowCounter + 1, 2));
					tblFacility.Cell(rowCounter + 1, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

					app.ActiveDocument.Bookmarks["CxPROPOSALxGroupExposure"].Range.Text = dataResult[0].GroupExposureCurr + " " +  dataResult[0].GroupExposureAmount;
					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "CxPROPOSALxRemarks", AppConstants.TableName.PAM_ProposalData, pamId, "Remarks", "Id");

					app.ActiveDocument.Bookmarks["CxPROPOSALxTenor"].Range.Text = dataResult[0].tenorYear + " year(s)  " + dataResult[0].tenorMonth + " month(s)";
					app.ActiveDocument.Bookmarks["CxPROPOSALxAverageLoanLife"].Range.Text = dataResult[0].averageLoanLifeYear + " year(s)  " + dataResult[0].averageLoanLifeMonth + " month(s)";
					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "CxPROPOSALxPricingxInterestRate", AppConstants.TableName.PAM_ProposalData, pamId, "PricingInterestRate", "Id");

					app.ActiveDocument.Bookmarks["CxPROPOSALxPricingxCommitmentFee"].Range.Text = dataResult[0].pricingCommitmentFee;
					app.ActiveDocument.Bookmarks["CxPROPOSALxPricingxFacility"].Range.Text = dataResult[0].pricingUpfrontFacilityFee;
					app.ActiveDocument.Bookmarks["CxPROPOSALxPricingxStructuringFee"].Range.Text = dataResult[0].pricingStructuringFee;
					app.ActiveDocument.Bookmarks["CxPROPOSALxPricingxArrangerFee"].Range.Text = dataResult[0].pricingArrangerFee;
					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "CxPROPOSALxCollateral", AppConstants.TableName.PAM_ProposalData, pamId, "PricingCollateral", "Id");					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "CxPROPOSALxOtherCondition", AppConstants.TableName.PAM_ProposalData, pamId, "PricingOtherConditions", "Id");
					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "CxPROPOSALxExceptionToIIFPolicy", AppConstants.TableName.PAM_ProposalData, pamId, "PricingExceptionToIIFPolicy", "Id");

					app.ActiveDocument.Bookmarks["CxPROPOSALxReviewPeriod"].Range.Text = dataResult[0].reviewPeriod;
					#endregion

					#region RECOMMENDATION					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "DxRECOMMENDATIONxKeyInvestment", AppConstants.TableName.PAM_RecommendationData, pamId, "KeyInvestmentRecommendation", "Id");					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "DxRECOMMENDATION", AppConstants.TableName.PAM_RecommendationData, pamId, "Recommendation", "Id");

					Table tblDealTeam = IIFCommon.createTable(app, "DxRECOMMENDATIONxDealTeam", 1, false);
					tblDealTeam.Borders.Enable = 0;
					rowCounter = 0;
					foreach (DataRow item in listDealTeam.Rows)
					{
						tblDealTeam.Rows.Add(ref missing);
						rowCounter++;
						tblDealTeam.Cell(rowCounter, 1).Range.Text = item[0].ToString();
						tblDealTeam.Cell(rowCounter, 1).Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
						tblDealTeam.Cell(rowCounter, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
					}

					app.ActiveDocument.Bookmarks["DxRECOMMENDATIONxCIO"].Range.Text = dataResult[0].AccountResponsibleCIOName;

					#endregion
				
					

					#region Attachment 
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "ProjectAnalysis", AppConstants.TableName.PAM_ProjectAnalysis, pamId);
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "HistoricalFinancialandFinancialProject", AppConstants.TableName.PAM_HistoricalFinancial, pamId);
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "Supplemental", AppConstants.TableName.PAM_Supplemental, pamId);
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "SocialEnvironmental", AppConstants.TableName.PAM_Social, pamId);
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "GroupStructure", AppConstants.TableName.PAM_GroupStructure, pamId);
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "TermSheet", AppConstants.TableName.PAM_TermSheet, pamId);
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "RiskRating", AppConstants.TableName.PAM_RiskRating, pamId);
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "KYCChecklists", AppConstants.TableName.PAM_KYCChecklists, pamId);
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "OtherBanksfacilities", AppConstants.TableName.PAM_OtherBanksFacilities, pamId);
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "IndustryAnalysis", AppConstants.TableName.PAM_IndustryAnalysis, pamId);
					
					System.Data.DataTable listLegalDue = db.ExecToDataTable(con, "Generate_Document_PAM_LegalDue_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });
					IIFCommon.createLegalSAndEDueOtherReportTable(app, listLegalDue, "LegalDuediligenceReportAttachment", "LegalDuediligenceReportDescription");
					
					System.Data.DataTable listSAndDue = db.ExecToDataTable(con, "Generate_Document_PAM_SAndDue_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });
					IIFCommon.createLegalSAndEDueOtherReportTable(app, listSAndDue, "SAndDuediligenceReportAttachment", "SAndDuediligenceReportDescription");
					
					System.Data.DataTable listOtherReport = db.ExecToDataTable(con, "Generate_Document_PAM_OtherReport_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });
					IIFCommon.createLegalSAndEDueOtherReportTable(app, listOtherReport, "OtherReportAttachment", "OtherReportDescription");					
					#endregion

					IIFCommon.finalizeDoc(doc);										

					//doc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
					//doc.SaveAs2(Path.Combine(temporaryFolderLocation, fileNamePDF), WdExportFormat.wdExportFormatPDF);					
					doc.SaveAs2(Path.Combine(temporaryFolderLocation, fileName));					
				}
				catch (Exception ex)
				{
					throw ex;
				}
				finally
				{
					doc.Close(WdSaveOptions.wdDoNotSaveChanges);					
				}
			}
			finally
			{
				app.Quit();
			}

			//File.Delete(destFile);
			//string destFilePDF = Path.Combine(temporaryFolderLocation, fileNamePDF);			
			//byte[] fileContent = File.ReadAllBytes(destFilePDF);

			FileMergeResult result = new FileMergeResult();
			//result.FileContent = fileContent;
			//result.FileName = fileNamePDF;			
			return result;
		}     
    }
}