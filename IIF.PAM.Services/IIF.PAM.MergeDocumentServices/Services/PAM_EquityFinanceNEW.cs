﻿using System;
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
    public class PAM_EquityFinanceNEW : BaseServices
    {
        public FileMergeResult MergePAMEquityFinance(SqlConnection con, long pamId, string foldertemplate, string temporaryFolderLocation)
        {
            DBHelper db = new DBHelper();

			//foldertemplate = ConfigurationManager.AppSettings["PAM_TEMPLATE_FOLDER_LOCATION"];
			//temporaryFolderLocation = ConfigurationManager.AppSettings["PAM_MERGE_FOLDER_LOCATION"];

			List<PAMData> dataResult = db.ExecToModel<PAMData>(con, "dbo.Generate_Document_PAM_Data_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

            System.Data.DataTable listBorrower = db.ExecToDataTable(con, "Generate_Document_PAM_Borrower_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });
			System.Data.DataTable listBorrowerCover = db.ExecToDataTable(con, "Generate_Document_PAM_Borrower_Cover_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

			System.Data.DataTable listFacility = db.ExecToDataTable(con, "Generate_Document_PAM_ProposalFacility_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

            System.Data.DataTable listDealTeam = db.ExecToDataTable(con, "Generate_Document_PAM_DealTeam_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

			System.Data.DataTable listDocVersion = db.ExecToDataTable(con, "Generate_Document_PAM_Version_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

			string fileName = "PAM-" + dataResult[0].ProductType + "-" + dataResult[0].ProjectCompanyName + "-" + dataResult[0].ProjectCode + ".docx";
            string fileNamePDF = "PAM-" + dataResult[0].ProductType + "-" + dataResult[0].ProjectCompanyName + "-" + dataResult[0].ProjectCode + ".pdf";
            string fileTemplateName = "PAM Equity Investment Template.docx";
            string fileTemplateFullName = foldertemplate.AppendPath("\\", fileTemplateName);

            string getfileName = Path.GetFileName(fileTemplateFullName);
            string destFile = Path.Combine(temporaryFolderLocation, fileName);

			IIFCommon.copyFromNetwork(fileTemplateFullName, destFile, foldertemplate, temporaryFolderLocation);

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
					int countSetBorrower = 0;
					string prevBorrower = "";
					string currentBorrower = "";
					List<String> lsBorrower = new List<string>();

					foreach (DataRow item in listBorrowerCover.Rows)
					{
						countBorrower++;
						prevBorrower = item[0].ToString().Trim().ToLower();

						if (!lsBorrower.Contains(prevBorrower))
						{
							if (countBorrower > 5)
							{
								continue;
							}
							countSetBorrower++;
							app.ActiveDocument.Bookmarks["CompanyName" + (countSetBorrower)].Range.Text = item[0].ToString();

							currentBorrower = item[0].ToString().Trim().ToLower();

							lsBorrower.Add(currentBorrower);
						}
					}


					app.ActiveDocument.Bookmarks["ProjectName"].Range.Text = dataResult[0].ProjectName;

					app.ActiveDocument.Bookmarks["ProjectCode"].Range.Text = dataResult[0].ProjectCode;
					System.Globalization.CultureInfo cult = new System.Globalization.CultureInfo("en-us");
					string dateToShow = string.Format(cult, "{0:dd-MMMM-yyyy}", dataResult[0].PAMDate);					
					app.ActiveDocument.Bookmarks["ProjectDate"].Range.Text = dateToShow;					

					//app.ActiveDocument.Bookmarks["FooterProjectCode"].Range.Text = dataResult[0].ProjectCode;
					#endregion

					#region PROJECT
					app.ActiveDocument.Bookmarks["AxPROJECTxProjectName"].Range.Text = dataResult[0].ProjectName;
					app.ActiveDocument.Bookmarks["AxPROJECTxSectorSubsector"].Range.Text = dataResult[0].SectorDesc + " - " + dataResult[0].SubSectorDesc;					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "AxPROJECTxFundingNeeds", AppConstants.TableName.PAM_ProjectData, pamId, "FundingNeeds", "Id");					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "AxPROJECTxDealStrategy", AppConstants.TableName.PAM_ProjectData, pamId, "DealStrategy", "Id");
					#endregion

					#region BORROWER
					app.ActiveDocument.Bookmarks["BxBORROWERxInvesteeCompany"].Range.Text = dataResult[0].ProjectCompanyName;

					Table tblShareholders = IIFCommon.createTable(app, "BxBORROWERxShareholders", 3, true);
					//header
					tblShareholders.Cell(1, 1).Range.Text = "Target Company";
					tblShareholders.Cell(1, 2).Range.Text = "Shareholders";
					tblShareholders.Cell(1, 3).Range.Text = "% ownership";

					string prevkey = "";
					string cellText = "";
					int rowCounter = 1;
					int rowTemp = 0;
					foreach (DataRow item in listBorrower.Rows)
					{
						tblShareholders.Rows.Add(ref missing);
						rowCounter++;

						prevkey = item[0].ToString().Trim().ToLower();
						if (cellText.Trim().ToLower() != prevkey.Trim().ToLower())
						{
							//merge kolom kalo value nya beda, mulai row ke 3
							if (rowCounter > 2 && (rowTemp != (rowCounter - 1)))
							{
								tblShareholders.Cell(rowTemp, 1).Merge(tblShareholders.Cell(rowCounter - 1, 1));
							}

							rowTemp = rowCounter;

							tblShareholders.Cell(rowCounter, 1).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
							tblShareholders.Cell(rowCounter, 1).Range.Text = item[0].ToString().Trim();
							cellText = item[0].ToString().Trim().ToLower();
							tblShareholders.Cell(rowCounter, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
						}

						tblShareholders.Cell(rowCounter, 2).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
						tblShareholders.Cell(rowCounter, 2).Range.Text = item[1].ToString().Trim();
						tblShareholders.Cell(rowCounter, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

						tblShareholders.Cell(rowCounter, 3).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
						tblShareholders.Cell(rowCounter, 3).Range.Text = item[2].ToString().Trim();
						tblShareholders.Cell(rowCounter, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					}

					//merge kolom untuk value trakhir
					if (rowCounter > 2 && (rowTemp != rowCounter))
					{
						tblShareholders.Cell(rowTemp, 1).Merge(tblShareholders.Cell(rowCounter, 1));
					}

					app.ActiveDocument.Bookmarks["BxBORROWERxUltimateBeneficialOwner"].Range.Text = dataResult[0].UltimateBeneficialOwner;
					
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxSP"].Range.Text = dataResult[0].SAndPRate;
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxMoodys"].Range.Text = dataResult[0].MoodysRate;
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxFitch"].Range.Text = dataResult[0].FitchRate;
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxPefindo"].Range.Text = dataResult[0].PefindoRate;
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxSAndECategory"].Range.Text = dataResult[0].SAndECategoryRate + "-" + dataResult[0].SAndECategoryType;
					app.ActiveDocument.Bookmarks["BxBORROWERxRatingxLQCBIChecking"].Range.Text = dataResult[0].LQCOrBICheckingRate;
					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "BxBORROWERxBusinessActivities", AppConstants.TableName.PAM_BorrowerOrTargetCompanyData, pamId, "BusinessActivities", "Id");
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "BxBORROWERxOtherInformation", AppConstants.TableName.PAM_BorrowerOrTargetCompanyData, pamId, "OtherInformation", "Id");
					#endregion
					
					#region PROPOSAL					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "CxPROPOSALxPurpose", AppConstants.TableName.PAM_ProposalData, pamId, "Purpose", "Id");

					app.ActiveDocument.Bookmarks["CxPROPOSALxApprovalAuthority"].Range.Text = dataResult[0].ApprovalAuthority;

					Table tblFacility = IIFCommon.createTable(app, "CxPROPOSALxInvestment", 2, true);
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

					app.ActiveDocument.Bookmarks["CxPROPOSALxGroupExposure"].Range.Text = dataResult[0].GroupExposureCurr + " " + dataResult[0].GroupExposureAmount;
					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "CxPROPOSALxRemarks", AppConstants.TableName.PAM_ProposalData, pamId, "Remarks", "Id");
					app.ActiveDocument.Bookmarks["CxPROPOSALxExpectedHoldingPeriod"].Range.Text = dataResult[0].ExpectedHoldingPeriodYear + " Year(s) " + dataResult[0].ExpectedHoldingPeriodMonth + " Month(s)";					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "CxPROPOSALxExitStrategy", AppConstants.TableName.PAM_ProposalData, pamId, "ExitStrategy", "Id");					
					this.FillBookmarkWithPAMAttachmentABNormal(app, con, "CxPROPOSALxExpectedReturn", AppConstants.TableName.PAM_ProposalData, pamId, "ExpectedReturn", "Id");					
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
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "TermSheet", AppConstants.TableName.PAM_TermSheet, pamId);
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "RiskRating", AppConstants.TableName.PAM_RiskRating, pamId);
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "KYCChecklists", AppConstants.TableName.PAM_KYCChecklists, pamId);
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "OtherBanksfacilities", AppConstants.TableName.PAM_OtherBanksFacilities, pamId);
					this.FillBookmarkWithPAMAttachmentNormal(app, con, "ShareValuationReport", AppConstants.TableName.PAM_ShareValuationReport, pamId);

					System.Data.DataTable listLegalDue = db.ExecToDataTable(con, "Generate_Document_PAM_LegalDue_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });
					IIFCommon.createLegalSAndEDueOtherReportTable(app, listLegalDue, "LegalDuediligenceReportAttachment", "LegalDuediligenceReportDescription", currFontFamily, currFontSize);

					System.Data.DataTable listSAndDue = db.ExecToDataTable(con, "Generate_Document_PAM_SAndDue_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });
					IIFCommon.createLegalSAndEDueOtherReportTable(app, listSAndDue, "SAndDuediligenceReportAttachment", "SAndDuediligenceReportDescription", currFontFamily, currFontSize);

					System.Data.DataTable listOtherReport = db.ExecToDataTable(con, "Generate_Document_PAM_OtherReport_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });
					IIFCommon.createLegalSAndEDueOtherReportTable(app, listOtherReport, "OtherReportAttachment", "OtherReportDescription", currFontFamily, currFontSize);
					#endregion

					IIFCommon.finalizeDoc(doc);
					IIFCommon.injectFooterPAM(doc, dataResult[0].ProjectCode);

					bool isPreview = false;
					try
					{
						if (dataResult[0].MWorkflowStatusId != null && dataResult[0].MWorkflowStatusId == 7)
							isPreview = true;
					}
					catch { }
					fileNamePDF = IIFCommon.fileNameFormat(listDocVersion, fileNamePDF, isPreview);

					doc.SaveAs2(Path.Combine(temporaryFolderLocation, fileNamePDF), WdExportFormat.wdExportFormatPDF);					
					//doc.SaveAs2(Path.Combine(temporaryFolderLocation, fileName));
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

			File.Delete(destFile);
			string destFilePDF = Path.Combine(temporaryFolderLocation, fileNamePDF);			
			byte[] fileContent = File.ReadAllBytes(destFilePDF);

			FileMergeResult result = new FileMergeResult();
			result.FileContent = fileContent;
			result.FileName = fileNamePDF;			
			return result;
		}
    }
}
