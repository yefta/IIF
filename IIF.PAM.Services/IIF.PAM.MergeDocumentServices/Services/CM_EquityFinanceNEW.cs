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
    public class CM_EquityFinanceNEW : BaseServices
    {
        public FileMergeResult MergeCMEquityFinance(SqlConnection con, long cmId, string foldertemplate, string temporaryFolderLocation)
        {
            DBHelper db = new DBHelper();

			//foldertemplate = ConfigurationManager.AppSettings["CM_TEMPLATE_FOLDER_LOCATION"];
			//temporaryFolderLocation = ConfigurationManager.AppSettings["CM_MERGE_FOLDER_LOCATION"];

			List<CMData> dataResult = db.ExecToModel<CMData>(con, "dbo.Generate_Document_CM_Data_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, cmId) });

            System.Data.DataTable listBorrower = db.ExecToDataTable(con, "Generate_Document_CM_Borrower_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, cmId) });
			System.Data.DataTable listBorrowerCover = db.ExecToDataTable(con, "Generate_Document_CM_Borrower_Cover_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, cmId) });

			System.Data.DataTable listFacility = db.ExecToDataTable(con, "Generate_Document_CM_ProposalFacility_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, cmId) });

            System.Data.DataTable listDealTeam = db.ExecToDataTable(con, "Generate_Document_CM_DealTeam_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, cmId) });


			string fileName = "CM-" + dataResult[0].ProductType + "-" + dataResult[0].CompanyName + "-" + dataResult[0].ProjectCode + ".docx";
			string fileNamePDF = "CM-" + dataResult[0].ProductType + "-" + dataResult[0].CompanyName + "-" + dataResult[0].ProjectCode + ".pdf";
            string fileTemplateName = "CM Equity Investment Template.docx";
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
					app.ActiveDocument.Bookmarks["Review"].Range.Text = dataResult[0].ReviewMemo;

					//app.ActiveDocument.Bookmarks["CompanyName"].Range.Text = dataResult[0].CompanyName;

					int countBorrower = 0;
					string prevBorrower = "";
					string currentBorrower = "";
					List<String> lsBorrower = new List<string>();
					Table tblCoverBorrower = IIFCommon.createTable(app, "CompanyName", 1, false);
					tblCoverBorrower.Borders.Enable = 0;
					foreach (DataRow item in listBorrowerCover.Rows)
					{
						countBorrower++;
						prevBorrower = item[0].ToString().Trim().ToLower();
						if (!lsBorrower.Contains(prevBorrower))
						{
							if (countBorrower > 1)
							{
								tblCoverBorrower.Rows.Add(ref missing);
							}
							tblCoverBorrower.Cell(countBorrower, 1).Range.Text = item[0].ToString();
							tblCoverBorrower.Cell(countBorrower, 1).Range.Font.Name = "Roboto Light";
							tblCoverBorrower.Cell(countBorrower, 1).Range.Font.Size = 18;
							tblCoverBorrower.Cell(countBorrower, 1).Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
							tblCoverBorrower.Cell(countBorrower, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

							currentBorrower = item[0].ToString().Trim().ToLower();

							lsBorrower.Add(currentBorrower);
						}
					}

					app.ActiveDocument.Bookmarks["ProjectName"].Range.Text = dataResult[0].ProjectName;

					app.ActiveDocument.Bookmarks["CMNumber"].Range.Text = IIFCommon.generateCMNumber(
						dataResult[0].ProjectCode
						, Convert.ToInt32(dataResult[0].CMNumber).ToString("00")
						, dataResult[0].ApprovalAuhority
						, dataResult[0].CMDate.ToString("MMM")
						, dataResult[0].CMDate.ToString("yyyy")
						);

					app.ActiveDocument.Bookmarks["ProjectCode"].Range.Text = dataResult[0].ProjectCode;
					app.ActiveDocument.Bookmarks["ProjectDate"].Range.Text = dataResult[0].CMDate.ToString("dd-MMMM-yyyy");

					app.ActiveDocument.Bookmarks["FooterDate"].Range.Text = dataResult[0].CMDate.ToString("MMM") + " " + dataResult[0].CMDate.ToString("yyyy");
					#endregion

					#region PROJECT
					app.ActiveDocument.Bookmarks["AxPROJECTxProjectName"].Range.Text = dataResult[0].ProjectName;
					app.ActiveDocument.Bookmarks["AxPROJECTxSectorSubsector"].Range.Text = dataResult[0].SectorDesc + " " + dataResult[0].SubSectorDesc;
					this.FillBookmarkWithCMAttachmentABNormal(app, con, "AxPROJECTxFundingNeeds", AppConstants.TableName.CM_ProjectData, cmId, "FundingNeeds", "Id");
					this.FillBookmarkWithCMAttachmentABNormal(app, con, "AxPROJECTxDealStrategy", AppConstants.TableName.CM_ProjectData, cmId, "DealStrategy", "Id");					
					#endregion

					#region BORROWER
					app.ActiveDocument.Bookmarks["BxBORROWERxInvesteeCompany"].Range.Text = dataResult[0].CompanyName;

					Table tblShareholders = IIFCommon.createTable(app, "BxBORROWERxShareholders", 3, true);
					//header
					tblShareholders.Cell(1, 1).Range.Text = "Investee";
					tblShareholders.Cell(1, 2).Range.Text = "Name";
					tblShareholders.Cell(1, 3).Range.Text = "% Ownership";

					string prevkey = "";
					string cellText = "";
					int rowCounter = 1;
					int rowTemp = 0;
					foreach (DataRow item in listBorrower.Rows)
					{
						tblShareholders.Rows.Add(ref missing);
						rowCounter++;

						prevkey = item[0].ToString();
						if (cellText.Trim() != prevkey.Trim())
						{
							//merge kolom kalo value nya beda, mulai row ke 3
							if (rowCounter > 2 && (rowTemp != (rowCounter - 1)))
							{
								tblShareholders.Cell(rowTemp, 1).Merge(tblShareholders.Cell(rowCounter - 1, 1));
							}

							rowTemp = rowCounter;

							tblShareholders.Cell(rowCounter, 1).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
							tblShareholders.Cell(rowCounter, 1).Range.Text = item[0].ToString().Trim();
							cellText = item[0].ToString();
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
					
					this.FillBookmarkWithCMAttachmentABNormal(app, con, "BxBORROWERxBusinessActivities", AppConstants.TableName.CM_BorrowerOrInvesteeCompanyData, cmId, "BusinessActivities", "Id");
					this.FillBookmarkWithCMAttachmentABNormal(app, con, "BxBORROWERxOtherInformation", AppConstants.TableName.CM_BorrowerOrInvesteeCompanyData, cmId, "OtherInformation", "Id");
					#endregion

					#region PROPOSAL
					this.FillBookmarkWithCMAttachmentABNormal(app, con, "CxPROPOSALxPurpose", AppConstants.TableName.CM_ProposalOrFacilityData, cmId, "Purpose", "Id");
					app.ActiveDocument.Bookmarks["CxPROPOSALxApprovalAuthority"].Range.Text = dataResult[0].ApprovalAuhority;

					Table tblFacility = IIFCommon.createTable(app, "CxPROPOSALxInvestment", 4, true);
					//header
					tblFacility.Cell(1, 1).Range.Text = "Type";
					tblFacility.Cell(1, 2).Range.Text = "Approved";
					tblFacility.Cell(1, 3).Range.Text = "Proposed";
					tblFacility.Cell(1, 4).Range.Text = "MTM";
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

						tblFacility.Cell(rowCounter, 3).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
						tblFacility.Cell(rowCounter, 3).Range.Text = item[3].ToString() + " " + item[4].ToString();
						tblFacility.Cell(rowCounter, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

						tblFacility.Cell(rowCounter, 4).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
						tblFacility.Cell(rowCounter, 4).Range.Text = item[5].ToString() + " " + item[6].ToString();
						tblFacility.Cell(rowCounter, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					}

					tblFacility.Rows.Add(ref missing);
					tblFacility.Cell(rowCounter + 1, 1).Range.Text = "Remarks : " + dataResult[0].FacilityOrInvestmentRemarks;
					tblFacility.Cell(rowCounter + 1, 1).Merge(tblFacility.Cell(rowCounter + 1, 4));
					tblFacility.Cell(rowCounter + 1, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

					app.ActiveDocument.Bookmarks["CxPROPOSALxGroupExposure"].Range.Text = dataResult[0].GroupExposureCurr + " " + dataResult[0].GroupExposureAmount;
					this.FillBookmarkWithCMAttachmentABNormal(app, con, "CxPROPOSALxRemarks", AppConstants.TableName.CM_ProposalOrFacilityData, cmId, "Remarks", "Id");
					app.ActiveDocument.Bookmarks["CxPROPOSALxExpectedHoldingPeriod"].Range.Text = dataResult[0].ExpectedHoldingPeriodYear + " Year(s) " + dataResult[0].ExpectedHoldingPeriodMonth + " Month(s)";

					this.FillBookmarkWithCMAttachmentABNormal(app, con, "CxPROPOSALxExpectedReturn", AppConstants.TableName.CM_ProposalOrFacilityData, cmId, "ExpectedReturn", "Id");
					this.FillBookmarkWithCMAttachmentABNormal(app, con, "CxPROPOSALxOtherCondition", AppConstants.TableName.CM_ProposalOrFacilityData, cmId, "PricingOtherConditions", "Id");										

					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexCurrency"].Range.Text = dataResult[0].LimitComplianceCurrency;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexAsFor"].Range.Text = Convert.ToDateTime(dataResult[0].FacilityLimitComplianceMonth.ToString() + "-1985").ToString("MMM") + " " + dataResult[0].FacilityLimitComplianceYear.ToString();
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexRiskRating"].Range.Text = dataResult[0].FacilityLimitComplianceIIFRate;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSecExposure"].Range.Text = dataResult[0].SectorDesc;

					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSPELxML"].Range.Text = dataResult[0].FacilityLimitComplianceSingleProjectExposureMaxLimit;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSPELxML"].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSPELxP"].Range.Text = dataResult[0].FacilityLimitComplianceSingleProjectExposureProposed;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSPELxP"].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSPELxR"].Range.Text = dataResult[0].SingleProjectExposureRemarks;

					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexPxML"].Range.Text = dataResult[0].FacilityLimitComplianceProductMaxLimit;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexPxML"].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexPxP"].Range.Text = dataResult[0].FacilityLimitComplianceProductProposed;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexPxP"].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexPxR"].Range.Text = dataResult[0].ProductRemarks;					

					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexGELxML"].Range.Text = dataResult[0].FacilityLimitComplianceGrupExposureMaxLimit;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexGELxML"].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexGELxP"].Range.Text = dataResult[0].FacilityLimitComplianceGrupExposureProposed;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexGELxP"].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexGELxR"].Range.Text = dataResult[0].GrupExposureRemarks;

					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSExML"].Range.Text = dataResult[0].FacilityLimitComplianceSectorExposureMaxLimit;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSExML"].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSExP"].Range.Text = dataResult[0].FacilityLimitComplianceSectorExposureProposed;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSExP"].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSExR"].Range.Text = dataResult[0].SectorExposureRemarks;

					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexNotes"].Range.Text = dataResult[0].notes;

					this.FillBookmarkWithCMAttachmentABNormal(app, con, "CxPROPOSALxExceptionToIIFPolicy", AppConstants.TableName.CM_ProposalOrFacilityData, cmId, "PricingExceptionToIIFPolicy", "Id");
					app.ActiveDocument.Bookmarks["CxPROPOSALxReviewPeriod"].Range.Text = dataResult[0].ProposalReviewPeriod;
					#endregion

					#region RECOMMENDATION
					this.FillBookmarkWithCMAttachmentABNormal(app, con, "DxRECOMMENDATIONxKeyInvestment", AppConstants.TableName.CM_RecommendationData, cmId, "KeyInvestmentRecommendation", "Id");
					this.FillBookmarkWithCMAttachmentABNormal(app, con, "DxRECOMMENDATION", AppConstants.TableName.CM_RecommendationData, cmId, "Recommendation", "Id");

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
					this.FillBookmarkWithCMAttachmentNormal(app, con, "PeriodicReview", AppConstants.TableName.CM_PeriodicReview, cmId);

					System.Data.DataTable listPreviousApproval = db.ExecToDataTable(con, "Generate_Document_CM_PreviousApproval_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@ProjectCode", SqlDbType.VarChar, dataResult[0].ProjectCode) });
					IIFCommon.createPreviousApproval(app, listPreviousApproval, "PreviousApprovals", dataResult, currFontFamily, currFontSize);

					this.FillBookmarkWithCMAttachmentNormal(app, con, "RiskRating", AppConstants.TableName.CM_RiskRating, cmId);
					this.FillBookmarkWithCMAttachmentNormal(app, con, "KYCChecklists", AppConstants.TableName.CM_KYCChecklists, cmId);
					this.FillBookmarkWithCMAttachmentNormal(app, con, "SandEReview", AppConstants.TableName.CM_SAndEReview, cmId);
					this.FillBookmarkWithCMAttachmentNormal(app, con, "OtherBanksfacilities", AppConstants.TableName.CM_OtherBanksFacilities, cmId);
					this.FillBookmarkWithCMAttachmentNormal(app, con, "ValuationReport", AppConstants.TableName.CM_ValuationReport, cmId);
					this.FillBookmarkWithCMAttachmentNormal(app, con, "OtherAttachment", AppConstants.TableName.CM_OtherAttachment, cmId);
					#endregion

					IIFCommon.finalizeDoc(doc);
					
					//doc.SaveAs2(Path.Combine(temporaryFolderLocation, fileNamePDF), WdExportFormat.wdExportFormatPDF);
					doc.SaveAs2(Path.Combine(temporaryFolderLocation, fileName));
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
