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
    public class CM_WaiverFinanceNEW : BaseServices
    {
        public FileMergeResult MergeCMWaiverFinance(SqlConnection con, long cmId, string foldertemplate, string temporaryFolderLocation)
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

            string fileTemplateName = "CM Waiver Template.docx";
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

					#region DOCUMENT2.DOCX
					//Application app2 = new Application();
					//Document myWordCollection1 = app2.Documents.Open(@"\\k2projectiif\c$\IIF\PAMTemplate\Document2.docx");
					//String myTes1 = myWordCollection1.Content.get_XML();
					//myWordCollection1.Close();					
					#endregion

					#region PROJECT
					app.ActiveDocument.Bookmarks["AxPROJECTxProjectName"].Range.Text = dataResult[0].ProjectName;
					app.ActiveDocument.Bookmarks["AxPROJECTxSectorSubsector"].Range.Text = dataResult[0].SectorDesc + " " + dataResult[0].SubSectorDesc;
					app.ActiveDocument.Bookmarks["AxPROJECTxFundingNeeds"].Range.InsertFile(ConvertHtmlAndFile.SaveToHtmlNew(dataResult[0].FundingNeeds, currFontFamily, currFontSize));
					//app.ActiveDocument.Bookmarks["AxPROJECTxFundingNeeds"].Range.InsertXML(myTes1);
					#endregion

					#region BORROWER
					app.ActiveDocument.Bookmarks["BxBORROWERxProjectCompany"].Range.Text = dataResult[0].CompanyName;

					Table tblShareholders = IIFCommon.createTable(app, "BxBORROWERxProjectSponsors", 3, true);
					//header
					tblShareholders.Cell(1, 1).Range.Text = "Project Company/ Investee/ Borrower";
					tblShareholders.Cell(1, 2).Range.Text = "Project Sponsors/Shareholders";
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
						if (cellText != prevkey)
						{
							//merge kolom kalo value nya beda, mulai row ke 3
							if (rowCounter > 2 && (rowTemp != (rowCounter - 1)))
							{
								tblShareholders.Cell(rowTemp, 1).Merge(tblShareholders.Cell(rowCounter - 1, 1));
							}

							rowTemp = rowCounter;

							tblShareholders.Cell(rowCounter, 1).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
							tblShareholders.Cell(rowCounter, 1).Range.Text = item[0].ToString();
							cellText = item[0].ToString();
							tblShareholders.Cell(rowCounter, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
						}

						tblShareholders.Cell(rowCounter, 2).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
						tblShareholders.Cell(rowCounter, 2).Range.Text = item[1].ToString();
						tblShareholders.Cell(rowCounter, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

						tblShareholders.Cell(rowCounter, 3).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
						tblShareholders.Cell(rowCounter, 3).Range.Text = item[2].ToString();
						tblShareholders.Cell(rowCounter, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					}

					//merge kolom untuk value trakhir
					if (rowCounter > 2 && (rowTemp != rowCounter))
					{
						tblShareholders.Cell(rowTemp, 1).Merge(tblShareholders.Cell(rowCounter, 1));
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
					


					app.ActiveDocument.Bookmarks["BxBORROWERxOtherInformation"].Range.InsertFile(ConvertHtmlAndFile.SaveToHtmlNew(dataResult[0].OtherInformation, currFontFamily, currFontSize));
					#endregion

					#region PROPOSAL
					app.ActiveDocument.Bookmarks["CxPROPOSALxPurpose"].Range.InsertFile(ConvertHtmlAndFile.SaveToHtmlNew(dataResult[0].Purpose, currFontFamily, currFontSize));
					app.ActiveDocument.Bookmarks["CxPROPOSALxApprovalAuthority"].Range.Text = dataResult[0].ApprovalAuhority;

					Table tblFacility = IIFCommon.createTable(app, "CxPROPOSALxFacility", 4, true);
					//header
					tblFacility.Cell(1, 1).Range.Text = "Type";
					tblFacility.Cell(1, 2).Range.Text = "Approved";
					tblFacility.Cell(1, 3).Range.Text = "Proposed";
					tblFacility.Cell(1, 4).Range.Text = "Outstanding";
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
						tblFacility.Cell(rowCounter, 3).Range.Text = "";

						tblFacility.Cell(rowCounter, 4).Shading.BackgroundPatternColor = WdColor.wdColorWhite;
						tblFacility.Cell(rowCounter, 4).Range.Text = "";
					}

					tblFacility.Rows.Add(ref missing);
					tblFacility.Cell(rowCounter + 1, 1).Range.Text = "Remarks : " + dataResult[0].FacilityOrInvestmentRemarks;
					tblFacility.Cell(rowCounter + 1, 1).Merge(tblFacility.Cell(rowCounter + 1, 4));
					tblFacility.Cell(rowCounter + 1, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

					app.ActiveDocument.Bookmarks["CxPROPOSALxGroupExposure"].Range.Text = dataResult[0].GroupExposureCurr + " " + dataResult[0].GroupExposureAmount;
					app.ActiveDocument.Bookmarks["CxPROPOSALxRemarks"].Range.InsertFile(ConvertHtmlAndFile.SaveToHtmlNew(dataResult[0].Remarks, currFontFamily, currFontSize));
					app.ActiveDocument.Bookmarks["CxPROPOSALxTenor"].Range.Text = dataResult[0].TenorYear + " year(s)  " + dataResult[0].TenorMonth + " month(s)";
					app.ActiveDocument.Bookmarks["CxPROPOSALxAverageLoanLife"].Range.Text = dataResult[0].AverageLoanLifeYear + " year(s)  " + dataResult[0].AverageLoanLifeMonth + " month(s)";

					app.ActiveDocument.Bookmarks["CxPROPOSALxPricingxInterestRate"].Range.InsertFile(ConvertHtmlAndFile.SaveToHtmlNew(dataResult[0].PricingInterestRate, currFontFamily, currFontSize));
					app.ActiveDocument.Bookmarks["CxPROPOSALxPricingxCommitmentFee"].Range.Text = dataResult[0].PricingCommitmentFee;
					app.ActiveDocument.Bookmarks["CxPROPOSALxPricingxFacility"].Range.Text = dataResult[0].PricingUpfrontFacilityFee;
					app.ActiveDocument.Bookmarks["CxPROPOSALxPricingxStructuringFee"].Range.Text = dataResult[0].PricingStructuringFee;
					app.ActiveDocument.Bookmarks["CxPROPOSALxPricingxArrangerFee"].Range.Text = dataResult[0].PricingArrangerFee;

					app.ActiveDocument.Bookmarks["CxPROPOSALxCollateral"].Range.InsertFile(ConvertHtmlAndFile.SaveToHtmlNew(dataResult[0].PricingCollateral, currFontFamily, currFontSize));
					app.ActiveDocument.Bookmarks["CxPROPOSALxOtherCondition"].Range.InsertFile(ConvertHtmlAndFile.SaveToHtmlNew(dataResult[0].PricingOtherConditions, currFontFamily, currFontSize));

					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSPELxML"].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceSingleProjectExposureMaxLimit);
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSPELxP"].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceSingleProjectExposureProposed);
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSPELxR"].Range.Text = dataResult[0].SingleProjectExposureRemarks;

					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexCurrency"].Range.Text = dataResult[0].LimitComplianceCurrency;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexAsFor"].Range.Text = Convert.ToDateTime("01-"+dataResult[0].FacilityLimitComplianceMonth.ToString()+"-1985").ToString("MMM") + " " + dataResult[0].FacilityLimitComplianceYear.ToString();
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexRiskRating"].Range.Text = dataResult[0].FacilityLimitComplianceIIFRate;
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSecExposure"].Range.Text = dataResult[0].SectorDesc;					

					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexPxML"].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceProductMaxLimit);
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexPxP"].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceProductProposed);
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexPxR"].Range.Text = dataResult[0].ProductRemarks;

					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexRRxML"].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceRiskRatingMaxLimit);
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexRRxP"].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceRiskRatingProposed);
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexRRxR"].Range.Text = dataResult[0].RiskRatingRemarks;

					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexGELxML"].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceGrupExposureMaxLimit);
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexGELxP"].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceGrupExposureProposed);
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexGELxR"].Range.Text = dataResult[0].GrupExposureRemarks;

					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSExML"].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceSectorExposureMaxLimit);
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSExP"].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceSectorExposureProposed);
					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexSExR"].Range.Text = dataResult[0].SectorExposureRemarks;

					app.ActiveDocument.Bookmarks["CxPROPOSALxLimitCompliancexNotes"].Range.Text = dataResult[0].notes;

					app.ActiveDocument.Bookmarks["CxPROPOSALxExceptionToIIFPolicy"].Range.InsertFile(ConvertHtmlAndFile.SaveToHtmlNew(dataResult[0].PricingExceptionToIIFPolicy, currFontFamily, currFontSize));
					app.ActiveDocument.Bookmarks["CxPROPOSALxReviewPeriod"].Range.Text = dataResult[0].ProposalReviewPeriod;
					#endregion

					#region RECOMMENDATION
					app.ActiveDocument.Bookmarks["DxRECOMMENDATIONxKeyInvestment"].Range.InsertFile(ConvertHtmlAndFile.SaveToHtmlNew(dataResult[0].KeyInvestmentRecommendation, currFontFamily, currFontSize));
					app.ActiveDocument.Bookmarks["DxRECOMMENDATION"].Range.InsertFile(ConvertHtmlAndFile.SaveToHtmlNew(dataResult[0].Recommendation, currFontFamily, currFontSize));

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
					this.FillBookmarkWithCMAttachmentNormal(app, con, "OtherAttachment", AppConstants.TableName.CM_OtherAttachment, cmId);
					#endregion

					IIFCommon.finalizeDoc(doc);

					//doc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
					doc.SaveAs2(Path.Combine(temporaryFolderLocation, fileNamePDF), WdExportFormat.wdExportFormatPDF);
					//doc.SaveAs2(Path.Combine(temporaryFolderLocation, fileName));
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
