using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;

using Microsoft.Office.Interop.Word;

using IIF.PAM.MergeDocumentServices.Helper;
using IIF.PAM.MergeDocumentServices.Models;

namespace IIF.PAM.MergeDocumentServices.Services
{
    public class CM_EquityFinance : BaseServices
    {
        public FileMergeResult MergeCMEquityFinance(SqlConnection con, long cmId, string foldertemplate, string temporaryFolderLocation)
        {
            DBHelper db = new DBHelper();

            List<CMData> dataResult = db.ExecToModel<CMData>(con, "dbo.Generate_Document_CM_Data_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, cmId) });

            System.Data.DataTable listBorrower = db.ExecToDataTable(con, "Generate_Document_CM_Borrower_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, cmId) });

            System.Data.DataTable listFacility = db.ExecToDataTable(con, "Generate_Document_CM_ProposalFacility_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, cmId) });

            System.Data.DataTable listDealTeam = db.ExecToDataTable(con, "Generate_Document_CM_DealTeam_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, cmId) });


			string fileName = "CM-" + dataResult[0].ProductType + "-" + dataResult[0].CompanyName + "-" + dataResult[0].ProjectCode + ".docx";
			string fileNamePDF = "CM-" + dataResult[0].ProductType + "-" + dataResult[0].CompanyName + "-" + dataResult[0].ProjectCode + ".pdf";
            string fileTemplateName = "CM Template - Equity.docx";
            string fileTemplateFullName = foldertemplate.AppendPath("\\", fileTemplateName);

            string getfileName = Path.GetFileName(fileTemplateFullName);
            string destFile = Path.Combine(temporaryFolderLocation, fileName);
            File.Copy(fileTemplateFullName, destFile, true);

            object missing = System.Reflection.Missing.Value;
            object readOnly = (object)false;
            Application app = new Application();

            try
            {
                Document doc = app.Documents.Open(destFile, ref missing, ref readOnly);
                app.Visible = false;
                try
                {
                    #region Cover
                    Range reviewMemo = app.ActiveDocument.Bookmarks["bmReviewMemo"].Range;
                    reviewMemo.Text = dataResult[0].ReviewMemo;
                    Range projCompanyName = app.ActiveDocument.Bookmarks["bmInvesteeCompanyName"].Range;
                    projCompanyName.Text = dataResult[0].CompanyName;
                    projCompanyName.Font.Name = "Roboto Light";
                    Range cmNumber = app.ActiveDocument.Bookmarks["bmCMNumber"].Range;
                    //cmNumber.Text = dataResult[0].CM;
                    cmNumber.Font.Name = "Roboto Light";

                    Range projName = app.ActiveDocument.Bookmarks["bmProjectName"].Range;
                    projName.Text = dataResult[0].ProjectName;
                    projName.Font.Name = "Roboto Light";
                    Range projCode = app.ActiveDocument.Bookmarks["bmProjectCode"].Range;
                    projCode.Text = dataResult[0].ProjectCode;
                    projCode.Font.Name = "Roboto Light";
                    Range projDate = app.ActiveDocument.Bookmarks["bmProjectDate"].Range;
                    projDate.Text = dataResult[0].CMDate.ToString("dd-MMMM-yyyy");
                    projDate.Font.Name = "Roboto Light";
                    #endregion

                    #region PROJECT
                    Range project = app.ActiveDocument.Bookmarks["bmProject"].Range;

                    Table tblproject = app.ActiveDocument.Tables.Add(project, 4, 2, WdDefaultTableBehavior.wdWord9TableBehavior);
                    tblproject.Range.Font.Name = "Roboto Light";
                    tblproject.Range.Font.Size = 10;
                    tblproject.set_Style("Table Grid");
                    tblproject.Columns[1].SetWidth(120, WdRulerStyle.wdAdjustFirstColumn);
                    tblproject.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                    tblproject.Cell(1, 1).Range.Text = "Project Name";
					//tblproject.Cell(1, 2).Range.Text = dataResult[0].ProjectName;
					tblproject.Cell(1, 2).Range.Text = "-";
					tblproject.Cell(2, 1).Range.Text = "Sector – Sub sector";
                    tblproject.Cell(2, 2).Range.Text = dataResult[0].SubSector + " " + dataResult[0].SubSector;
                    tblproject.Cell(3, 1).Range.Text = "Funding Needs";
                    tblproject.Cell(3, 2).Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].FundingNeeds));
                    tblproject.Cell(4, 1).Range.Text = "Deal Strategy";
                    tblproject.Cell(4, 2).Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].DealStrategy));
                    #endregion

                    #region BORROWER
                    Range borrower = app.ActiveDocument.Bookmarks["bmInvesteeCompany"].Range;

                    Table tblborrower = app.ActiveDocument.Tables.Add(borrower, 3, 4, WdDefaultTableBehavior.wdWord9TableBehavior);
                    tblborrower.Range.Font.Name = "Roboto Light";
                    tblborrower.Range.Font.Size = 10;
                    tblborrower.set_Style("Table Grid");
                    tblproject.Columns[1].SetWidth(90, WdRulerStyle.wdAdjustFirstColumn);
                    tblborrower.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                    tblborrower.Cell(1, 1).Range.Text = "Investee  Company";
                    tblborrower.Cell(1, 2).Merge(tblborrower.Cell(1, 4));

                    tblborrower.Cell(2, 1).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Cell(2, 1).Range.Text = "Shareholders";
                    tblborrower.Cell(2, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblborrower.Cell(2, 2).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblborrower.Cell(2, 2).Range.Text = "Investee";
                    tblborrower.Cell(2, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblborrower.Cell(2, 3).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblborrower.Cell(2, 3).Range.Text = "Name";
                    tblborrower.Cell(2, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblborrower.Cell(2, 4).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblborrower.Cell(2, 4).Range.Text = "% ownership";

                    string prevkey = "";
                    string cellText = "";
                    int rowCount = 3;
                    foreach (DataRow item in listBorrower.Rows)
                    {
                        prevkey = item[0].ToString() + "\r\a";
                        if (cellText != prevkey)
                        {
                            tblborrower.Rows[rowCount].Cells[2].Range.Text = item[0].ToString();
                            cellText = tblborrower.Rows[rowCount].Cells[2].Range.Text;
                            tblborrower.Cell(1, 2).Range.Text = item[0].ToString() + " ";
                        }

                        tblborrower.Rows.Add(ref missing);
                        tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                        tblborrower.Rows[rowCount].Cells[3].Merge(tblborrower.Rows[rowCount].Cells[4]);
                        tblborrower.Rows[rowCount].Cells[3].Range.Text = item[1].ToString();
                        tblborrower.Rows[rowCount].Cells[4].Range.Text = item[2].ToString();
                        rowCount++;
                    }

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[1].Range.Text = "Ultimate Beneficial Owner";
                    tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                    tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    tblborrower.Rows[rowCount].Cells[2].Range.Text = "";
                    int rowtemp = rowCount;

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[1].Range.Text = "Rating";
                    tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                    tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    tblborrower.Rows[rowCount].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblborrower.Rows[rowCount].Cells[2].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblborrower.Rows[rowCount].Cells[2].Range.Text = "External Rating";
                    tblborrower.Rows[rowCount].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblborrower.Rows[rowCount].Cells[3].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblborrower.Rows[rowCount].Cells[3].Range.Text = "S&E Category";
                    tblborrower.Rows[rowCount].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblborrower.Rows[rowCount].Cells[4].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblborrower.Rows[rowCount].Cells[4].Range.Text = "LQC/BI Checking";

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblborrower.Rows[rowCount].Cells[2].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                    tblborrower.Rows[rowCount].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblborrower.Rows[rowCount].Cells[3].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                    tblborrower.Rows[rowCount].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblborrower.Rows[rowCount].Cells[4].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;

                    tblborrower.Rows[rowCount].Cells[2].Range.Text = "S&P: ";
                    tblborrower.Rows[rowCount].Cells[3].Range.Text = "";
                    tblborrower.Rows[rowCount].Cells[4].Range.Text = "";
                    tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[2].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[3].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[4].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[2].Range.Text = "Moodys: " + dataResult[0].MoodysRate;
                    tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[2].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[3].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[4].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[2].Range.Text = "Fitch: " + dataResult[0].FitchRate;
                    tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[2].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[3].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[4].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[2].Range.Text = "Pefindo: " + dataResult[0].PefindoRate;

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[1].Range.Text = "Business Activities";
                    tblborrower.Rows[rowCount].Cells[2].Merge(tblborrower.Rows[rowCount].Cells[4]);
                    tblborrower.Rows[rowCount].Cells[2].Range.Text = dataResult[0].BusinessActivities;
                    tblborrower.Rows[rowCount].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[1].Range.Text = "Other information";
                    tblborrower.Rows[rowCount].Cells[2].Range.Text = dataResult[0].OtherInformation;
                    tblborrower.Rows[rowCount].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;

                    tblborrower.Rows[rowtemp].Cells[2].Merge(tblborrower.Rows[rowtemp].Cells[4]);
                    #endregion

                    #region PROPOSAL
                    Range proposal = app.ActiveDocument.Bookmarks["bmProposal"].Range;
                    Table tblProposal = app.ActiveDocument.Tables.Add(proposal, 3, 5, WdDefaultTableBehavior.wdWord9TableBehavior);
                    tblProposal.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                    tblProposal.Range.Font.Name = "Roboto Light";
                    tblProposal.Range.Font.Size = 10;
                    tblProposal.set_Style("Table Grid");
                    tblProposal.Columns[1].SetWidth(app.InchesToPoints(1.8f / 2.54f), WdRulerStyle.wdAdjustNone);
                    int rowcount = 2;

                    tblProposal.Cell(1, 1).Range.Text = "Purpose";
                    tblProposal.Cell(1, 2).Merge(tblProposal.Cell(1, 3));
                    tblProposal.Cell(1, 2).Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].Purpose));

                    tblProposal.Cell(2, 1).Range.Text = "Approval Authority";
                    tblProposal.Cell(2, 2).Merge(tblProposal.Cell(2, 3));
                    tblProposal.Cell(2, 2).Range.Text = dataResult[0].ApprovalAuhority;

                    tblProposal.Cell(3, 1).Range.Text = "Investment";
                    tblProposal.Cell(3, 2).Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblProposal.Cell(3, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblProposal.Cell(3, 2).Range.Text = "Type";
                    tblProposal.Cell(3, 3).Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblProposal.Cell(3, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblProposal.Cell(3, 3).Range.Text = "Approved";
                    tblProposal.Cell(3, 4).Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblProposal.Cell(3, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblProposal.Cell(3, 4).Range.Text = "Proposed";
                    tblProposal.Cell(3, 5).Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblProposal.Cell(3, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblProposal.Cell(3, 5).Range.Text = "MTM";

                    foreach (DataRow item in listFacility.Rows)
                    {
                        tblProposal.Rows.Add(ref missing);
                        rowcount++;
                        tblProposal.Rows[rowcount].Cells[2].Range.Text = item[0].ToString();
                        tblProposal.Rows[rowcount].Cells[3].Range.Text = item[1].ToString() + " " + item[2].ToString();
                        tblProposal.Rows[rowcount].Cells[4].Range.Text = item[3].ToString() + " " + item[4].ToString();
                        tblProposal.Rows[rowcount].Cells[5].Range.Text = item[5].ToString() + " " + item[6].ToString();
                    }

                    rowcount++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcount].Cells[1].Range.Text = "Group Exposure";
                    tblProposal.Rows[rowcount].Cells[2].Range.Text = dataResult[0].GroupExposureCurr + dataResult[0].GroupExposureAmount;

                    rowcount++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcount].Cells[1].Range.Text = "Remarks";
                    tblProposal.Rows[rowcount].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].Remarks));

                    rowcount++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcount].Cells[1].Range.Text = "Expected Holding Period";
                    tblProposal.Rows[rowcount].Cells[2].Range.Text = dataResult[0].TenorYear + " year(s)  " + dataResult[0].TenorMonth + " month(s)";

                    rowcount++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcount].Cells[1].Range.Text = "Expected Return";
                    tblProposal.Rows[rowcount].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].PricingOtherConditions));

                    rowcount++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcount].Cells[1].Range.Text = "Other Conditions";
                    tblProposal.Rows[rowcount].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].PricingOtherConditions));

                    rowcount++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcount].Cells[1].Range.Text = "Limit Compliance";
                    tblProposal.Cell(rowcount, 2).Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblProposal.Cell(rowcount, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblProposal.Rows[rowcount].Cells[2].Range.Text = "IDR  million, as of xx1] ";
                    tblProposal.Cell(rowcount, 3).Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblProposal.Cell(rowcount, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblProposal.Rows[rowcount].Cells[3].Range.Text = "Max Limit";
                    tblProposal.Cell(rowcount, 4).Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblProposal.Cell(rowcount, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblProposal.Rows[rowcount].Cells[4].Range.Text = "Proposed";
                    tblProposal.Cell(rowcount, 5).Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblProposal.Cell(rowcount, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblProposal.Rows[rowcount].Cells[5].Range.Text = "Remarks";

                    rowcount++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Columns[rowcount].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                    tblProposal.Rows[rowcount].Cells[2].Range.Text = "Single Project Exposure limit";
                    tblProposal.Cell(rowcount, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblProposal.Rows[rowcount].Cells[3].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceSingleProjectExposureMaxLimit);
                    tblProposal.Cell(rowcount, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblProposal.Rows[rowcount].Cells[4].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceSingleProjectExposureProposed);
                    tblProposal.Cell(rowcount, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblProposal.Rows[rowcount].Cells[5].Range.Text = dataResult[0].SingleProjectExposureRemarks;
                    tblProposal.Cell(rowcount, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                    rowcount++;
                    tblProposal.Rows[rowcount].Cells[2].Range.Text = "Product";
                    tblProposal.Cell(rowcount, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblProposal.Rows[rowcount].Cells[3].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceProductMaxLimit);
                    tblProposal.Cell(rowcount, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblProposal.Rows[rowcount].Cells[4].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceProductProposed);
                    tblProposal.Cell(rowcount, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblProposal.Rows[rowcount].Cells[5].Range.Text = dataResult[0].ProductRemarks;
                    tblProposal.Cell(rowcount, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                    rowcount++;
                    tblProposal.Rows[rowcount].Cells[2].Range.Text = "Group Exposure Limit";
                    tblProposal.Cell(rowcount, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblProposal.Rows[rowcount].Cells[3].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceGrupExposureMaxLimit);
                    tblProposal.Cell(rowcount, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblProposal.Rows[rowcount].Cells[4].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceGrupExposureProposed);
                    tblProposal.Cell(rowcount, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblProposal.Rows[rowcount].Cells[5].Range.Text = dataResult[0].GrupExposureRemarks;
                    tblProposal.Cell(rowcount, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                    rowcount++;
                    tblProposal.Rows[rowcount].Cells[2].Range.Text = "Sector exposure";
                    tblProposal.Cell(rowcount, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblProposal.Rows[rowcount].Cells[3].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceSectorExposureMaxLimit);
                    tblProposal.Cell(rowcount, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblProposal.Rows[rowcount].Cells[4].Range.Text = Convert.ToString(dataResult[0].FacilityLimitComplianceSectorExposureProposed);
                    tblProposal.Cell(rowcount, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblProposal.Rows[rowcount].Cells[5].Range.Text = dataResult[0].SectorExposureRemarks;
                    tblProposal.Cell(rowcount, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                    rowcount++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcount].Cells[1].Range.Text = "Exception to IIF Policy";
                    tblProposal.Rows[rowcount].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].PricingExceptionToIIFPolicy));

                    rowcount++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcount].Cells[1].Range.Text = "Review Period";
                    tblProposal.Rows[rowcount].Cells[2].Range.Text = dataResult[0].ProposalReviewPeriod;
                    #endregion

                    #region D.Recommendation
                    Range keyInvestment = app.ActiveDocument.Bookmarks["bmRecommendation"].Range;
                    Paragraph paragraph = doc.Content.Paragraphs.Add(keyInvestment);
                    paragraph.Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].KeyInvestmentRecommendation));

                    Range recommendation = app.ActiveDocument.Bookmarks["bmRecommendation"].Range;
                    paragraph.Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].Recommendation));

                    Range accountResponsible = app.ActiveDocument.Bookmarks["bmRecommendation"].Range;
                    Table tblAccountResponsible = app.ActiveDocument.Tables.Add(accountResponsible, 1, 3, WdDefaultTableBehavior.wdWord9TableBehavior);
                    tblAccountResponsible.Range.Font.Name = "Roboto Light";
                    tblAccountResponsible.Range.Font.Size = 10;
                    tblAccountResponsible.set_Style("Table Grid");
                    tblAccountResponsible.Columns[1].SetWidth(130, WdRulerStyle.wdAdjustFirstColumn);
                    tblAccountResponsible.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                    tblAccountResponsible.Cell(1, 1).Range.Text = "Account Responsible";
                    tblAccountResponsible.Cell(1, 2).Range.Text = "Deal Team";
                    tblAccountResponsible.Cell(1, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblAccountResponsible.Cell(1, 2).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    tblAccountResponsible.Cell(1, 3).Range.Text = "CIO";
                    tblAccountResponsible.Cell(1, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblAccountResponsible.Cell(1, 3).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray10;

                    int rowCountDealTeam = 1;
                    foreach (DataRow item in listDealTeam.Rows)
                    {
                        rowCountDealTeam++;
                        tblAccountResponsible.Rows.Add(ref missing);

                        tblAccountResponsible.Rows[rowCountDealTeam].Cells[2].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                        tblAccountResponsible.Rows[rowCountDealTeam].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        tblAccountResponsible.Rows[rowCountDealTeam].Cells[3].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                        tblAccountResponsible.Rows[rowCountDealTeam].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        tblAccountResponsible.Rows[rowCountDealTeam].Cells[2].Range.Text = item[0].ToString();
                        if (rowCountDealTeam <= 2)
                        {
                            tblAccountResponsible.Rows[rowCountDealTeam].Cells[3].Range.Text = dataResult[0].AccountResponsibleCIOName;
                        }
                    }
                    #endregion

                    #region Attachment 
                    this.FillBookmarkWithCMAttachmentType1(app, con, "bmPeriodicReview", AppConstants.TableName.CM_PeriodicReview, cmId);
                    this.FillBookmarkWithCMAttachmentType1(app, con, "bmPreviousApproval", AppConstants.TableName.CM_PreviousApprovals, cmId);
                    this.FillBookmarkWithCMAttachmentType1(app, con, "bmRiskRating", AppConstants.TableName.CM_RiskRating, cmId);
                    this.FillBookmarkWithCMAttachmentType1(app, con, "bmKYCchecklist", AppConstants.TableName.CM_KYCChecklists, cmId);
                    this.FillBookmarkWithCMAttachmentType1(app, con, "bmSAndEReview", AppConstants.TableName.CM_SAndEReview, cmId);
                    this.FillBookmarkWithCMAttachmentType1(app, con, "bmOtherbankFacilities", AppConstants.TableName.CM_OtherBanksFacilities, cmId);
                    this.FillBookmarkWithCMAttachmentType1(app, con, "bmValutationReport", AppConstants.TableName.CM_ValuationReport, cmId);
                    this.FillBookmarkWithCMAttachmentType1(app, con, "bmOtherAttachment", AppConstants.TableName.CM_OtherAttachment, cmId);
					#endregion

					doc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
					doc.SaveAs2(Path.Combine(temporaryFolderLocation, fileNamePDF), WdExportFormat.wdExportFormatPDF);
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
