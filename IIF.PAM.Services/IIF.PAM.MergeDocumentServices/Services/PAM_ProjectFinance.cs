﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;

using Microsoft.Office.Interop.Word;

using IIF.PAM.MergeDocumentServices.Helper;
using IIF.PAM.MergeDocumentServices.Models;

namespace IIF.PAM.MergeDocumentServices.Services
{
    public class PAM_ProjectFinance : BaseServices
    {
        public FileMergeResult MergePAMProjectFinance(SqlConnection con, long pamId, string foldertemplate, string temporaryFolderLocation)
        {
            DBHelper db = new DBHelper();

            List<PAMData> dataResult = db.ExecToModel<PAMData>(con, "dbo.Generate_Document_PAM_Data_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

            System.Data.DataTable listBorrower = db.ExecToDataTable(con, "Generate_Document_PAM_Borrower_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

            System.Data.DataTable listFacility = db.ExecToDataTable(con, "Generate_Document_PAM_ProposalFacility_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

            System.Data.DataTable listDealTeam = db.ExecToDataTable(con, "Generate_Document_PAM_DealTeam_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

            string fileName = "PAM-" + dataResult[0].ProductType + "-" + dataResult[0].ProjectCompanyName + "-" + dataResult[0].ProjectCode + ".docx";
            string fileNamePDF = "PAM-" + dataResult[0].ProductType + "-" + dataResult[0].ProjectCompanyName + "-" + dataResult[0].ProjectCode + ".pdf";
            string fileTemplateName = "PAM Template - Project Finance.docx";
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
                    #region header & Footer
                    //doc.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
                    //Object oMissing = System.Reflection.Missing.Value;
                    //doc.ActiveWindow.Selection.TypeText("Review Memorandum – Project Finance");
                    //Object TotalPages = WdFieldType.wdFieldNumPages;
                    //Object CurrentPage = WdFieldType.wdFieldPage;
                    //doc.ActiveWindow.Selection.TypeText("\t\t");
                    //doc.ActiveWindow.Selection.Fields.Add(doc.ActiveWindow.Selection.Range, ref CurrentPage, ref oMissing, ref oMissing);
                    #endregion

                    #region Cover
                    this.SetBookmarkText(app, "ProjectCompanyName", dataResult[0].ProjectCompanyName);
                    this.SetBookmarkText(app, "ProjectName", dataResult[0].ProjectName);
                    this.SetBookmarkText(app, "ProjectCode", dataResult[0].ProjectCode);
                    this.SetBookmarkText(app, "ProjectDate", dataResult[0].PAMDate.ToString("dd-MMMM-yyyy"));
                    #endregion

                    #region A.Project
                    Range project = app.ActiveDocument.Bookmarks["Project"].Range;
                    Table tblproject = app.ActiveDocument.Tables.Add(project, 6, 2, WdDefaultTableBehavior.wdWord9TableBehavior);
                    tblproject.Range.Font.Name = "Roboto Light";
                    tblproject.Range.Font.Size = 10;
                    tblproject.set_Style("Table Grid");

                    tblproject.Columns[1].SetWidth(120, WdRulerStyle.wdAdjustFirstColumn);
                    tblproject.Cell(1, 1).Range.Text = "Project Description";
                    tblproject.Cell(1, 2).Range.Text = dataResult[0].ProjectDescription.ToString();
                    tblproject.Cell(2, 1).Range.Text = "Sector – Sub sector";
                    tblproject.Cell(2, 2).Range.Text = dataResult[0].SubSector + dataResult[0].SubSector;
                    tblproject.Cell(3, 1).Range.Text = "Project Cost";
                    tblproject.Cell(3, 2).Range.Text = dataResult[0].ProjectCostCurr + " " + dataResult[0].ProjectCostAmount;
                    tblproject.Cell(4, 1).Range.Text = "Project Scope";
                    tblproject.Cell(4, 2).Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].ProjectScope));
                    tblproject.Cell(4, 1).Range.Font.Name = "Roboto Light";
                    tblproject.Cell(4, 1).Range.Font.Size = 10;
                    tblproject.Cell(5, 1).Range.Text = "Project Structure";
                    tblproject.Cell(5, 2).Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].ProjectStructure));
                    tblproject.Cell(5, 1).Range.Font.Name = "Roboto Light";
                    tblproject.Cell(5, 1).Range.Font.Size = 10;
                    tblproject.Cell(6, 1).Range.Text = "Deal Strategy";
                    tblproject.Cell(6, 2).Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].DealStrategy));
                    tblproject.Cell(6, 1).Range.Font.Name = "Roboto Light";
                    tblproject.Cell(6, 1).Range.Font.Size = 10;
                    #endregion

                    #region B.Borrower
                    Range borrower = app.ActiveDocument.Bookmarks["Borrower"].Range;
                    Table tblborrower = app.ActiveDocument.Tables.Add(borrower, 2, 5, WdDefaultTableBehavior.wdWord9TableBehavior);
                    tblborrower.Range.Font.Name = "Roboto Light";
                    tblborrower.Range.Font.Size = 10;
                    tblborrower.set_Style("Table Grid");
                    tblborrower.Columns[1].SetWidth(100, WdRulerStyle.wdAdjustFirstColumn);

                    tblborrower.Cell(1, 1).Range.Text = "Project Company";

                    string prevkey = "";
                    string cellText = "";
                    foreach (DataRow items in listBorrower.Rows)
                    {
                        prevkey = items[0].ToString() + "\r\a";
                        if (cellText != prevkey)
                        {
                            tblborrower.Cell(1, 2).Range.Text = items[0].ToString();
                            cellText = tblborrower.Cell(1, 2).Range.Text;
                        }
                    }

                    tblborrower.Cell(2, 1).Range.Text = "Project Sponsors";
                    tblborrower.Cell(2, 1).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;

                    tblborrower.Cell(2, 2).SetWidth(140, WdRulerStyle.wdAdjustFirstColumn);
                    tblborrower.Cell(2, 2).Range.Text = "Project Company";
                    tblborrower.Cell(2, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblborrower.Cell(2, 2).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;

                    tblborrower.Cell(2, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblborrower.Cell(2, 3).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
                    tblborrower.Cell(2, 3).Range.Text = "Name Project Sponsors";

                    tblborrower.Cell(2, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblborrower.Cell(2, 5).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
                    tblborrower.Cell(2, 5).Range.Text = "% ownership";

                    string prevkey1 = "";
                    string cellText1 = "";
                    int rowCount = 2;
                    foreach (DataRow item in listBorrower.Rows)
                    {
                        rowCount++;
                        tblborrower.Rows.Add(ref missing);

                        prevkey1 = item[0].ToString() + "\r\a";
                        if (cellText1 != prevkey1)
                        {
                            tblborrower.Rows[rowCount].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                            tblborrower.Rows[rowCount].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                            tblborrower.Rows[rowCount].Cells[5].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                            tblborrower.Rows[rowCount].Cells[2].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                            tblborrower.Rows[rowCount].Cells[3].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                            tblborrower.Rows[rowCount].Cells[5].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;

                            tblborrower.Rows[rowCount].Cells[2].Range.Text = item[0].ToString();
                            cellText1 = tblborrower.Rows[rowCount].Cells[2].Range.Text;
                            tblborrower.Cell(1, 2).Range.Text = item[0].ToString() + " ";
                            tblborrower.Cell(1, 2).Range.Font.Name = "Roboto Light";
                            tblborrower.Cell(1, 2).Range.Font.Size = 10;
                        }
                        tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
                        tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                        tblborrower.Rows[rowCount].Cells[3].Range.Text = item[1].ToString();
                        var ownership = Convert.ToDecimal(item[2]);
                        tblborrower.Rows[rowCount].Cells[5].Range.Text = ownership.ToString("#,#");
                        tblborrower.Rows[rowCount].Cells[5].Range.Font.Name = "Roboto Light";
                        tblborrower.Rows[rowCount].Cells[5].Range.Font.Size = 10;

                    }

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[1].Range.Text = "Ultimate Beneficial Owner";
                    tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                    tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    int rowtemp = rowCount;

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[1].Range.Text = "Rating";
                    tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblborrower.Rows[rowCount].Cells[2].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
                    tblborrower.Rows[rowCount].Cells[2].Range.Text = "IIF Rating";
                    tblborrower.Rows[rowCount].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblborrower.Rows[rowCount].Cells[3].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
                    tblborrower.Rows[rowCount].Cells[3].Range.Text = "External Rating";
                    tblborrower.Rows[rowCount].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblborrower.Rows[rowCount].Cells[4].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
                    tblborrower.Rows[rowCount].Cells[4].Range.Text = "S&E Category";
                    tblborrower.Rows[rowCount].Cells[5].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblborrower.Rows[rowCount].Cells[5].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
                    tblborrower.Rows[rowCount].Cells[5].Range.Text = "LQC/BI Checking";

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblborrower.Rows[rowCount].Cells[2].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                    tblborrower.Rows[rowCount].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblborrower.Rows[rowCount].Cells[3].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                    tblborrower.Rows[rowCount].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblborrower.Rows[rowCount].Cells[4].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                    tblborrower.Rows[rowCount].Cells[5].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblborrower.Rows[rowCount].Cells[5].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;

                    tblborrower.Rows[rowCount].Cells[2].Range.Text = dataResult[0].IIFRate;
                    tblborrower.Rows[rowCount].Cells[3].Range.Text = "S&P: " + dataResult[0].SAndPRate;
                    tblborrower.Rows[rowCount].Cells[4].Range.Text = dataResult[0].SAndECategoryRate;
                    tblborrower.Rows[rowCount].Cells[5].Range.Text = dataResult[0].LQCOrBICheckingRate;

                    tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[2].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[3].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[4].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[5].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;

                    tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[2].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[3].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[4].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblborrower.Rows[rowCount].Cells[5].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[2].Range.Text = "Rating Date: " + dataResult[0].IIFRatingDate.Value.ToString("dd-MMM-yyyy");
                    tblborrower.Rows[rowCount].Cells[3].Range.Text = "Moodys: " + dataResult[0].MoodysRate;

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[3].Range.Text = "Fitch: " + dataResult[0].FitchRate;

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[3].Range.Text = "Pefindo: " + dataResult[0].PefindoRate;

                    rowCount++;
                    tblborrower.Rows.Add(ref missing);
                    tblborrower.Rows[rowCount].Cells[1].Range.Text = "Other information";
                    tblborrower.Rows[rowCount].Cells[2].Merge(tblborrower.Rows[rowCount].Cells[5]);
                    tblborrower.Rows[rowCount].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].OtherInformation));
                    tblborrower.Rows[rowCount].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    tblborrower.Rows[rowCount].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                    tblborrower.Rows[rowCount].Cells[2].Range.Font.Name = "Robaoto Light";
                    tblborrower.Rows[rowCount].Cells[2].Range.Font.Size = 10;

                    tblborrower.Cell(1, 2).Merge(tblborrower.Cell(1, 5));
                    tblborrower.Cell(2, 3).Merge(tblborrower.Cell(2, 4));
                    tblborrower.Cell(2, 3).SetWidth(130, WdRulerStyle.wdAdjustSameWidth);

                    for (int a = 1; a <= listBorrower.Rows.Count; a++)
                    {
                        tblborrower.Cell(2 + a, 3).Merge(tblborrower.Cell(2 + a, 4));
                        tblborrower.Cell(2 + a, 3).SetWidth(130, WdRulerStyle.wdAdjustSameWidth);
                    }
                    tblborrower.Rows[rowtemp].Cells[2].Merge(tblborrower.Rows[rowtemp].Cells[5]);
                    #endregion

                    #region C.Proposal
                    Range proposal = app.ActiveDocument.Bookmarks["Proposal"].Range;
                    Table tblProposal = app.ActiveDocument.Tables.Add(proposal, 3, 3, WdDefaultTableBehavior.wdWord9TableBehavior);
                    tblProposal.Range.Font.Name = "Roboto Light";
                    tblProposal.Range.Font.Size = 10;
                    tblProposal.set_Style("Table Grid");
                    tblProposal.Columns[1].SetWidth(130, WdRulerStyle.wdAdjustFirstColumn);
                    tblProposal.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                    tblProposal.Cell(1, 1).Range.Text = "Purpose";
                    tblProposal.Cell(1, 2).Merge(tblProposal.Cell(1, 3));
                    tblProposal.Cell(1, 2).Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].Purpose));
                    tblProposal.Cell(1, 2).Range.Font.Name = "Roboto Light";

                    tblProposal.Cell(2, 1).Range.Text = "Approval Authority";
                    tblProposal.Cell(2, 2).Merge(tblProposal.Cell(2, 3));
                    tblProposal.Cell(2, 2).Range.Text = dataResult[0].ApprovalAuthority;

                    tblProposal.Cell(3, 1).Range.Text = "Facility";
                    tblProposal.Cell(3, 2).Shading.BackgroundPatternColor = WdColor.wdColorGray40;
                    tblProposal.Cell(3, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblProposal.Cell(3, 2).Range.Text = "Type";
                    tblProposal.Cell(3, 3).Shading.BackgroundPatternColor = WdColor.wdColorGray40;
                    tblProposal.Cell(3, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblProposal.Cell(3, 3).Range.Text = "Amount";

                    int rowcounttblProposal = 3;
                    foreach (DataRow item in listFacility.Rows)
                    {
                        rowcounttblProposal++;
                        tblProposal.Rows.Add(ref missing);
                        tblProposal.Rows[rowcounttblProposal].Cells[2].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                        tblProposal.Rows[rowcounttblProposal].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        tblProposal.Rows[rowcounttblProposal].Cells[3].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                        tblProposal.Rows[rowcounttblProposal].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = item[0].ToString();
                        var amount = Convert.ToDecimal(item[2]);
                        tblProposal.Rows[rowcounttblProposal].Cells[3].Range.Text = item[1].ToString() + " " + amount.ToString("#,#");
                    }

                    rowcounttblProposal++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tblProposal.Rows[rowcounttblProposal].Cells[3].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                    tblProposal.Rows[rowcounttblProposal].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                    tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Group Exposure";
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = dataResult[0].GroupExposureCurr + " " + dataResult[0].GroupExposureAmount;
                    int GroupExpoCount = rowcounttblProposal;

                    rowcounttblProposal++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Remarks";
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].Remarks));
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Font.Name = "Roboto Light";
                    int RemarksCount = rowcounttblProposal;

                    rowcounttblProposal++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Tenor";
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = dataResult[0].tenorYear + " year(s)  " + dataResult[0].tenorMonth + " month(s)";
                    int tenorCount = rowcounttblProposal;

                    rowcounttblProposal++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Average Loan Life";
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = dataResult[0].averageLoanLifeYear + " year(s)  " + dataResult[0].averageLoanLifeMonth + " month(s)";
                    int AverageCount = rowcounttblProposal;

                    rowcounttblProposal++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Pricing";
                    tblProposal.Rows[rowcounttblProposal].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = "Interest rate";
                    tblProposal.Rows[rowcounttblProposal].Cells[3].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].pricingInterestRate));
                    tblProposal.Rows[rowcounttblProposal].Cells[3].Range.Font.Name = "Roboto Light";

                    rowcounttblProposal++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = "Commitment Fee";
                    tblProposal.Rows[rowcounttblProposal].Cells[3].Range.Text = dataResult[0].pricingCommitmentFee;

                    rowcounttblProposal++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = "Upfront Fee";
                    tblProposal.Rows[rowcounttblProposal].Cells[3].Range.Text = dataResult[0].pricingUpfrontFacilityFee;

                    rowcounttblProposal++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = "Structuring Fee";
                    tblProposal.Rows[rowcounttblProposal].Cells[3].Range.Text = dataResult[0].pricingStructuringFee;

                    rowcounttblProposal++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcounttblProposal].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = "Arranger Fee";
                    tblProposal.Rows[rowcounttblProposal].Cells[3].Range.Text = dataResult[0].pricingArrangerFee;

                    rowcounttblProposal++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Collateral";
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].pricingCollateral));
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Font.Name = "Roboto Light";
                    tblProposal.Cell(rowcounttblProposal, 2).Merge(tblProposal.Cell(rowcounttblProposal, 3));

                    rowcounttblProposal++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Other Conditions";
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].pricingOtherConditions));
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Font.Name = "Roboto Light";

                    rowcounttblProposal++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Exception to IIF Policy";
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].pricingExceptionToIIFPolicy));
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Font.Name = "Roboto Light";

                    rowcounttblProposal++;
                    tblProposal.Rows.Add(ref missing);
                    tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Review Period";
                    tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = dataResult[0].reviewPeriod;

                    tblProposal.Cell(GroupExpoCount, 2).Merge(tblProposal.Cell(GroupExpoCount, 3));
                    tblProposal.Cell(RemarksCount, 2).Merge(tblProposal.Cell(RemarksCount, 3));
                    tblProposal.Cell(tenorCount, 2).Merge(tblProposal.Cell(tenorCount, 3));
                    tblProposal.Cell(AverageCount, 2).Merge(tblProposal.Cell(AverageCount, 3));
                    #endregion

                    #region D.Recommendation
                    Range keyInvestment = app.ActiveDocument.Bookmarks["KeyInvestment"].Range;
                    keyInvestment.Font.Name = "Roboto Light";
                    keyInvestment.Font.Size = 10;
                    Paragraph paragkeyInvestment = doc.Content.Paragraphs.Add(keyInvestment);
                    paragkeyInvestment.Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].KeyInvestmentRecommendation));
                    paragkeyInvestment.Range.Font.Name = "Roboto Light";

                    Range recommendation = app.ActiveDocument.Bookmarks["Recommendation"].Range;
                    recommendation.Font.Name = "Roboto Light";
                    recommendation.Font.Size = 10;
                    Paragraph paragrecommendation = doc.Content.Paragraphs.Add(recommendation);
                    paragrecommendation.Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].Recommendation));
                    paragrecommendation.Range.Font.Name = "Roboto Light";
                    paragrecommendation.Range.Font.Size = 10;

                    Range accountResponsible = app.ActiveDocument.Bookmarks["AccountResponsible"].Range;
                    Table tblAccountResponsible = app.ActiveDocument.Tables.Add(accountResponsible, 1, 3, WdDefaultTableBehavior.wdWord9TableBehavior);
                    tblAccountResponsible.Range.Font.Name = "Roboto Light";
                    tblAccountResponsible.Range.Font.Size = 10;
                    tblAccountResponsible.set_Style("Table Grid");
                    tblAccountResponsible.Columns[1].SetWidth(130, WdRulerStyle.wdAdjustFirstColumn);
                    tblAccountResponsible.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                    tblAccountResponsible.Cell(1, 1).Range.Text = "Account Responsible";
                    tblAccountResponsible.Cell(1, 2).Range.Text = "Deal Team";
                    tblAccountResponsible.Cell(1, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblAccountResponsible.Cell(1, 2).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
                    tblAccountResponsible.Cell(1, 3).Range.Text = "CIO";
                    tblAccountResponsible.Cell(1, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tblAccountResponsible.Cell(1, 3).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;

                    int rowCountDealTeam = 1;
                    foreach (DataRow item in listDealTeam.Rows)
                    {
                        rowCountDealTeam++;
                        tblAccountResponsible.Rows.Add(ref missing);
                        tblAccountResponsible.Cell(rowCountDealTeam, 1).Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
                        tblAccountResponsible.Cell(rowCountDealTeam, 1).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
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
                    tblAccountResponsible.Cell(rowCountDealTeam, 1).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                    #endregion

                    #region Attachment
                    this.FillBookmarkWithPAMAttachmentType1(app, con, "ProjectAnalysis", AppConstants.TableName.PAM_ProjectAnalysis, pamId);
                    this.FillBookmarkWithPAMAttachmentType1(app, con, "HistoricalFinancialandFinancialProject", AppConstants.TableName.PAM_HistoricalFinancial, pamId);
                    this.FillBookmarkWithPAMAttachmentType1(app, con, "Supplemental", AppConstants.TableName.PAM_Supplemental, pamId);
                    this.FillBookmarkWithPAMAttachmentType1(app, con, "SocialEnvironmental", AppConstants.TableName.PAM_Social, pamId);
                    this.FillBookmarkWithPAMAttachmentType1(app, con, "GroupStructure", AppConstants.TableName.PAM_GroupStructure, pamId);
                    this.FillBookmarkWithPAMAttachmentType1(app, con, "TermSheet", AppConstants.TableName.PAM_TermSheet, pamId);
                    this.FillBookmarkWithPAMAttachmentType1(app, con, "riskRating", AppConstants.TableName.PAM_RiskRating, pamId);
                    this.FillBookmarkWithPAMAttachmentType1(app, con, "kycChecklists", AppConstants.TableName.PAM_KYCChecklists, pamId);
                    this.FillBookmarkWithPAMAttachmentType1(app, con, "Otherbanks", AppConstants.TableName.PAM_OtherBanksFacilities, pamId);
                    this.FillBookmarkWithPAMAttachmentType1(app, con, "IndustryAnalysis", AppConstants.TableName.PAM_IndustryAnalysis, pamId);

                    this.FillBookmarkWithPAMAttachmentType2(app, con, "LegalDuediligenceReport", AppConstants.TableName.PAM_LegalDueDiligenceReport, pamId);
                    this.FillBookmarkWithPAMAttachmentType2(app, con, "SAndEDueDiligence", AppConstants.TableName.PAM_SAndEDueDiligence, pamId);
                    this.FillBookmarkWithPAMAttachmentType2(app, con, "OtherReport", AppConstants.TableName.PAM_OtherReports, pamId);
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

        //public FileMergeResult MergePAMProjectFinance(SqlConnection con, long pamId, string foldertemplate, string temporaryFolderLocation)
        //{
        //    DBHelper db = new DBHelper();

        //    List<PAMData> dataResult = db.ExecToModel<PAMData>(con, "dbo.Generate_Document_PAM_Data_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

        //    System.Data.DataTable listBorrower = db.ExecToDataTable(con, "Generate_Document_PAM_Borrower_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

        //    System.Data.DataTable listFacility = db.ExecToDataTable(con, "Generate_Document_PAM_ProposalFacility_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

        //    System.Data.DataTable listDealTeam = db.ExecToDataTable(con, "Generate_Document_PAM_DealTeam_SP", CommandType.StoredProcedure, new List<SqlParameter> { this.NewSqlParameter("@Id", SqlDbType.BigInt, pamId) });

        //    string fileName = "PAM-" + dataResult[0].ProductType + "-" + dataResult[0].ProjectCompanyName + "-" + dataResult[0].ProjectCode + ".docx";
        //    string fileNamePDF = "PAM-" + dataResult[0].ProductType + "-" + dataResult[0].ProjectCompanyName + "-" + dataResult[0].ProjectCode + ".pdf";
        //    string fileTemplateName = "PAM Template - Project Finance.docx";
        //    string fileTemplateFullName = foldertemplate.AppendPath("\\", fileTemplateName);

        //    string getfileName = Path.GetFileName(fileTemplateFullName);
        //    string destFile = Path.Combine(temporaryFolderLocation, fileName);
        //    File.Copy(fileTemplateFullName, destFile, true);

        //    object missing = System.Reflection.Missing.Value;
        //    object readOnly = (object)false;
        //    Application app = new Application();

        //    try
        //    {
        //        Document doc = app.Documents.Open(destFile, ref missing, ref readOnly);
        //        app.Visible = false;
        //        try
        //        {
        //            #region header & Footer
        //            //doc.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
        //            //Object oMissing = System.Reflection.Missing.Value;
        //            //doc.ActiveWindow.Selection.TypeText("Review Memorandum – Project Finance");
        //            //Object TotalPages = WdFieldType.wdFieldNumPages;
        //            //Object CurrentPage = WdFieldType.wdFieldPage;
        //            //doc.ActiveWindow.Selection.TypeText("\t\t");
        //            //doc.ActiveWindow.Selection.Fields.Add(doc.ActiveWindow.Selection.Range, ref CurrentPage, ref oMissing, ref oMissing);
        //            #endregion

        //            #region Cover
        //            this.SetBookmarkText(app, "ProjectCompanyName", dataResult[0].ProjectCompanyName);
        //            this.SetBookmarkText(app, "ProjectName", dataResult[0].ProjectName);
        //            this.SetBookmarkText(app, "ProjectCode", dataResult[0].ProjectCode);
        //            this.SetBookmarkText(app, "ProjectDate", dataResult[0].PAMDate.ToString("dd-MMMM-yyyy"));
        //            #endregion

        //            #region A.Project
        //            Range project = app.ActiveDocument.Bookmarks["Project"].Range;
        //            Table tblproject = app.ActiveDocument.Tables.Add(project, 6, 2, WdDefaultTableBehavior.wdWord9TableBehavior);
        //            tblproject.Range.Font.Name = "Roboto Light";
        //            tblproject.Range.Font.Size = 10;
        //            tblproject.set_Style("Table Grid");

        //            tblproject.Columns[1].SetWidth(120, WdRulerStyle.wdAdjustFirstColumn);
        //            tblproject.Cell(1, 1).Range.Text = "Project Description";
        //            tblproject.Cell(1, 2).Range.Text = dataResult[0].ProjectDescription.ToString();
        //            tblproject.Cell(2, 1).Range.Text = "Sector – Sub sector";
        //            tblproject.Cell(2, 2).Range.Text = dataResult[0].SubSector + dataResult[0].SubSector;
        //            tblproject.Cell(3, 1).Range.Text = "Project Cost";
        //            tblproject.Cell(3, 2).Range.Text = dataResult[0].ProjectCostCurr + " " + dataResult[0].ProjectCostAmount;
        //            tblproject.Cell(4, 1).Range.Text = "Project Scope";
        //            tblproject.Cell(4, 2).Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].ProjectScope));
        //            tblproject.Cell(4, 1).Range.Font.Name = "Roboto Light";
        //            tblproject.Cell(4, 1).Range.Font.Size = 10;
        //            tblproject.Cell(5, 1).Range.Text = "Project Structure";
        //            tblproject.Cell(5, 2).Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].ProjectStructure));
        //            tblproject.Cell(5, 1).Range.Font.Name = "Roboto Light";
        //            tblproject.Cell(5, 1).Range.Font.Size = 10;
        //            tblproject.Cell(6, 1).Range.Text = "Deal Strategy";
        //            tblproject.Cell(6, 2).Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].DealStrategy));
        //            tblproject.Cell(6, 1).Range.Font.Name = "Roboto Light";
        //            tblproject.Cell(6, 1).Range.Font.Size = 10;
        //            #endregion

        //            #region B.Borrower
        //            Range borrower = app.ActiveDocument.Bookmarks["Borrower"].Range;
        //            Table tblborrower = app.ActiveDocument.Tables.Add(borrower, 2, 5, WdDefaultTableBehavior.wdWord9TableBehavior);
        //            tblborrower.Range.Font.Name = "Roboto Light";
        //            tblborrower.Range.Font.Size = 10;
        //            tblborrower.set_Style("Table Grid");
        //            tblborrower.Columns[1].SetWidth(100, WdRulerStyle.wdAdjustFirstColumn);

        //            tblborrower.Cell(1, 1).Range.Text = "Project Company";

        //            string prevkey = "";
        //            string cellText = "";
        //            foreach (DataRow items in listBorrower.Rows)
        //            {
        //                prevkey = items[0].ToString() + "\r\a";
        //                if (cellText != prevkey)
        //                {
        //                    tblborrower.Cell(1, 2).Range.Text = items[0].ToString();
        //                    cellText = tblborrower.Cell(1, 2).Range.Text;
        //                }
        //            }

        //            tblborrower.Cell(2, 1).Range.Text = "Project Sponsors";
        //            tblborrower.Cell(2, 1).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;

        //            tblborrower.Cell(2, 2).SetWidth(140, WdRulerStyle.wdAdjustFirstColumn);
        //            tblborrower.Cell(2, 2).Range.Text = "Project Company";
        //            tblborrower.Cell(2, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            tblborrower.Cell(2, 2).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;

        //            tblborrower.Cell(2, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            tblborrower.Cell(2, 3).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
        //            tblborrower.Cell(2, 3).Range.Text = "Name Project Sponsors";

        //            tblborrower.Cell(2, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            tblborrower.Cell(2, 5).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
        //            tblborrower.Cell(2, 5).Range.Text = "% ownership";

        //            string prevkey1 = "";
        //            string cellText1 = "";
        //            int rowCount = 2;
        //            foreach (DataRow item in listBorrower.Rows)
        //            {
        //                rowCount++;
        //                tblborrower.Rows.Add(ref missing);

        //                prevkey1 = item[0].ToString() + "\r\a";
        //                if (cellText1 != prevkey1)
        //                {
        //                    tblborrower.Rows[rowCount].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
        //                    tblborrower.Rows[rowCount].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
        //                    tblborrower.Rows[rowCount].Cells[5].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
        //                    tblborrower.Rows[rowCount].Cells[2].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
        //                    tblborrower.Rows[rowCount].Cells[3].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
        //                    tblborrower.Rows[rowCount].Cells[5].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;

        //                    tblborrower.Rows[rowCount].Cells[2].Range.Text = item[0].ToString();
        //                    cellText1 = tblborrower.Rows[rowCount].Cells[2].Range.Text;
        //                    tblborrower.Cell(1, 2).Range.Text = item[0].ToString() + " ";
        //                    tblborrower.Cell(1, 2).Range.Font.Name = "Roboto Light";
        //                    tblborrower.Cell(1, 2).Range.Font.Size = 10;
        //                }
        //                tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
        //                tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
        //                tblborrower.Rows[rowCount].Cells[3].Range.Text = item[1].ToString();
        //                var ownership = Convert.ToDecimal(item[2]);
        //                tblborrower.Rows[rowCount].Cells[5].Range.Text = ownership.ToString("#,#");
        //                tblborrower.Rows[rowCount].Cells[5].Range.Font.Name = "Roboto Light";
        //                tblborrower.Rows[rowCount].Cells[5].Range.Font.Size = 10;

        //            }

        //            rowCount++;
        //            tblborrower.Rows.Add(ref missing);
        //            tblborrower.Rows[rowCount].Cells[1].Range.Text = "Ultimate Beneficial Owner";
        //            tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
        //            tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
        //            int rowtemp = rowCount;

        //            rowCount++;
        //            tblborrower.Rows.Add(ref missing);
        //            tblborrower.Rows[rowCount].Cells[1].Range.Text = "Rating";
        //            tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
        //            tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
        //            tblborrower.Rows[rowCount].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            tblborrower.Rows[rowCount].Cells[2].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
        //            tblborrower.Rows[rowCount].Cells[2].Range.Text = "IIF Rating";
        //            tblborrower.Rows[rowCount].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            tblborrower.Rows[rowCount].Cells[3].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
        //            tblborrower.Rows[rowCount].Cells[3].Range.Text = "External Rating";
        //            tblborrower.Rows[rowCount].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            tblborrower.Rows[rowCount].Cells[4].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
        //            tblborrower.Rows[rowCount].Cells[4].Range.Text = "S&E Category";
        //            tblborrower.Rows[rowCount].Cells[5].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            tblborrower.Rows[rowCount].Cells[5].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
        //            tblborrower.Rows[rowCount].Cells[5].Range.Text = "LQC/BI Checking";

        //            rowCount++;
        //            tblborrower.Rows.Add(ref missing);
        //            tblborrower.Rows[rowCount].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        //            tblborrower.Rows[rowCount].Cells[2].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
        //            tblborrower.Rows[rowCount].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        //            tblborrower.Rows[rowCount].Cells[3].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
        //            tblborrower.Rows[rowCount].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        //            tblborrower.Rows[rowCount].Cells[4].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
        //            tblborrower.Rows[rowCount].Cells[5].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        //            tblborrower.Rows[rowCount].Cells[5].Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;

        //            tblborrower.Rows[rowCount].Cells[2].Range.Text = dataResult[0].IIFRate;
        //            tblborrower.Rows[rowCount].Cells[3].Range.Text = "S&P: " + dataResult[0].SAndPRate;
        //            tblborrower.Rows[rowCount].Cells[4].Range.Text = dataResult[0].SAndECategoryRate;
        //            tblborrower.Rows[rowCount].Cells[5].Range.Text = dataResult[0].LQCOrBICheckingRate;

        //            tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
        //            tblborrower.Rows[rowCount].Cells[2].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
        //            tblborrower.Rows[rowCount].Cells[3].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
        //            tblborrower.Rows[rowCount].Cells[4].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
        //            tblborrower.Rows[rowCount].Cells[5].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;

        //            tblborrower.Rows[rowCount].Cells[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
        //            tblborrower.Rows[rowCount].Cells[2].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
        //            tblborrower.Rows[rowCount].Cells[3].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
        //            tblborrower.Rows[rowCount].Cells[4].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
        //            tblborrower.Rows[rowCount].Cells[5].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;

        //            rowCount++;
        //            tblborrower.Rows.Add(ref missing);
        //            tblborrower.Rows[rowCount].Cells[2].Range.Text = "Rating Date: " + dataResult[0].IIFRatingDate.Value.ToString("dd-MMM-yyyy");
        //            tblborrower.Rows[rowCount].Cells[3].Range.Text = "Moodys: " + dataResult[0].MoodysRate;

        //            rowCount++;
        //            tblborrower.Rows.Add(ref missing);
        //            tblborrower.Rows[rowCount].Cells[3].Range.Text = "Fitch: " + dataResult[0].FitchRate;

        //            rowCount++;
        //            tblborrower.Rows.Add(ref missing);
        //            tblborrower.Rows[rowCount].Cells[3].Range.Text = "Pefindo: " + dataResult[0].PefindoRate;

        //            rowCount++;
        //            tblborrower.Rows.Add(ref missing);
        //            tblborrower.Rows[rowCount].Cells[1].Range.Text = "Other information";
        //            tblborrower.Rows[rowCount].Cells[2].Merge(tblborrower.Rows[rowCount].Cells[5]);
        //            tblborrower.Rows[rowCount].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].OtherInformation));
        //            tblborrower.Rows[rowCount].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
        //            tblborrower.Rows[rowCount].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
        //            tblborrower.Rows[rowCount].Cells[2].Range.Font.Name = "Robaoto Light";
        //            tblborrower.Rows[rowCount].Cells[2].Range.Font.Size = 10;

        //            tblborrower.Cell(1, 2).Merge(tblborrower.Cell(1, 5));
        //            tblborrower.Cell(2, 3).Merge(tblborrower.Cell(2, 4));
        //            tblborrower.Cell(2, 3).SetWidth(130, WdRulerStyle.wdAdjustSameWidth);

        //            for (int a = 1; a <= listBorrower.Rows.Count; a++)
        //            {
        //                tblborrower.Cell(2 + a, 3).Merge(tblborrower.Cell(2 + a, 4));
        //                tblborrower.Cell(2 + a, 3).SetWidth(130, WdRulerStyle.wdAdjustSameWidth);
        //            }
        //            tblborrower.Rows[rowtemp].Cells[2].Merge(tblborrower.Rows[rowtemp].Cells[5]);
        //            #endregion

        //            #region C.Proposal
        //            Range proposal = app.ActiveDocument.Bookmarks["Proposal"].Range;
        //            Table tblProposal = app.ActiveDocument.Tables.Add(proposal, 3, 3, WdDefaultTableBehavior.wdWord9TableBehavior);
        //            tblProposal.Range.Font.Name = "Roboto Light";
        //            tblProposal.Range.Font.Size = 10;
        //            tblProposal.set_Style("Table Grid");
        //            tblProposal.Columns[1].SetWidth(130, WdRulerStyle.wdAdjustFirstColumn);
        //            tblProposal.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

        //            tblProposal.Cell(1, 1).Range.Text = "Purpose";
        //            tblProposal.Cell(1, 2).Merge(tblProposal.Cell(1, 3));
        //            tblProposal.Cell(1, 2).Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].Purpose));
        //            tblProposal.Cell(1, 2).Range.Font.Name = "Roboto Light";

        //            tblProposal.Cell(2, 1).Range.Text = "Approval Authority";
        //            tblProposal.Cell(2, 2).Merge(tblProposal.Cell(2, 3));
        //            tblProposal.Cell(2, 2).Range.Text = dataResult[0].ApprovalAuthority;

        //            tblProposal.Cell(3, 1).Range.Text = "Facility";
        //            tblProposal.Cell(3, 2).Shading.BackgroundPatternColor = WdColor.wdColorGray40;
        //            tblProposal.Cell(3, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            tblProposal.Cell(3, 2).Range.Text = "Type";
        //            tblProposal.Cell(3, 3).Shading.BackgroundPatternColor = WdColor.wdColorGray40;
        //            tblProposal.Cell(3, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            tblProposal.Cell(3, 3).Range.Text = "Amount";

        //            int rowcounttblProposal = 3;
        //            foreach (DataRow item in listFacility.Rows)
        //            {
        //                rowcounttblProposal++;
        //                tblProposal.Rows.Add(ref missing);
        //                tblProposal.Rows[rowcounttblProposal].Cells[2].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
        //                tblProposal.Rows[rowcounttblProposal].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        //                tblProposal.Rows[rowcounttblProposal].Cells[3].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
        //                tblProposal.Rows[rowcounttblProposal].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        //                tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = item[0].ToString();
        //                var amount = Convert.ToDecimal(item[2]);
        //                tblProposal.Rows[rowcounttblProposal].Cells[3].Range.Text = item[1].ToString() + " " + amount.ToString("#,#");
        //            }

        //            rowcounttblProposal++;
        //            tblProposal.Rows.Add(ref missing);
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        //            tblProposal.Rows[rowcounttblProposal].Cells[3].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
        //            tblProposal.Rows[rowcounttblProposal].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

        //            tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Group Exposure";
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = dataResult[0].GroupExposureCurr + " " + dataResult[0].GroupExposureAmount;
        //            int GroupExpoCount = rowcounttblProposal;

        //            rowcounttblProposal++;
        //            tblProposal.Rows.Add(ref missing);
        //            tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Remarks";
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].Remarks));
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Font.Name = "Roboto Light";
        //            int RemarksCount = rowcounttblProposal;

        //            rowcounttblProposal++;
        //            tblProposal.Rows.Add(ref missing);
        //            tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Tenor";
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = dataResult[0].tenorYear + " year(s)  " + dataResult[0].tenorMonth + " month(s)";
        //            int tenorCount = rowcounttblProposal;

        //            rowcounttblProposal++;
        //            tblProposal.Rows.Add(ref missing);
        //            tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Average Loan Life";
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = dataResult[0].averageLoanLifeYear + " year(s)  " + dataResult[0].averageLoanLifeMonth + " month(s)";
        //            int AverageCount = rowcounttblProposal;

        //            rowcounttblProposal++;
        //            tblProposal.Rows.Add(ref missing);
        //            tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Pricing";
        //            tblProposal.Rows[rowcounttblProposal].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = "Interest rate";
        //            tblProposal.Rows[rowcounttblProposal].Cells[3].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].pricingInterestRate));
        //            tblProposal.Rows[rowcounttblProposal].Cells[3].Range.Font.Name = "Roboto Light";

        //            rowcounttblProposal++;
        //            tblProposal.Rows.Add(ref missing);
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = "Commitment Fee";
        //            tblProposal.Rows[rowcounttblProposal].Cells[3].Range.Text = dataResult[0].pricingCommitmentFee;

        //            rowcounttblProposal++;
        //            tblProposal.Rows.Add(ref missing);
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = "Upfront Fee";
        //            tblProposal.Rows[rowcounttblProposal].Cells[3].Range.Text = dataResult[0].pricingUpfrontFacilityFee;

        //            rowcounttblProposal++;
        //            tblProposal.Rows.Add(ref missing);
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = "Structuring Fee";
        //            tblProposal.Rows[rowcounttblProposal].Cells[3].Range.Text = dataResult[0].pricingStructuringFee;

        //            rowcounttblProposal++;
        //            tblProposal.Rows.Add(ref missing);
        //            tblProposal.Rows[rowcounttblProposal].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = "Arranger Fee";
        //            tblProposal.Rows[rowcounttblProposal].Cells[3].Range.Text = dataResult[0].pricingArrangerFee;

        //            rowcounttblProposal++;
        //            tblProposal.Rows.Add(ref missing);
        //            tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Collateral";
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].pricingCollateral));
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Font.Name = "Roboto Light";
        //            tblProposal.Cell(rowcounttblProposal, 2).Merge(tblProposal.Cell(rowcounttblProposal, 3));

        //            rowcounttblProposal++;
        //            tblProposal.Rows.Add(ref missing);
        //            tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Other Conditions";
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].pricingOtherConditions));
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Font.Name = "Roboto Light";

        //            rowcounttblProposal++;
        //            tblProposal.Rows.Add(ref missing);
        //            tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Exception to IIF Policy";
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].pricingExceptionToIIFPolicy));
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Font.Name = "Roboto Light";

        //            rowcounttblProposal++;
        //            tblProposal.Rows.Add(ref missing);
        //            tblProposal.Rows[rowcounttblProposal].Cells[1].Range.Text = "Review Period";
        //            tblProposal.Rows[rowcounttblProposal].Cells[2].Range.Text = dataResult[0].reviewPeriod;

        //            tblProposal.Cell(GroupExpoCount, 2).Merge(tblProposal.Cell(GroupExpoCount, 3));
        //            tblProposal.Cell(RemarksCount, 2).Merge(tblProposal.Cell(RemarksCount, 3));
        //            tblProposal.Cell(tenorCount, 2).Merge(tblProposal.Cell(tenorCount, 3));
        //            tblProposal.Cell(AverageCount, 2).Merge(tblProposal.Cell(AverageCount, 3));
        //            #endregion

        //            #region D.Recommendation
        //            Range keyInvestment = app.ActiveDocument.Bookmarks["KeyInvestment"].Range;
        //            keyInvestment.Font.Name = "Roboto Light";
        //            keyInvestment.Font.Size = 10;
        //            Paragraph paragkeyInvestment = doc.Content.Paragraphs.Add(keyInvestment);
        //            paragkeyInvestment.Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].KeyInvestmentRecommendation));
        //            paragkeyInvestment.Range.Font.Name = "Roboto Light";

        //            Range recommendation = app.ActiveDocument.Bookmarks["Recommendation"].Range;
        //            recommendation.Font.Name = "Roboto Light";
        //            recommendation.Font.Size = 10;
        //            Paragraph paragrecommendation = doc.Content.Paragraphs.Add(recommendation);
        //            paragrecommendation.Range.InsertFile(ConvertHtmlAndFile.SaveToHtml(dataResult[0].Recommendation));
        //            paragrecommendation.Range.Font.Name = "Roboto Light";
        //            paragrecommendation.Range.Font.Size = 10;

        //            Range accountResponsible = app.ActiveDocument.Bookmarks["AccountResponsible"].Range;
        //            Table tblAccountResponsible = app.ActiveDocument.Tables.Add(accountResponsible, 1, 3, WdDefaultTableBehavior.wdWord9TableBehavior);
        //            tblAccountResponsible.Range.Font.Name = "Roboto Light";
        //            tblAccountResponsible.Range.Font.Size = 10;
        //            tblAccountResponsible.set_Style("Table Grid");
        //            tblAccountResponsible.Columns[1].SetWidth(130, WdRulerStyle.wdAdjustFirstColumn);
        //            tblAccountResponsible.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

        //            tblAccountResponsible.Cell(1, 1).Range.Text = "Account Responsible";
        //            tblAccountResponsible.Cell(1, 2).Range.Text = "Deal Team";
        //            tblAccountResponsible.Cell(1, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            tblAccountResponsible.Cell(1, 2).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;
        //            tblAccountResponsible.Cell(1, 3).Range.Text = "CIO";
        //            tblAccountResponsible.Cell(1, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            tblAccountResponsible.Cell(1, 3).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray40;

        //            int rowCountDealTeam = 1;
        //            foreach (DataRow item in listDealTeam.Rows)
        //            {
        //                rowCountDealTeam++;
        //                tblAccountResponsible.Rows.Add(ref missing);
        //                tblAccountResponsible.Cell(rowCountDealTeam, 1).Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
        //                tblAccountResponsible.Cell(rowCountDealTeam, 1).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
        //                tblAccountResponsible.Rows[rowCountDealTeam].Cells[2].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
        //                tblAccountResponsible.Rows[rowCountDealTeam].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        //                tblAccountResponsible.Rows[rowCountDealTeam].Cells[3].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
        //                tblAccountResponsible.Rows[rowCountDealTeam].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        //                tblAccountResponsible.Rows[rowCountDealTeam].Cells[2].Range.Text = item[0].ToString();
        //                if (rowCountDealTeam <= 2)
        //                {
        //                    tblAccountResponsible.Rows[rowCountDealTeam].Cells[3].Range.Text = dataResult[0].AccountResponsibleCIOName;
        //                }
        //            }
        //            tblAccountResponsible.Cell(rowCountDealTeam, 1).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
        //            #endregion

        //            #region Attachment
        //            this.FillBookmarkWithPAMAttachmentType1(app, con, "ProjectAnalysis", AppConstants.TableName.PAM_ProjectAnalysis, pamId);
        //            this.FillBookmarkWithPAMAttachmentType1(app, con, "HistoricalFinancialandFinancialProject", AppConstants.TableName.PAM_HistoricalFinancial, pamId);
        //            this.FillBookmarkWithPAMAttachmentType1(app, con, "Supplemental", AppConstants.TableName.PAM_Supplemental, pamId);
        //            this.FillBookmarkWithPAMAttachmentType1(app, con, "SocialEnvironmental", AppConstants.TableName.PAM_Social, pamId);
        //            this.FillBookmarkWithPAMAttachmentType1(app, con, "GroupStructure", AppConstants.TableName.PAM_GroupStructure, pamId);
        //            this.FillBookmarkWithPAMAttachmentType1(app, con, "TermSheet", AppConstants.TableName.PAM_TermSheet, pamId);
        //            this.FillBookmarkWithPAMAttachmentType1(app, con, "riskRating", AppConstants.TableName.PAM_RiskRating, pamId);
        //            this.FillBookmarkWithPAMAttachmentType1(app, con, "kycChecklists", AppConstants.TableName.PAM_KYCChecklists, pamId);
        //            this.FillBookmarkWithPAMAttachmentType1(app, con, "Otherbanks", AppConstants.TableName.PAM_OtherBanksFacilities, pamId);
        //            this.FillBookmarkWithPAMAttachmentType1(app, con, "IndustryAnalysis", AppConstants.TableName.PAM_IndustryAnalysis, pamId);

        //            this.FillBookmarkWithPAMAttachmentType2(app, con, "LegalDuediligenceReport", AppConstants.TableName.PAM_LegalDueDiligenceReport, pamId);
        //            this.FillBookmarkWithPAMAttachmentType2(app, con, "SAndEDueDiligence", AppConstants.TableName.PAM_SAndEDueDiligence, pamId);
        //            this.FillBookmarkWithPAMAttachmentType2(app, con, "OtherReport", AppConstants.TableName.PAM_OtherReports, pamId);
        //            #endregion

        //            doc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
        //            doc.SaveAs2(Path.Combine(temporaryFolderLocation, fileNamePDF), WdExportFormat.wdExportFormatPDF);
        //        }
        //        finally
        //        {
        //            doc.Close(WdSaveOptions.wdDoNotSaveChanges);
        //        }
        //    }
        //    finally
        //    {
        //        app.Quit();
        //    }

        //    File.Delete(destFile);
        //    string destFilePDF = Path.Combine(temporaryFolderLocation, fileNamePDF);
        //    byte[] fileContent = File.ReadAllBytes(destFilePDF);

        //    FileMergeResult result = new FileMergeResult();
        //    result.FileContent = fileContent;
        //    result.FileName = fileNamePDF;
        //    return result;
        //}
    }
}