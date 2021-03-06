﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using IIF.PAM.MergeDocumentServices.Models;
using Microsoft.Office.Interop.Word;

using ParagraphsList =
	System.Collections.Generic.List<Microsoft.Office.Interop.Word.Paragraph>;

namespace IIF.PAM.MergeDocumentServices.Helper
{
	public class IIFCommon
	{
		public static Table createTable(Application app, string bookmarkString, int columnCount, bool fillHeaderBgColor)
		{
			Range myRange = app.ActiveDocument.Bookmarks[bookmarkString].Range;

			Table res = app.ActiveDocument.Tables.Add(myRange, 1, columnCount, WdDefaultTableBehavior.wdWord9TableBehavior);
			res.Range.Font.Name = "Roboto Light";
			res.Range.Font.Size = 10;
			res.set_Style("Table Grid");
			res.AllowAutoFit = true;			
			res.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);			
			res.PreferredWidth = 100;			

			res.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

			if (fillHeaderBgColor)
				res.Rows[1].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray10;

			return res;
		}

		public static void createPreviousApproval(Application app, System.Data.DataTable listData, string bookmarkName, List<CMData> dataResult, string currFontFamily, float currFontSize)
		{
			Table tblPreviousApproval = IIFCommon.createTable(app, bookmarkName, 5, true);
			tblPreviousApproval.Borders.Enable = 1;

			//header
			tblPreviousApproval.Cell(1, 1).Range.Text = "Type Document";
			tblPreviousApproval.Cell(1, 2).Range.Text = "No. Document";
			tblPreviousApproval.Cell(1, 3).Range.Text = "Approval";
			tblPreviousApproval.Cell(1, 4).Range.Text = "Approval Date";
			tblPreviousApproval.Cell(1, 5).Range.Text = "Purpose";

			int rowCounter = 1;
			object missing = System.Reflection.Missing.Value;

			string cmNumber = IIFCommon.generateCMNumber(
				dataResult[0].ProjectCode
				, Convert.ToInt32(dataResult[0].CMNumber).ToString("00")
				, dataResult[0].ApprovalAuhority
				, dataResult[0].CMDate.ToString("MMM")
				, dataResult[0].CMDate.ToString("yyyy")
				);

			foreach (DataRow item in listData.Rows)
			{
				if (item[3].ToString() == null || item[3].ToString().Trim().Length == 0) //ApprovalDate
					continue;

				tblPreviousApproval.Rows.Add(ref missing);
				rowCounter++;

				tblPreviousApproval.Cell(rowCounter, 1).Range.Text = item[1].ToString();
				tblPreviousApproval.Cell(rowCounter, 1).Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
				tblPreviousApproval.Cell(rowCounter, 1).Range.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				tblPreviousApproval.Cell(rowCounter, 1).Range.Font.Name = currFontFamily;
				tblPreviousApproval.Cell(rowCounter, 1).Range.Font.Size = currFontSize;

				tblPreviousApproval.Cell(rowCounter, 2).Range.Text = item[2].ToString();
				tblPreviousApproval.Cell(rowCounter, 2).Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
				tblPreviousApproval.Cell(rowCounter, 2).Range.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				tblPreviousApproval.Cell(rowCounter, 2).Range.Font.Name = currFontFamily;
				tblPreviousApproval.Cell(rowCounter, 2).Range.Font.Size = currFontSize;

				tblPreviousApproval.Cell(rowCounter, 3).Range.Text = item[3].ToString();
				tblPreviousApproval.Cell(rowCounter, 3).Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
				tblPreviousApproval.Cell(rowCounter, 3).Range.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				tblPreviousApproval.Cell(rowCounter, 3).Range.Font.Name = currFontFamily;
				tblPreviousApproval.Cell(rowCounter, 3).Range.Font.Size = currFontSize;

				tblPreviousApproval.Cell(rowCounter, 4).Range.Text = Convert.ToDateTime(item[4].ToString()).ToString("dd MMM yyyy");
				tblPreviousApproval.Cell(rowCounter, 4).Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
				tblPreviousApproval.Cell(rowCounter, 4).Range.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				tblPreviousApproval.Cell(rowCounter, 4).Range.Font.Name = currFontFamily;
				tblPreviousApproval.Cell(rowCounter, 4).Range.Font.Size = currFontSize;

				string htmlResult = ConvertHtmlAndFile.SaveToFile(item[5].ToString());

				#region delete empty paragraph
				Application app2 = new Application();
				Document sourceDocument = app2.Documents.Open(htmlResult);
				object start = sourceDocument.Content.Start;
				object end = sourceDocument.Content.End;
				Microsoft.Office.Interop.Word.Range myRange = sourceDocument.Range(ref start, ref end);
				myRange.Select();
				FindEmptyParagraphsAndDelete(sourceDocument);
				//myRange.set_Style(ref Normal);
				sourceDocument.Save();
				sourceDocument.Close(WdSaveOptions.wdSaveChanges);
				app2.Quit();
				#endregion

				tblPreviousApproval.Cell(rowCounter, 5).Range.InsertFile(htmlResult);
				tblPreviousApproval.Cell(rowCounter, 5).Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
				tblPreviousApproval.Cell(rowCounter, 5).Range.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				tblPreviousApproval.Cell(rowCounter, 5).Range.Font.Name = currFontFamily;
				tblPreviousApproval.Cell(rowCounter, 5).Range.Font.Size = currFontSize;
			}

		}

		public static void FindEmptyParagraphsAndDelete(Document document)
		{
			ParagraphsList list = new ParagraphsList();
			foreach (Paragraph para in document.Content.Paragraphs)
			{
				if ((para.Range.End - para.Range.Start) <= 2)
					list.Add(para);
			}

			foreach (Paragraph myPar in list)
			{
				myPar.Range.Delete();
			}

		}

		public static string generateCMNumber(string projectCode, string noUrut, string wewenangPemutus, string bulan, string tahun)
		{
			string res = "";

			string borrowerCode = projectCode.Substring(5, 3);

			res = borrowerCode + "-" + noUrut + "/" + wewenangPemutus + "/" + bulan + "/" + tahun;
			return res;
		}

		public static string readValueHTMLString(string htmlString, string tag)
		{
			string res = "";

			htmlString = htmlString.Split(new string[] { "</" + tag + ">" }, StringSplitOptions.None)[0];
			htmlString = htmlString.Split(new string[] { "<" + tag + ">" }, StringSplitOptions.None)[1];

			res = htmlString;

			return res;
		}

		public static void createLegalSAndEDueOtherReportTable(Application app, System.Data.DataTable listData, string bookmarkNameAttachment, string bookmarkNameDescription, string currFontFamily, float currFontSize)
		{
			object missing = System.Reflection.Missing.Value;
			Table tblAttachment = IIFCommon.createTable(app, bookmarkNameAttachment, 1, false);
			tblAttachment.Columns[1].Width = 250;
			tblAttachment.Borders.Enable = 0;
			int rowCounter = 0;
			foreach (DataRow item in listData.Rows)
			{
				tblAttachment.Rows.Add(ref missing);
				rowCounter++;
				string valueAttachmentName = IIFCommon.readValueHTMLString(item[0].ToString(), "name");
				if (valueAttachmentName.ToLower() == "scnull")
					valueAttachmentName = "-";
				tblAttachment.Cell(rowCounter, 1).Range.Text = valueAttachmentName;
				tblAttachment.Cell(rowCounter, 1).Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
				tblAttachment.Cell(rowCounter, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
			}
			tblAttachment.Rows[rowCounter + 1].Delete();

			Table tblDescription = IIFCommon.createTable(app, bookmarkNameDescription, 1, false);
			tblDescription.Columns[1].Width = 250;
			tblDescription.Borders.Enable = 0;
			rowCounter = 0;
			foreach (DataRow item in listData.Rows)
			{
				tblDescription.Rows.Add(ref missing);
				rowCounter++;
				tblDescription.Cell(rowCounter, 1).Range.Text = item[1].ToString();
				tblDescription.Cell(rowCounter, 1).Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
				tblDescription.Cell(rowCounter, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				tblDescription.Cell(rowCounter, 1).Range.Font.Name = currFontFamily;
				tblDescription.Cell(rowCounter, 1).Range.Font.Size = currFontSize;
			}
			tblDescription.Rows[rowCounter + 1].Delete();
		}

		public static void copyFromNetwork(string source, string destination, string foldertemplate, string temporaryFolderLocation)
		{
			string NETWORK_USER_NAME = "";
			string NETWORK_USER_PASSWORD = "";

			try
			{
				NETWORK_USER_NAME = ConfigurationManager.AppSettings["NETWORK_USER_NAME"];
			}
			catch { NETWORK_USER_NAME = ""; }

			try
			{
				NETWORK_USER_PASSWORD = ConfigurationManager.AppSettings["NETWORK_USER_PASSWORD"];
			}
			catch { NETWORK_USER_PASSWORD = ""; }

			if (!String.IsNullOrEmpty(NETWORK_USER_NAME) && !String.IsNullOrEmpty(NETWORK_USER_PASSWORD))
			{
				IIF.PAM.MergeDocumentServices.Helper.NetworkShare.DisconnectFromShare(foldertemplate, true);
				IIF.PAM.MergeDocumentServices.Helper.NetworkShare.DisconnectFromShare(temporaryFolderLocation, true);

				IIF.PAM.MergeDocumentServices.Helper.NetworkShare.ConnectToShare(foldertemplate, NETWORK_USER_NAME, NETWORK_USER_PASSWORD);
				IIF.PAM.MergeDocumentServices.Helper.NetworkShare.ConnectToShare(temporaryFolderLocation, NETWORK_USER_NAME, NETWORK_USER_PASSWORD);
			}

			File.Copy(source, destination, true);

			if (!String.IsNullOrEmpty(NETWORK_USER_NAME) && !String.IsNullOrEmpty(NETWORK_USER_PASSWORD))
			{
				IIF.PAM.MergeDocumentServices.Helper.NetworkShare.DisconnectFromShare(foldertemplate, false);
				IIF.PAM.MergeDocumentServices.Helper.NetworkShare.DisconnectFromShare(temporaryFolderLocation, false);
			}
		}

		public static void finalizeDoc(Document doc)
		{
			try
			{
				doc.TablesOfContents[1].UpdatePageNumbers();
			}
			catch { }
			try
			{
				doc.Revisions.AcceptAll();
			}
			catch { }
			try
			{
				doc.DeleteAllComments();
			}
			catch { }

			try
			{
				object start = doc.Content.Start;
				object end = doc.Content.End;
				Microsoft.Office.Interop.Word.Range myRange = doc.Range(ref start, ref end);
				myRange.Select();
				myRange.Font.Name = "Roboto Light";			
			}
			catch { }

			foreach (Section mySec in doc.Sections)
			{
				if (mySec.Index % 2 == 0) //section genap -> landscape
				{
					mySec.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
				}
			}
		}

		public static void injectFooterPAM(Document doc, string projectCode)
		{
			try
			{
				object currentPage = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage;
				string privateStr = "\tPrivate & Confidential";

				foreach (Section mySec in doc.Sections)
				{
					Range footerRangePage = mySec.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
					footerRangePage.Font.Size = 9;
					footerRangePage.Fields.Add(footerRangePage, currentPage);
					footerRangePage.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					footerRangePage.InsertBefore(projectCode + "\t");
					footerRangePage.InsertAfter(privateStr);					
				}								
			}
			catch { }
		}		

		public static void injectFooterCM(Document doc, string footerDate, string cmType)
		{
			try
			{
				object currentPage = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage;

				if (footerDate.Contains("0001"))
					footerDate = "";

				foreach (Section mySec in doc.Sections)
				{
					Range footerRangePage = mySec.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
					footerRangePage.Fields.Add(footerRangePage, currentPage);
					footerRangePage.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
					footerRangePage.Font.Size = 9;
					if (cmType == "Waiver")
						footerRangePage.InsertBefore("Credit Memorandum – Project/Corporate Loan/Equity " + footerDate + "\t\t");
					if (cmType == "Project")
						footerRangePage.InsertBefore("Periodic Review Memorandum – Project Loan " + footerDate + "\t\t");
					if (cmType == "Equity")
						footerRangePage.InsertBefore("Periodic Review Memorandum – Equity " + footerDate + "\t\t");
					if (cmType == "Corporate")
						footerRangePage.InsertBefore("Periodic Review Memorandum – Corporate Loan " + footerDate + "\t\t");
				}
			}
			catch { }
		}

		public static string fileNameFormat(System.Data.DataTable listDocVersion, string fileNamePDF, bool isPreview = false)
		{
			string res = "";
			try
			{
				string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileNamePDF);
				if (listDocVersion.Rows.Count == 0)
				{
					if(isPreview)
						res = fileNameWithoutExt + "-v0.5.pdf";
					else
						res = fileNameWithoutExt + "-v1.0.pdf";
				}
				else
				{
					if (string.IsNullOrEmpty(listDocVersion.Rows[0]["LastVersion"].ToString()))
					{
						if (isPreview)
							res = fileNameWithoutExt + "-v0.5.pdf";
						else
							res = fileNameWithoutExt + "-v1.0.pdf";
					}
					else
					{
						int lastVersion = Convert.ToInt32(listDocVersion.Rows[0]["LastVersion"]);

						if (isPreview)
							res = fileNameWithoutExt + "-v" + (lastVersion) + ".5.pdf";
						else
							res = fileNameWithoutExt + "-v" + (lastVersion + 1) + ".0.pdf";
					}
				}

				return res;
			}
			catch { return null; }
		}
		
	}
}
