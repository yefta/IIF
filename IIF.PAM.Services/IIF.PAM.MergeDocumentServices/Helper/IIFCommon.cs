using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;

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
			res.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
			res.PreferredWidth = 100;

			res.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

			if(fillHeaderBgColor)
				res.Rows[1].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray10;
			
			return res;
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

		public static void createLegalSAndEDueOtherReportTable(Application app, System.Data.DataTable listData, string bookmarkNameAttachment, string bookmarkNameDescription)
		{
			object missing = System.Reflection.Missing.Value;
			Table tblAttachment = IIFCommon.createTable(app, bookmarkNameAttachment, 1, false);
			tblAttachment.Borders.Enable = 0;
			int rowCounter = 0;
			foreach (DataRow item in listData.Rows)
			{
				tblAttachment.Rows.Add(ref missing);
				rowCounter++;
				string valueAttachmentName = IIFCommon.readValueHTMLString(item[0].ToString(), "name");
				tblAttachment.Cell(rowCounter, 1).Range.Text = valueAttachmentName;
				tblAttachment.Cell(rowCounter, 1).Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
				tblAttachment.Cell(rowCounter, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
			}
			tblAttachment.Rows[rowCounter + 1].Delete();

			Table tblDescription = IIFCommon.createTable(app, bookmarkNameDescription, 1, false);
			tblDescription.Borders.Enable = 0;
			rowCounter = 0;
			foreach (DataRow item in listData.Rows)
			{
				tblDescription.Rows.Add(ref missing);
				rowCounter++;
				tblDescription.Cell(rowCounter, 1).Range.Text = item[1].ToString();
				tblDescription.Cell(rowCounter, 1).Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
				tblDescription.Cell(rowCounter, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
			}
			tblDescription.Rows[rowCounter + 1].Delete();
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
		}
	}
}
