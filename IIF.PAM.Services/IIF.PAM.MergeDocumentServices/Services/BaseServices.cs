using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;

using Microsoft.Office.Interop.Word;

using log4net;

using IIF.PAM.MergeDocumentServices.Helper;
using IIF.PAM.MergeDocumentServices.Models;
using System.IO;
using System.Threading;

namespace IIF.PAM.MergeDocumentServices.Services
{
    public class BaseServices
    {
        private ILog _Logger = null;
        protected ILog Logger
        {
            get
            {
                if (_Logger == null)
                {
                    _Logger = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);                    
                }
                return _Logger;
            }
        }

        protected SqlParameter NewSqlParameter(string parameterName, SqlDbType dbType, object value)
        {
            SqlParameter result = new SqlParameter(parameterName, dbType);
            result.Value = value;
            return result;
        }

        protected void FillBookmarkWithPAMAttachmentNormal(Application app, SqlConnection con, string bookmarkName, string tableName, long pamId)
        {
            this.Logger.Info(tableName);
		
			object Normal = "Normal";

			string query = "SELECT";
            query = query + " Attachment";
            query = query + " FROM " + tableName;
            query = query + " WHERE PAMId = @PAMId";
            query = query + " ORDER BY OrderNumber DESC";

            using (SqlCommand cmd = con.CreateCommand())
            {
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = query;
                cmd.Parameters.Add(this.NewSqlParameter("PAMId", SqlDbType.BigInt, pamId));
                
                Range rangeBookmark = app.ActiveDocument.Bookmarks[bookmarkName].Range;				

				using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string fileContent = dr.GetNullableString(0);
						if (fileContent != null)
						{
							string htmlResult = ConvertHtmlAndFile.SaveToFile(fileContent);

							try
							{
								Application app2 = new Application();
								Document sourceDocument = app2.Documents.Open(htmlResult);
								object start = sourceDocument.Content.Start;
								object end = sourceDocument.Content.End;
								Microsoft.Office.Interop.Word.Range myRange = sourceDocument.Range(ref start, ref end);
								myRange.Select();
								//myRange.set_Style(ref Normal);								
								sourceDocument.Save();
								sourceDocument.Close(WdSaveOptions.wdSaveChanges);
								app2.Quit();

								rangeBookmark.InsertFile(htmlResult);
							}
							catch (Exception ex)
							{
								throw ex;
							}
							finally
							{
								File.Delete(htmlResult);
							}
						}
                    }
                }

				
			}            
        }


		protected void FillBookmarkWithPAMAttachmentABNormal(Application app, SqlConnection con, string bookmarkName, string tableName, long pamId, string columnName, string columnPAMId)
		{
			this.Logger.Info(tableName);

			object Normal = "Normal";

			string query = "SELECT";
			query = query + " " + columnName;
			query = query + " FROM " + tableName;
			query = query + " WHERE " + columnPAMId + " = @PAMId";			

			using (SqlCommand cmd = con.CreateCommand())
			{
				cmd.CommandType = CommandType.Text;
				cmd.CommandText = query;
				cmd.Parameters.Add(this.NewSqlParameter("PAMId", SqlDbType.BigInt, pamId));

				Range rangeBookmark = app.ActiveDocument.Bookmarks[bookmarkName].Range;				

				using (SqlDataReader dr = cmd.ExecuteReader())
				{
					while (dr.Read())
					{
						string fileContent = dr.GetNullableString(0);
						if (fileContent != null)
						{
							string htmlResult = ConvertHtmlAndFile.SaveToFile(fileContent);

							try
							{
								Application app2 = new Application();
								Document sourceDocument = app2.Documents.Open(htmlResult);
								object start = sourceDocument.Content.Start;
								object end = sourceDocument.Content.End;
								Microsoft.Office.Interop.Word.Range myRange = sourceDocument.Range(ref start, ref end);
								myRange.Select();
								//myRange.set_Style(ref Normal);
								sourceDocument.Save();
								sourceDocument.Close(WdSaveOptions.wdSaveChanges);
								app2.Quit();

								rangeBookmark.InsertFile(htmlResult);
							}
							catch (Exception ex)
							{
								throw ex;
							}
							finally
							{
								File.Delete(htmlResult);
							}
						}
					}
				}
			}
		}


		protected void FillBookmarkWithCMAttachmentNormal(Application app, SqlConnection con, string bookmarkName, string tableName, long cmId)
        {
            this.Logger.Info(tableName);

			object Normal = "Normal";

			string query = "SELECT";
            query = query + " Attachment";
            query = query + " FROM " + tableName;
            query = query + " WHERE CMId = @CMId";
            query = query + " ORDER BY OrderNumber DESC";

            using (SqlCommand cmd = con.CreateCommand())
            {
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = query;
                cmd.Parameters.Add(this.NewSqlParameter("CMId", SqlDbType.BigInt, cmId));

                Range rangeBookmark = app.ActiveDocument.Bookmarks[bookmarkName].Range;
                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string fileContent = dr.GetNullableString(0);
						if (fileContent != null)
						{
							string htmlResult = ConvertHtmlAndFile.SaveToFile(fileContent);

							try
							{
								Application app2 = new Application();
								Document sourceDocument = app2.Documents.Open(htmlResult);
								object start = sourceDocument.Content.Start;
								object end = sourceDocument.Content.End;
								Microsoft.Office.Interop.Word.Range myRange = sourceDocument.Range(ref start, ref end);
								myRange.Select();
								//myRange.set_Style(ref Normal);
								sourceDocument.Save();
								sourceDocument.Close(WdSaveOptions.wdSaveChanges);
								app2.Quit();

								rangeBookmark.InsertFile(htmlResult);
							}
							catch (Exception ex)
							{
								throw ex;
							}
							finally
							{
								File.Delete(htmlResult);
							}
						}
					}
                }
            }
        }

		protected void FillBookmarkWithCMAttachmentABNormal(Application app, SqlConnection con, string bookmarkName, string tableName, long cmId, string columnName, string columnCMId)
		{
			this.Logger.Info(tableName);

			object Normal = "Normal";

			string query = "SELECT";
			query = query + " " + columnName;
			query = query + " FROM " + tableName;
			query = query + " WHERE " + columnCMId + " = @CMId";			

			using (SqlCommand cmd = con.CreateCommand())
			{
				cmd.CommandType = CommandType.Text;
				cmd.CommandText = query;
				cmd.Parameters.Add(this.NewSqlParameter("CMId", SqlDbType.BigInt, cmId));

				Range rangeBookmark = app.ActiveDocument.Bookmarks[bookmarkName].Range;
				using (SqlDataReader dr = cmd.ExecuteReader())
				{
					while (dr.Read())
					{
						string fileContent = dr.GetNullableString(0);
						if (fileContent != null)
						{
							string htmlResult = ConvertHtmlAndFile.SaveToFile(fileContent);

							try
							{
								Application app2 = new Application();
								Document sourceDocument = app2.Documents.Open(htmlResult);
								object start = sourceDocument.Content.Start;
								object end = sourceDocument.Content.End;
								Microsoft.Office.Interop.Word.Range myRange = sourceDocument.Range(ref start, ref end);
								myRange.Select();
								//myRange.set_Style(ref Normal);
								sourceDocument.Save();
								sourceDocument.Close(WdSaveOptions.wdSaveChanges);
								app2.Quit();

								rangeBookmark.InsertFile(htmlResult);
							}
							catch (Exception ex)
							{
								throw ex;
							}
							finally
							{
								File.Delete(htmlResult);
							}
						}
					}
				}
			}
		}

		protected void SetBookmarkText(Application app, string bookmarkName, string value)
        {
            Range rangeBookmark = app.ActiveDocument.Bookmarks[bookmarkName].Range;
            rangeBookmark.Font.Name = "Roboto Light";
            rangeBookmark.Font.Size = 10;
            rangeBookmark.Text = value;
        }
    }
}