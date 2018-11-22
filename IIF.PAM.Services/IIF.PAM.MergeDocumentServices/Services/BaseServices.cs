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

        protected void FillBookmarkWithPAMAttachmentType1(Application app, SqlConnection con, string bookmarkName, string tableName, long pamId)
        {
            this.Logger.Info(tableName);            

            string query = "SELECT";
            query = query + " Attachment";
            query = query + " FROM " + tableName;
            query = query + " WHERE PAMId = @PAMId";
            query = query + " ORDER BY OrderNumber";

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
                            rangeBookmark.InsertFile(htmlResult);
                            rangeBookmark.Font.Name = "Roboto Light";
                        }
                    }
                }
            }            
        }

        protected void FillBookmarkWithPAMAttachmentType2(Application app, SqlConnection con, string bookmarkName, string tableName, long pamId)
        {
            this.Logger.Info(tableName);            

            string query = "SELECT";
            query = query + " Description";
            query = query + " FROM " + tableName;
            query = query + " WHERE PAMId = @PAMId";
            query = query + " ORDER BY Id";

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
                        string description = dr.GetNullableString(0);
                        rangeBookmark.InsertBefore(description);
                        rangeBookmark.Font.Name = "Roboto Light";
                    }
                }
            }
        }

        protected void FillBookmarkWithCMAttachmentType1(Application app, SqlConnection con, string bookmarkName, string tableName, long cmId)
        {
            this.Logger.Info(tableName);            

            string query = "SELECT";
            query = query + " Attachment";
            query = query + " FROM " + tableName;
            query = query + " WHERE CMId = @CMId";
            query = query + " ORDER BY OrderNumber";

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
                            rangeBookmark.InsertFile(htmlResult);
                            rangeBookmark.Font.Name = "Roboto Light";
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