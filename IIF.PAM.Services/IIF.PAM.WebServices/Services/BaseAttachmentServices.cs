using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using System.Web;
using System.Web.Configuration;
using System.Web.Services;
using System.Xml.Linq;

using IIF.PAM.Utilities;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class BaseAttachmentServices : BaseServices
    {
        protected void ParsePAMXDocument(XDocument xDoc, BasePAMAttachment data, string downloadMethodName)
        {
            data.FullFileName = xDoc.Root.Element("name").Value;

            string asmxUrl = this.AppConfig.WebServiceUrl;
            asmxUrl = asmxUrl.AppendUrlPath(typeof(IIFPAMAttachmentServices).Name + ".asmx");

            if (data.FullFileName == "scnull")
            {
                data.DownloadHyperlink = null;
            }
            else
            {
                data.DownloadHyperlink = "<hyperlink><link>" + asmxUrl.AppendUrlPath(downloadMethodName + "?id=" + data.Id.ToString()) + "</link><display>" + HttpUtility.HtmlEncode(data.FullFileName) + "</display></hyperlink>";
            }
        }

        protected void ParseCMXDocument(XDocument xDoc, BaseCMAttachment data, string downloadMethodName)
        {
            data.FullFileName = xDoc.Root.Element("name").Value;

            string asmxUrl = this.AppConfig.WebServiceUrl;
            asmxUrl = asmxUrl.AppendUrlPath(typeof(IIFCMAttachmentServices).Name + ".asmx");

            if (data.FullFileName == "scnull")
            {
                data.DownloadHyperlink = null;
            }
            else
            {
                data.DownloadHyperlink = "<hyperlink><link>" + asmxUrl.AppendUrlPath(downloadMethodName + "?id=" + data.Id.ToString()) + "</link><display>" + HttpUtility.HtmlEncode(data.FullFileName) + "</display></hyperlink>";
            }
        }

        protected List<AttachmentType> ListPAMAttachmentType1<AttachmentType>(long pamId, string tableName, string downloadMethodName)
            where AttachmentType : BasePAMAttachment, IAttachmentType1, new()
        {
            List<AttachmentType> result = new List<AttachmentType>();
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();

                string query = "SELECT Id, PAMId, Attachment, OrderNumber";
                query = query + " , CreatedByFQN, CreatedBy, CreatedOn";
                query = query + " , ModifiedByFQN, ModifiedBy, ModifiedOn";
                query = query + " FROM " + tableName;
                query = query + " WHERE PAMId = @PAMId";
                query = query + " ORDER BY Id";

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = query;
                    cmd.Parameters.Add(this.NewSqlParameter("PAMId", SqlDbType.BigInt, pamId));

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_Id = dr.GetOrdinal("Id");
                        int indexOf_PAMId = dr.GetOrdinal("PAMId");
                        int indexOf_Attachment = dr.GetOrdinal("Attachment");
                        int indexOf_OrderNumber = dr.GetOrdinal("OrderNumber");
                        int indexOf_CreatedByFQN = dr.GetOrdinal("CreatedByFQN");
                        int indexOf_CreatedBy = dr.GetOrdinal("CreatedBy");
                        int indexOf_CreatedOn = dr.GetOrdinal("CreatedOn");
                        int indexOf_ModifiedByFQN = dr.GetOrdinal("ModifiedByFQN");
                        int indexOf_ModifiedBy = dr.GetOrdinal("ModifiedBy");
                        int indexOf_ModifiedOn = dr.GetOrdinal("ModifiedOn");

                        while (dr.Read())
                        {
                            AttachmentType data = new AttachmentType();
                            data.Id = dr.GetInt64(indexOf_Id);
                            data.PAMId = dr.GetInt64(indexOf_PAMId);
                            data.OrderNumber = dr.GetInt32(indexOf_OrderNumber);

                            data.CreatedByFQN = dr.GetString(indexOf_CreatedByFQN);
                            data.CreatedBy = dr.GetString(indexOf_CreatedBy);
                            data.CreatedOn = dr.GetDateTime(indexOf_CreatedOn);
                            data.ModifiedByFQN = dr.GetString(indexOf_ModifiedByFQN);
                            data.ModifiedBy = dr.GetString(indexOf_ModifiedBy);
                            data.ModifiedOn = dr.GetDateTime(indexOf_ModifiedOn);

                            string attachmentValue = dr.GetString(indexOf_Attachment);

                            if (!string.IsNullOrEmpty(attachmentValue))
                            {
                                XDocument xDoc = new XDocument();
                                xDoc = XDocument.Parse(attachmentValue);
                                this.ParsePAMXDocument(xDoc, data, downloadMethodName);
                            }

                            result.Add(data);
                        }
                    }
                }
                con.Close();
            }
            return result;
        }

        protected List<AttachmentType> ListPAMAttachmentType2<AttachmentType>(long pamId, string tableName, string downloadMethodName)
            where AttachmentType : BasePAMAttachment, IAttachmentType2, new()
        {
            List<AttachmentType> result = new List<AttachmentType>();
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();

                string query = "SELECT Id, PAMId, Attachment, Description";
                query = query + " , CreatedByFQN, CreatedBy, CreatedOn";
                query = query + " , ModifiedByFQN, ModifiedBy, ModifiedOn";
                query = query + " FROM " + tableName;
                query = query + " WHERE PAMId = @PAMId";
                query = query + " ORDER BY Id";

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = query;
                    cmd.Parameters.Add(this.NewSqlParameter("PAMId", SqlDbType.BigInt, pamId));

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_Id = dr.GetOrdinal("Id");
                        int indexOf_PAMId = dr.GetOrdinal("PAMId");
                        int indexOf_Attachment = dr.GetOrdinal("Attachment");
                        int indexOf_Description = dr.GetOrdinal("Description");
                        int indexOf_CreatedByFQN = dr.GetOrdinal("CreatedByFQN");
                        int indexOf_CreatedBy = dr.GetOrdinal("CreatedBy");
                        int indexOf_CreatedOn = dr.GetOrdinal("CreatedOn");
                        int indexOf_ModifiedByFQN = dr.GetOrdinal("ModifiedByFQN");
                        int indexOf_ModifiedBy = dr.GetOrdinal("ModifiedBy");
                        int indexOf_ModifiedOn = dr.GetOrdinal("ModifiedOn");

                        while (dr.Read())
                        {
                            AttachmentType data = new AttachmentType();
                            data.Id = dr.GetInt64(indexOf_Id);
                            data.PAMId = dr.GetInt64(indexOf_PAMId);
                            data.Description = dr.GetNullableString(indexOf_Description);

                            data.CreatedByFQN = dr.GetString(indexOf_CreatedByFQN);
                            data.CreatedBy = dr.GetString(indexOf_CreatedBy);
                            data.CreatedOn = dr.GetDateTime(indexOf_CreatedOn);
                            data.ModifiedByFQN = dr.GetString(indexOf_ModifiedByFQN);
                            data.ModifiedBy = dr.GetString(indexOf_ModifiedBy);
                            data.ModifiedOn = dr.GetDateTime(indexOf_ModifiedOn);

                            string attachmentValue = dr.GetString(indexOf_Attachment);

                            if (!string.IsNullOrEmpty(attachmentValue))
                            {
                                XDocument xDoc = new XDocument();
                                xDoc = XDocument.Parse(attachmentValue);
                                this.ParsePAMXDocument(xDoc, data, downloadMethodName);
                            }
                            result.Add(data);
                        }
                    }
                }
                con.Close();
            }
            return result;
        }

        protected List<AttachmentType> ListPAMAttachmentType3<AttachmentType>(long pamId, int? mWorkflowStatusIdWhenAdded, int? roleIdWhenAdded, string snWhenAdded_NOT, string tableName, string downloadMethodName)
            where AttachmentType : BasePAMAttachment, IAttachmentType3, new()
        {
            List<AttachmentType> result = new List<AttachmentType>();
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();

                string query = "SELECT Id, PAMId, Attachment";
                query = query + " , MWorkflowStatusIdWhenAdded, SNWhenAdded, Description";
                query = query + " , CreatedByFQN, CreatedBy, CreatedOn";
                query = query + " , ModifiedByFQN, ModifiedBy, ModifiedOn";
                query = query + " FROM " + tableName;
                query = query + " WHERE PAMId = @PAMId";
                if (mWorkflowStatusIdWhenAdded.HasValue)
                {
                    query = query + " AND MWorkflowStatusIdWhenAdded = @MWorkflowStatusIdWhenAdded";
                }
                if (roleIdWhenAdded.HasValue)
                {
                    switch (roleIdWhenAdded.Value)
                    {
                        case 3:
                            query = query + " AND MWorkflowStatusIdWhenAdded IN (7)";
                            break;
                        case 6:
                            query = query + " AND MWorkflowStatusIdWhenAdded IN (8, 10, 11, 12, 13)";
                            break;
                        case 7:
                            query = query + " AND MWorkflowStatusIdWhenAdded IN (9)";
                            break;
                        default:
                            break;
                    }
                }
                query = query + " AND SNWhenAdded <> @SNWhenAdded";
                query = query + " ORDER BY Id";

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = query;
                    cmd.Parameters.Add(this.NewSqlParameter("PAMId", SqlDbType.BigInt, pamId));
                    if (mWorkflowStatusIdWhenAdded.HasValue)
                    {
                        cmd.Parameters.Add(this.NewSqlParameter("MWorkflowStatusIdWhenAdded", SqlDbType.Int, mWorkflowStatusIdWhenAdded));
                    }
                    cmd.Parameters.Add(this.NewSqlParameter("SNWhenAdded", SqlDbType.VarChar, snWhenAdded_NOT));

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_Id = dr.GetOrdinal("Id");
                        int indexOf_PAMId = dr.GetOrdinal("PAMId");
                        int indexOf_Attachment = dr.GetOrdinal("Attachment");
                        int indexOf_MWorkflowStatusIdWhenAdded = dr.GetOrdinal("MWorkflowStatusIdWhenAdded");
                        int indexOf_SNWhenAdded = dr.GetOrdinal("SNWhenAdded");
                        int indexOf_Description = dr.GetOrdinal("Description");
                        int indexOf_CreatedByFQN = dr.GetOrdinal("CreatedByFQN");
                        int indexOf_CreatedBy = dr.GetOrdinal("CreatedBy");
                        int indexOf_CreatedOn = dr.GetOrdinal("CreatedOn");
                        int indexOf_ModifiedByFQN = dr.GetOrdinal("ModifiedByFQN");
                        int indexOf_ModifiedBy = dr.GetOrdinal("ModifiedBy");
                        int indexOf_ModifiedOn = dr.GetOrdinal("ModifiedOn");

                        while (dr.Read())
                        {
                            AttachmentType data = new AttachmentType();
                            data.Id = dr.GetInt64(indexOf_Id);
                            data.PAMId = dr.GetInt64(indexOf_PAMId);
                            data.MWorkflowStatusIdWhenAdded = dr.GetInt32(indexOf_MWorkflowStatusIdWhenAdded);
                            data.SNWhenAdded = dr.GetNullableString(indexOf_SNWhenAdded);
                            data.Description = dr.GetNullableString(indexOf_Description);

                            data.CreatedByFQN = dr.GetString(indexOf_CreatedByFQN);
                            data.CreatedBy = dr.GetString(indexOf_CreatedBy);
                            data.CreatedOn = dr.GetDateTime(indexOf_CreatedOn);
                            data.ModifiedByFQN = dr.GetString(indexOf_ModifiedByFQN);
                            data.ModifiedBy = dr.GetString(indexOf_ModifiedBy);
                            data.ModifiedOn = dr.GetDateTime(indexOf_ModifiedOn);

                            string attachmentValue = dr.GetString(indexOf_Attachment);

                            if (!string.IsNullOrEmpty(attachmentValue))
                            {
                                XDocument xDoc = new XDocument();
                                xDoc = XDocument.Parse(attachmentValue);
                                this.ParsePAMXDocument(xDoc, data, downloadMethodName);
                            }

                            result.Add(data);
                        }
                    }
                }
                con.Close();
            }
            return result;
        }

        protected List<AttachmentType> ListCMAttachmentType1<AttachmentType>(long cmId, string tableName, string downloadMethodName)
            where AttachmentType : BaseCMAttachment, IAttachmentType1, new()
        {
            List<AttachmentType> result = new List<AttachmentType>();
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();

                string query = "SELECT Id, CMId, Attachment, OrderNumber";
                query = query + " , CreatedByFQN, CreatedBy, CreatedOn";
                query = query + " , ModifiedByFQN, ModifiedBy, ModifiedOn";
                query = query + " FROM " + tableName;
                query = query + " WHERE CMId = @CMId";
                query = query + " ORDER BY Id";

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = query;
                    cmd.Parameters.Add(this.NewSqlParameter("CMId", SqlDbType.BigInt, cmId));

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_Id = dr.GetOrdinal("Id");
                        int indexOf_CMId = dr.GetOrdinal("CMId");
                        int indexOf_Attachment = dr.GetOrdinal("Attachment");
                        int indexOf_OrderNumber = dr.GetOrdinal("OrderNumber");
                        int indexOf_CreatedByFQN = dr.GetOrdinal("CreatedByFQN");
                        int indexOf_CreatedBy = dr.GetOrdinal("CreatedBy");
                        int indexOf_CreatedOn = dr.GetOrdinal("CreatedOn");
                        int indexOf_ModifiedByFQN = dr.GetOrdinal("ModifiedByFQN");
                        int indexOf_ModifiedBy = dr.GetOrdinal("ModifiedBy");
                        int indexOf_ModifiedOn = dr.GetOrdinal("ModifiedOn");

                        while (dr.Read())
                        {
                            AttachmentType data = new AttachmentType();
                            data.Id = dr.GetInt64(indexOf_Id);
                            data.CMId = dr.GetInt64(indexOf_CMId);
                            data.OrderNumber = dr.GetInt32(indexOf_OrderNumber);

                            data.CreatedByFQN = dr.GetString(indexOf_CreatedByFQN);
                            data.CreatedBy = dr.GetString(indexOf_CreatedBy);
                            data.CreatedOn = dr.GetDateTime(indexOf_CreatedOn);
                            data.ModifiedByFQN = dr.GetString(indexOf_ModifiedByFQN);
                            data.ModifiedBy = dr.GetString(indexOf_ModifiedBy);
                            data.ModifiedOn = dr.GetDateTime(indexOf_ModifiedOn);

                            string attachmentValue = dr.GetString(indexOf_Attachment);

                            if (!string.IsNullOrEmpty(attachmentValue))
                            {
                                XDocument xDoc = new XDocument();
                                xDoc = XDocument.Parse(attachmentValue);
                                this.ParseCMXDocument(xDoc, data, downloadMethodName);
                            }

                            result.Add(data);
                        }
                    }
                }
                con.Close();
            }
            return result;
        }

        protected List<AttachmentType> ListCMAttachmentType3<AttachmentType>(long cmId, int? mWorkflowStatusIdWhenAdded, int? roleIdWhenAdded, string snWhenAdded_NOT, string tableName, string downloadMethodName)
            where AttachmentType : BaseCMAttachment, IAttachmentType3, new()
        {
            List<AttachmentType> result = new List<AttachmentType>();
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();

                string query = "SELECT Id, CMId, Attachment";
                query = query + " , MWorkflowStatusIdWhenAdded, SNWhenAdded, Description";
                query = query + " , CreatedByFQN, CreatedBy, CreatedOn";
                query = query + " , ModifiedByFQN, ModifiedBy, ModifiedOn";
                query = query + " FROM " + tableName;
                query = query + " WHERE CMId = @CMId";
                if (mWorkflowStatusIdWhenAdded.HasValue)
                {
                    query = query + " AND MWorkflowStatusIdWhenAdded = @MWorkflowStatusIdWhenAdded";
                }
                if (roleIdWhenAdded.HasValue)
                {
                    switch (roleIdWhenAdded.Value)
                    {
                        case 3:
                            query = query + " AND MWorkflowStatusIdWhenAdded IN (7)";
                            break;
                        case 6:
                            query = query + " AND MWorkflowStatusIdWhenAdded IN (8, 10, 11, 12, 13)";
                            break;
                        case 7:
                            query = query + " AND MWorkflowStatusIdWhenAdded IN (9)";
                            break;
                        default:
                            break;
                    }
                }
                query = query + " AND SNWhenAdded <> @SNWhenAdded";
                query = query + " ORDER BY Id";

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = query;
                    cmd.Parameters.Add(this.NewSqlParameter("CMId", SqlDbType.BigInt, cmId));
                    if (mWorkflowStatusIdWhenAdded.HasValue)
                    {
                        cmd.Parameters.Add(this.NewSqlParameter("MWorkflowStatusIdWhenAdded", SqlDbType.Int, mWorkflowStatusIdWhenAdded));
                    }
                    cmd.Parameters.Add(this.NewSqlParameter("SNWhenAdded", SqlDbType.VarChar, snWhenAdded_NOT));

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_Id = dr.GetOrdinal("Id");
                        int indexOf_CMId = dr.GetOrdinal("CMId");
                        int indexOf_Attachment = dr.GetOrdinal("Attachment");
                        int indexOf_MWorkflowStatusIdWhenAdded = dr.GetOrdinal("MWorkflowStatusIdWhenAdded");
                        int indexOf_SNWhenAdded = dr.GetOrdinal("SNWhenAdded");
                        int indexOf_Description = dr.GetOrdinal("Description");
                        int indexOf_CreatedByFQN = dr.GetOrdinal("CreatedByFQN");
                        int indexOf_CreatedBy = dr.GetOrdinal("CreatedBy");
                        int indexOf_CreatedOn = dr.GetOrdinal("CreatedOn");
                        int indexOf_ModifiedByFQN = dr.GetOrdinal("ModifiedByFQN");
                        int indexOf_ModifiedBy = dr.GetOrdinal("ModifiedBy");
                        int indexOf_ModifiedOn = dr.GetOrdinal("ModifiedOn");

                        while (dr.Read())
                        {
                            AttachmentType data = new AttachmentType();
                            data.Id = dr.GetInt64(indexOf_Id);
                            data.CMId = dr.GetInt64(indexOf_CMId);
                            data.MWorkflowStatusIdWhenAdded = dr.GetInt32(indexOf_MWorkflowStatusIdWhenAdded);
                            data.SNWhenAdded = dr.GetNullableString(indexOf_SNWhenAdded);
                            data.Description = dr.GetNullableString(indexOf_Description);

                            data.CreatedByFQN = dr.GetString(indexOf_CreatedByFQN);
                            data.CreatedBy = dr.GetString(indexOf_CreatedBy);
                            data.CreatedOn = dr.GetDateTime(indexOf_CreatedOn);
                            data.ModifiedByFQN = dr.GetString(indexOf_ModifiedByFQN);
                            data.ModifiedBy = dr.GetString(indexOf_ModifiedBy);
                            data.ModifiedOn = dr.GetDateTime(indexOf_ModifiedOn);

                            string attachmentValue = dr.GetString(indexOf_Attachment);

                            if (!string.IsNullOrEmpty(attachmentValue))
                            {
                                XDocument xDoc = new XDocument();
                                xDoc = XDocument.Parse(attachmentValue);
                                this.ParseCMXDocument(xDoc, data, downloadMethodName);
                            }

                            result.Add(data);
                        }
                    }
                }
                con.Close();
            }
            return result;
        }

        public XDocument GetAttachmentContentAllType(long id, string tableName)
        {
            XDocument result = null;
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();

                string query = "SELECT Attachment";
                query = query + " FROM " + tableName;
                query = query + " WHERE Id = @Id";

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = query;
                    cmd.Parameters.Add(this.NewSqlParameter("Id", SqlDbType.BigInt, id));

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_Attachment = dr.GetOrdinal("Attachment");
                        while (dr.Read())
                        {
                            string attachmentValue = dr.GetString(indexOf_Attachment);
                            if (!string.IsNullOrEmpty(attachmentValue))
                            {
                                result = XDocument.Parse(attachmentValue);
                            }
                        }
                    }
                }
                con.Close();
            }

            return result;
        }
    }
}