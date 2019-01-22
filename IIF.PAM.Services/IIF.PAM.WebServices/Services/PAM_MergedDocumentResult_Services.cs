using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_MergedDocumentResult_Services : BaseAttachmentServices
    {
        public List<PAM_MergedDocumentResultAttachment> ListMergedDocumentResultAttachment(long pamId, bool isPreview)
        {
            List<PAM_MergedDocumentResultAttachment> result = new List<PAM_MergedDocumentResultAttachment>();
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();

                string query = "SELECT Id, PAMId, Attachment, IsForHistory, IsPreview";
                query = query + " , CreatedByFQN, CreatedBy, CreatedOn";
                query = query + " , ModifiedByFQN, ModifiedBy, ModifiedOn";
                query = query + " FROM [dbo].[PAM_MergedDocumentResult]";
                query = query + " WHERE PAMId = @PAMId";
                query = query + " AND IsPreview = @IsPreview";
                query = query + " ORDER BY ModifiedOn";

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = query;
                    cmd.Parameters.Add(this.NewSqlParameter("PAMId", SqlDbType.BigInt, pamId));
                    cmd.Parameters.Add(this.NewSqlParameter("IsPreview", SqlDbType.Bit, isPreview));

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_Id = dr.GetOrdinal("Id");
                        int indexOf_PAMId = dr.GetOrdinal("PAMId");
                        int indexOf_Attachment = dr.GetOrdinal("Attachment");
                        int indexOf_IsForHistory = dr.GetOrdinal("IsForHistory");
                        int indexOf_IsPreview = dr.GetOrdinal("IsPreview");
                        int indexOf_CreatedByFQN = dr.GetOrdinal("CreatedByFQN");
                        int indexOf_CreatedBy = dr.GetOrdinal("CreatedBy");
                        int indexOf_CreatedOn = dr.GetOrdinal("CreatedOn");
                        int indexOf_ModifiedByFQN = dr.GetOrdinal("ModifiedByFQN");
                        int indexOf_ModifiedBy = dr.GetOrdinal("ModifiedBy");
                        int indexOf_ModifiedOn = dr.GetOrdinal("ModifiedOn");

                        while (dr.Read())
                        {
                            PAM_MergedDocumentResultAttachment data = new PAM_MergedDocumentResultAttachment();
                            data.Id = dr.GetInt64(indexOf_Id);
                            data.PAMId = dr.GetInt64(indexOf_PAMId);
                            data.IsForHistory = dr.GetBoolean(indexOf_IsForHistory);
                            data.IsPreview = dr.GetBoolean(indexOf_IsPreview);

                            data.CreatedByFQN = dr.GetString(indexOf_CreatedByFQN);
                            data.CreatedBy = dr.GetString(indexOf_CreatedBy);
                            data.CreatedOn = dr.GetDateTime(indexOf_CreatedOn);
                            data.ModifiedByFQN = dr.GetString(indexOf_ModifiedByFQN);
                            data.ModifiedBy = dr.GetString(indexOf_ModifiedBy);
                            data.ModifiedOn = dr.GetDateTime(indexOf_ModifiedOn);

                            XDocument xDoc = new XDocument();
                            xDoc = XDocument.Parse(dr.GetString(indexOf_Attachment));
                            this.ParsePAMXDocument(xDoc, data, "Get_PAM_MergedDocumentResultAttachment_Content");

                            result.Add(data);
                        }
                    }
                }
                con.Close();
            }
            return result;
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_MergedDocumentResult]");
        }
    }
}