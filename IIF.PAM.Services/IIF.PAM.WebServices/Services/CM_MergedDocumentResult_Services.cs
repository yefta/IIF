using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Web;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_MergedDocumentResult_Services : BaseAttachmentServices
    {
        public List<CM_MergedDocumentResultAttachment> ListMergedDocumentResultAttachment(long cmId)
        {
            List<CM_MergedDocumentResultAttachment> result = new List<CM_MergedDocumentResultAttachment>();
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();

                string query = "SELECT Id, CMId, Attachment, IsForHistory";
                query = query + " , CreatedByFQN, CreatedBy, CreatedOn";
                query = query + " , ModifiedByFQN, ModifiedBy, ModifiedOn";
                query = query + " FROM [dbo].[CM_MergedDocumentResult]";
                query = query + " WHERE CMId = @CMId";
                query = query + " ORDER BY ModifiedOn";

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
                        int indexOf_IsForHistory = dr.GetOrdinal("IsForHistory");
                        int indexOf_CreatedByFQN = dr.GetOrdinal("CreatedByFQN");
                        int indexOf_CreatedBy = dr.GetOrdinal("CreatedBy");
                        int indexOf_CreatedOn = dr.GetOrdinal("CreatedOn");
                        int indexOf_ModifiedByFQN = dr.GetOrdinal("ModifiedByFQN");
                        int indexOf_ModifiedBy = dr.GetOrdinal("ModifiedBy");
                        int indexOf_ModifiedOn = dr.GetOrdinal("ModifiedOn");

                        while (dr.Read())
                        {
                            CM_MergedDocumentResultAttachment data = new CM_MergedDocumentResultAttachment();
                            data.Id = dr.GetInt64(indexOf_Id);
                            data.CMId = dr.GetInt64(indexOf_CMId);
                            data.IsForHistory = dr.GetBoolean(indexOf_IsForHistory);

                            data.CreatedByFQN = dr.GetString(indexOf_CreatedByFQN);
                            data.CreatedBy = dr.GetString(indexOf_CreatedBy);
                            data.CreatedOn = dr.GetDateTime(indexOf_CreatedOn);
                            data.ModifiedByFQN = dr.GetString(indexOf_ModifiedByFQN);
                            data.ModifiedBy = dr.GetString(indexOf_ModifiedBy);
                            data.ModifiedOn = dr.GetDateTime(indexOf_ModifiedOn);

                            XDocument xDoc = new XDocument();
                            xDoc = XDocument.Parse(dr.GetString(indexOf_Attachment));
                            this.ParseCMXDocument(xDoc, data, "Get_CM_MergedDocumentResultAttachment_Content");

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
            return this.GetAttachmentContentAllType(id, "[dbo].[CM_MergedDocumentResult]");
        }
    }
}