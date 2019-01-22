using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;

using IIF.PAM.Utilities;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_Services : BaseServices
    {
        public void DownloadCMToSharedFolder(long id)
        {
            Dictionary<string, List<AttachmentWithContent>> dictAttachmentWithContent = new Dictionary<string, List<AttachmentWithContent>>();

            string cmNumber = string.Empty;
            DateTime? CMDate = null;
            string projectCode = string.Empty;
            string companyName = string.Empty;

            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();

                string query = "";
                query = query + Environment.NewLine + "WITH [CTE_Src] AS ( ";
                query = query + Environment.NewLine + "\t" + "SELECT";
                query = query + Environment.NewLine + "\t" + "\t" + "[Src].[Id]";
                query = query + Environment.NewLine + "\t" + "\t" + ",[Src].[CMNumber]";
                query = query + Environment.NewLine + "\t" + "\t" + ",[Src].[CMDate]";
                query = query + Environment.NewLine + "\t" + "FROM [dbo].[CM] AS [Src]";

                query = query + Environment.NewLine + "\t" + "WHERE [Src].[Id] = @Id";
                query = query + Environment.NewLine + ")";

                query = query + Environment.NewLine + "SELECT";
                query = query + Environment.NewLine + "\t" + "[Src].[Id]";
                query = query + Environment.NewLine + "\t" + ",[Src].[CMNumber]";
                query = query + Environment.NewLine + "\t" + ",[Src].[CMDate]";
                query = query + Environment.NewLine + "\t" + ",[SrcProject].[ProjectCode]";
                query = query + Environment.NewLine + "\t" + ",[SrcCompany].[ProjectCompanyOrInvesteeOrBorrower]";
                query = query + Environment.NewLine + "FROM [CTE_Src] AS [Src]";

                query = query + Environment.NewLine + "INNER JOIN [dbo].[CM_ProjectData] AS [SrcProject]";
                query = query + Environment.NewLine + "ON [Src].[Id] = [SrcProject].[Id]";

                query = query + Environment.NewLine + "INNER JOIN [dbo].[CM_BorrowerOrInvesteeCompanyData] AS [SrcCompany]";
                query = query + Environment.NewLine + "ON [Src].[Id] = [SrcCompany].[Id]";

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = query;
                    cmd.Parameters.Add(this.NewSqlParameter("Id", SqlDbType.BigInt, id));

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_CMNumber = dr.GetOrdinal("CMNumber");
                        int indexOf_CMDate = dr.GetOrdinal("CMDate");
                        int indexOf_ProjectCode = dr.GetOrdinal("ProjectCode");
                        int indexOf_ProjectCompanyOrInvesteeOrBorrower = dr.GetOrdinal("ProjectCompanyOrInvesteeOrBorrower");

                        while (dr.Read())
                        {
                            cmNumber = dr.GetNullableString(indexOf_CMNumber);
                            CMDate = dr.GetNullableDateTime(indexOf_CMDate);
                            projectCode = dr.GetString(indexOf_ProjectCode);
                            companyName = dr.GetString(indexOf_ProjectCompanyOrInvesteeOrBorrower);
                        }
                    }
                }

                query = "";
                query = query + Environment.NewLine + " SELECT ";
                query = query + Environment.NewLine + "\t" + " 8 AS [AttachmentType]";
                query = query + Environment.NewLine + "\t" + " ,[Attachment]";
                query = query + Environment.NewLine + "\t" + " ,NULL AS [Description]";
                query = query + Environment.NewLine + " FROM [dbo].[CM_MergedDocumentResult]";
                query = query + Environment.NewLine + " WHERE [CMId] = @CMId";
                query = query + Environment.NewLine + " AND [IsPreview] = 0";
                query = query + Environment.NewLine + " AND [IsForHistory] = 0";

                query = query + Environment.NewLine + " UNION ";
                query = query + Environment.NewLine + " SELECT ";
                query = query + Environment.NewLine + "\t" + " 9 AS [AttachmentType]";
                query = query + Environment.NewLine + "\t" + " ,[Attachment]";
                query = query + Environment.NewLine + "\t" + " ,[Description]";
                query = query + Environment.NewLine + " FROM [dbo].[CM_OtherSupportingDocument]";
                query = query + Environment.NewLine + " WHERE [CMId] = @CMId";

                query = query + Environment.NewLine + " UNION ";
                query = query + Environment.NewLine + " SELECT ";
                query = query + Environment.NewLine + "\t" + " 10 AS [AttachmentType]";
                query = query + Environment.NewLine + "\t" + " ,[Attachment]";
                query = query + Environment.NewLine + "\t" + " ,[Description]";
                query = query + Environment.NewLine + " FROM [dbo].[CM_DocumentRelatedToBoDDecision]";
                query = query + Environment.NewLine + " WHERE [CMId] = @CMId";

                query = query + Environment.NewLine + " UNION ";
                query = query + Environment.NewLine + " SELECT ";
                query = query + Environment.NewLine + "\t" + " 11 AS [AttachmentType]";
                query = query + Environment.NewLine + "\t" + " ,[Attachment]";
                query = query + Environment.NewLine + "\t" + " ,[Description]";
                query = query + Environment.NewLine + " FROM [dbo].[CM_DocumentRelatedToBoCDecision]";
                query = query + Environment.NewLine + " WHERE [CMId] = @CMId";

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = query;
                    cmd.Parameters.Add(this.NewSqlParameter("CMId", SqlDbType.BigInt, id));

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_AttachmentType = dr.GetOrdinal("AttachmentType");
                        int indexOf_Attachment = dr.GetOrdinal("Attachment");
                        int indexOf_Description = dr.GetOrdinal("Description");

                        while (dr.Read())
                        {
                            AttachmentWithContent data = new AttachmentWithContent();
                            data.AttachmentType = (AttachmentTypeConstants)dr.GetInt32(indexOf_AttachmentType);
                            data.Attachment = dr.GetNullableString(indexOf_Attachment);
                            data.Description = dr.GetNullableString(indexOf_Description);
                            data.ParseAttachment();
                            if (!string.IsNullOrEmpty(data.FileName))
                            {
                                if (!dictAttachmentWithContent.ContainsKey(data.AttachmentTypeDisplayName))
                                {
                                    dictAttachmentWithContent.Add(data.AttachmentTypeDisplayName, new List<AttachmentWithContent>());
                                }
                                dictAttachmentWithContent[data.AttachmentTypeDisplayName].Add(data);
                            }
                        }
                    }
                }
                con.Close();
            }

            if (string.IsNullOrEmpty(cmNumber))
            {
                cmNumber = "1";
            }
            string folderPath = this.AppConfig.DMSSharedFolderLocation.AppendFolderPath(companyName + "_" + projectCode + "_" + cmNumber);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            foreach (KeyValuePair<string, List<AttachmentWithContent>> kvPair in dictAttachmentWithContent)
            {
                string docTypeFolder = folderPath;
                if (kvPair.Value.Count > 1)
                {
                    if (CMDate.HasValue)
                    {
                        docTypeFolder = docTypeFolder.AppendFolderPath(kvPair.Key + "_" + CMDate.Value.ToString("dd MMM yyyy"));
                    }
                    else
                    {
                        docTypeFolder = docTypeFolder.AppendFolderPath(kvPair.Key);
                    }
                }
                if (!Directory.Exists(docTypeFolder))
                {
                    Directory.CreateDirectory(docTypeFolder);
                }
                foreach (AttachmentWithContent attachmentWithContent in kvPair.Value)
                {
                    string fileName = docTypeFolder.AppendFolderPath(attachmentWithContent.FileName);
                    string metadataFileName = fileName + ".metadata.txt";
                    File.WriteAllBytes(fileName, attachmentWithContent.FileContent);

                    using (StreamWriter fileMetadata = new StreamWriter(metadataFileName))
                    {
                        fileMetadata.WriteLine("Customer Name=" + companyName);
                        fileMetadata.WriteLine("Document Type=" + attachmentWithContent.AttachmentTypeDMSMetadataDisplayName);
                        fileMetadata.WriteLine("Description=" + attachmentWithContent.Description);
                    }
                }
            }
        }
    }
}