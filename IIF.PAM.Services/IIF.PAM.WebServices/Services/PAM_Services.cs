using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;

using IIF.PAM.Utilities;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_Services : BaseServices
    {
        public void DownloadPAMToSharedFolder(long id)
        {
            Dictionary<string, List<AttachmentWithContent>> dictAttachmentWithContent = new Dictionary<string, List<AttachmentWithContent>>();

            DateTime? pamDate = null;
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
                query = query + Environment.NewLine + "\t" + "\t" + ",[Src].[PAMDate]";
                query = query + Environment.NewLine + "\t" + "FROM [dbo].[PAM] AS [Src]";

                query = query + Environment.NewLine + "\t" + "WHERE [Src].[Id] = @Id";
                query = query + Environment.NewLine + ")";

                query = query + Environment.NewLine + "SELECT";
                query = query + Environment.NewLine + "\t" + "[Src].[Id]";
                query = query + Environment.NewLine + "\t" + ",[Src].[PAMDate]";
                query = query + Environment.NewLine + "\t" + ",[SrcProject].[ProjectCode]";
                query = query + Environment.NewLine + "\t" + ",[SrcCompany].[ProjectCompanyOrBorrowerCompanyOrTargetCompany]";
                query = query + Environment.NewLine + "FROM [CTE_Src] AS [Src]";

                query = query + Environment.NewLine + "INNER JOIN [dbo].[PAM_ProjectData] AS [SrcProject]";
                query = query + Environment.NewLine + "ON [Src].[Id] = [SrcProject].[Id]";

                query = query + Environment.NewLine + "INNER JOIN [dbo].[PAM_BorrowerOrTargetCompanyData] AS [SrcCompany]";
                query = query + Environment.NewLine + "ON [Src].[Id] = [SrcCompany].[Id]";

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = query;
                    cmd.Parameters.Add(this.NewSqlParameter("Id", SqlDbType.BigInt, id));

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_PAMDate = dr.GetOrdinal("PAMDate");
                        int indexOf_ProjectCode = dr.GetOrdinal("ProjectCode");
                        int indexOf_ProjectCompanyOrBorrowerCompanyOrTargetCompany = dr.GetOrdinal("ProjectCompanyOrBorrowerCompanyOrTargetCompany");

                        while (dr.Read())
                        {
                            pamDate = dr.GetNullableDateTime(indexOf_PAMDate);
                            projectCode = dr.GetString(indexOf_ProjectCode);
                            companyName = dr.GetString(indexOf_ProjectCompanyOrBorrowerCompanyOrTargetCompany);
                        }
                    }
                }

                query = "";
                query = query + Environment.NewLine + " SELECT ";
                query = query + Environment.NewLine + "\t" + " 1 AS [AttachmentType]";
                query = query + Environment.NewLine + "\t" + " ,[Attachment]";
                query = query + Environment.NewLine + "\t" + " ,NULL AS [Description]";
                query = query + Environment.NewLine + " FROM [dbo].[PAM_MergedDocumentResult]";
                query = query + Environment.NewLine + " WHERE [PAMId] = @PAMId";
                query = query + Environment.NewLine + " AND [IsPreview] = 0";
                query = query + Environment.NewLine + " AND [IsForHistory] = 0";

                query = query + Environment.NewLine + " UNION ";
                query = query + Environment.NewLine + " SELECT ";
                query = query + Environment.NewLine + "\t" + " 2 AS [AttachmentType]";
                query = query + Environment.NewLine + "\t" + " ,[Attachment]";
                query = query + Environment.NewLine + "\t" + " ,[Description]";
                query = query + Environment.NewLine + " FROM [dbo].[PAM_LegalDueDiligenceReport]";
                query = query + Environment.NewLine + " WHERE [PAMId] = @PAMId";

                query = query + Environment.NewLine + " UNION ";
                query = query + Environment.NewLine + " SELECT ";
                query = query + Environment.NewLine + "\t" + " 3 AS [AttachmentType]";
                query = query + Environment.NewLine + "\t" + " ,[Attachment]";
                query = query + Environment.NewLine + "\t" + " ,[Description]";
                query = query + Environment.NewLine + " FROM [dbo].[PAM_SAndEDueDiligence]";
                query = query + Environment.NewLine + " WHERE [PAMId] = @PAMId";

                query = query + Environment.NewLine + " UNION ";
                query = query + Environment.NewLine + " SELECT ";
                query = query + Environment.NewLine + "\t" + " 4 AS [AttachmentType]";
                query = query + Environment.NewLine + "\t" + " ,[Attachment]";
                query = query + Environment.NewLine + "\t" + " ,[Description]";
                query = query + Environment.NewLine + " FROM [dbo].[PAM_OtherReports]";
                query = query + Environment.NewLine + " WHERE [PAMId] = @PAMId";

                query = query + Environment.NewLine + " UNION ";
                query = query + Environment.NewLine + " SELECT ";
                query = query + Environment.NewLine + "\t" + " 5 AS [AttachmentType]";
                query = query + Environment.NewLine + "\t" + " ,[Attachment]";
                query = query + Environment.NewLine + "\t" + " ,[Description]";
                query = query + Environment.NewLine + " FROM [dbo].[PAM_OtherSupportingDocument]";
                query = query + Environment.NewLine + " WHERE [PAMId] = @PAMId";

                query = query + Environment.NewLine + " UNION ";
                query = query + Environment.NewLine + " SELECT ";
                query = query + Environment.NewLine + "\t" + " 6 AS [AttachmentType]";
                query = query + Environment.NewLine + "\t" + " ,[Attachment]";
                query = query + Environment.NewLine + "\t" + " ,[Description]";
                query = query + Environment.NewLine + " FROM [dbo].[PAM_DocumentRelatedToBoDDecision]";
                query = query + Environment.NewLine + " WHERE [PAMId] = @PAMId";

                query = query + Environment.NewLine + " UNION ";
                query = query + Environment.NewLine + " SELECT ";
                query = query + Environment.NewLine + "\t" + " 7 AS [AttachmentType]";
                query = query + Environment.NewLine + "\t" + " ,[Attachment]";
                query = query + Environment.NewLine + "\t" + " ,[Description]";
                query = query + Environment.NewLine + " FROM [dbo].[PAM_DocumentRelatedToBoCDecision]";
                query = query + Environment.NewLine + " WHERE [PAMId] = @PAMId";

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = query;
                    cmd.Parameters.Add(this.NewSqlParameter("PAMId", SqlDbType.BigInt, id));

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

            string folderPath = this.AppConfig.DMSSharedFolderLocation.AppendFolderPath(companyName + "_" + projectCode);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            foreach (KeyValuePair<string, List<AttachmentWithContent>> kvPair in dictAttachmentWithContent)
            {
                string docTypeFolder = folderPath;
                if (kvPair.Value.Count > 1)
                {
                    if (pamDate.HasValue)
                    {
                        docTypeFolder = docTypeFolder.AppendFolderPath(kvPair.Key + "_" + pamDate.Value.ToString("dd MMM yyyy"));
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