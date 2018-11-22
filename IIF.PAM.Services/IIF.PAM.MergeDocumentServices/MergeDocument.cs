using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

using IIF.PAM.MergeDocumentServices.Models;
using IIF.PAM.MergeDocumentServices.Services;

namespace IIF.PAM.MergeDocumentServices
{
    public class MergeDocument : BaseServices
    {
        public void MergePAMDocument(long id, string connectionString, string folderTemplateLocation, string temporaryFolderLocation, string mergeByFQN, string mergeBy)
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                string query = string.Empty;
                query = query + "SELECT";
                query = query + " [MProductTypeId]";
                query = query + " FROM [dbo].[PAM]";
                query = query + " WHERE [Id] = @Id";
                int? productTypeId = null;
                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = query;
                    cmd.Parameters.Add(this.NewSqlParameter("Id", SqlDbType.BigInt, id));

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_MProductTypeId = dr.GetOrdinal("MProductTypeId");

                        while (dr.Read())
                        {
                            productTypeId = dr.GetInt32(indexOf_MProductTypeId);
                        }
                    }
                }
                if (productTypeId.HasValue)
                {
                    FileMergeResult fileMergeResult = null;                    
                    switch (productTypeId.Value)
                    {
                        case 1:
                            PAM_ProjectFinance svcPAMProjectFinance = new PAM_ProjectFinance();
                            fileMergeResult = svcPAMProjectFinance.MergePAMProjectFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        case 2:
                            PAM_CorporateFinance svcPAMCorporateFinance = new PAM_CorporateFinance();
                            fileMergeResult = svcPAMCorporateFinance.MergePAMCorporateFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        case 3:
                            PAM_EquityFinance svcEquityFinance = new PAM_EquityFinance();
                            fileMergeResult = svcEquityFinance.MergePAMEquityFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        default:
                            break;
                    }
                    
                    SaveMergeResultToDatabase svcSave = new SaveMergeResultToDatabase();
                    svcSave.SavePAMToDatabase(con, id, fileMergeResult.FileContent, mergeByFQN, mergeBy, fileMergeResult.FileName);
                }
                con.Close();
            }
        }

        public void MergeCMDocument(long id, string connectionString, string folderTemplateLocation, string temporaryFolderLocation, string mergeByFQN, string mergeBy)
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                string query = string.Empty;
                query = query + "SELECT";
                query = query + " [MProductTypeId]";
                query = query + " FROM [dbo].[CM]";
                query = query + " WHERE [Id] = @Id";
                int? productTypeId = null;
                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = query;
                    cmd.Parameters.Add(this.NewSqlParameter("Id", SqlDbType.BigInt, id));

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_MProductTypeId = dr.GetOrdinal("MProductTypeId");

                        while (dr.Read())
                        {
                            productTypeId = dr.GetInt32(indexOf_MProductTypeId);
                        }
                    }
                }
                if (productTypeId.HasValue)
                {
                    FileMergeResult fileMergeResult = null;
                    switch (productTypeId.Value)
                    {
                        case 1:
                            CM_ProjectFinance svcCMProjectFinance = new CM_ProjectFinance();
                            fileMergeResult = svcCMProjectFinance.MergeCMProjectFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        case 2:
                            CM_CorporateFinance svcCMCorporateFinance = new CM_CorporateFinance();
                            fileMergeResult = svcCMCorporateFinance.MergeCMCorporateFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        case 3:
                            CM_EquityFinance svcCMEquityFinance = new CM_EquityFinance();
                            fileMergeResult = svcCMEquityFinance.MergeCMEquityFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        case 4:
                            CM_WaiverFinance svcCMWaiver = new CM_WaiverFinance();
                            fileMergeResult = svcCMWaiver.MergeCMWaiverFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        default:
                            break;
                    }

                    SaveMergeResultToDatabase svcSave = new SaveMergeResultToDatabase();
                    svcSave.SaveCMToDatabase(con, id, fileMergeResult.FileContent, mergeByFQN, mergeBy, fileMergeResult.FileName);
                }
                con.Close();
            }
        }
    }
}
