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
		string version = "28Des_13:35PM";
        public void MergePAMDocument(long id, string connectionString, string folderTemplateLocation, string temporaryFolderLocation, string mergeByFQN, string mergeBy)
        {
			log4net.Config.XmlConfigurator.Configure();
			this.Logger.Info("MergePAMDocument_" + version);

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
                            PAM_ProjectFinanceNEW svcPAMProjectFinance = new PAM_ProjectFinanceNEW();
                            fileMergeResult = svcPAMProjectFinance.MergePAMProjectFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        case 2:
							PAM_CorporateFinanceNEW svcPAMCorporateFinance = new PAM_CorporateFinanceNEW();
                            fileMergeResult = svcPAMCorporateFinance.MergePAMCorporateFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        case 3:
                            PAM_EquityFinanceNEW svcEquityFinance = new PAM_EquityFinanceNEW();
                            fileMergeResult = svcEquityFinance.MergePAMEquityFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        default:
                            break;
                    }

					string query1 = string.Empty;
					query1 = query1 + "SELECT";
					query1 = query1 + " [MWorkflowStatusId]";
					query1 = query1 + " FROM [dbo].[PAM]";
					query1 = query1 + " WHERE [Id] = @Id";
					int? mWorkflowStatusId = null;
					using (SqlCommand cmd = con.CreateCommand())
					{
						cmd.CommandType = CommandType.Text;
						cmd.CommandText = query1;
						cmd.Parameters.Add(this.NewSqlParameter("Id", SqlDbType.BigInt, id));

						using (SqlDataReader dr = cmd.ExecuteReader())
						{
							int indexOf_mWorkflowStatusId = dr.GetOrdinal("MWorkflowStatusId");

							while (dr.Read())
							{
								mWorkflowStatusId = dr.GetInt32(indexOf_mWorkflowStatusId);
							}
						}
					}
					bool isPreview = false;
					try
					{
						if (mWorkflowStatusId != null && mWorkflowStatusId == 7)
							isPreview = true;
					}
					catch { }

					SaveMergeResultToDatabase svcSave = new SaveMergeResultToDatabase();
                    svcSave.SavePAMToDatabase(con, id, fileMergeResult.FileContent, mergeByFQN, mergeBy, fileMergeResult.FileName, isPreview);
                }
                con.Close();
            }
        }

        public void MergeCMDocument(long id, string connectionString, string folderTemplateLocation, string temporaryFolderLocation, string mergeByFQN, string mergeBy)
        {
			log4net.Config.XmlConfigurator.Configure();
			this.Logger.Info("MergeCMDocument_" + version);

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
                            CM_ProjectFinanceNEW svcCMProjectFinance = new CM_ProjectFinanceNEW();
                            fileMergeResult = svcCMProjectFinance.MergeCMProjectFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        case 2:
                            CM_CorporateFinanceNEW svcCMCorporateFinance = new CM_CorporateFinanceNEW();
                            fileMergeResult = svcCMCorporateFinance.MergeCMCorporateFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        case 3:
                            CM_EquityFinanceNEW svcCMEquityFinance = new CM_EquityFinanceNEW();
                            fileMergeResult = svcCMEquityFinance.MergeCMEquityFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        case 4:
                            CM_WaiverFinanceNEW svcCMWaiver = new CM_WaiverFinanceNEW();
                            fileMergeResult = svcCMWaiver.MergeCMWaiverFinance(con, id, folderTemplateLocation, temporaryFolderLocation);
                            break;
                        default:
                            break;
                    }

					string query1 = string.Empty;
					query1 = query1 + "SELECT";
					query1 = query1 + " [MWorkflowStatusId]";
					query1 = query1 + " FROM [dbo].[CM]";
					query1 = query1 + " WHERE [Id] = @Id";
					int? mWorkflowStatusId = null;
					using (SqlCommand cmd = con.CreateCommand())
					{
						cmd.CommandType = CommandType.Text;
						cmd.CommandText = query1;
						cmd.Parameters.Add(this.NewSqlParameter("Id", SqlDbType.BigInt, id));

						using (SqlDataReader dr = cmd.ExecuteReader())
						{
							int indexOf_mWorkflowStatusId = dr.GetOrdinal("MWorkflowStatusId");

							while (dr.Read())
							{
								mWorkflowStatusId = dr.GetInt32(indexOf_mWorkflowStatusId);
							}
						}
					}
					bool isPreview = false;
					try
					{
						if (mWorkflowStatusId != null && mWorkflowStatusId == 7)
							isPreview = true;
					}
					catch { }

					SaveMergeResultToDatabase svcSave = new SaveMergeResultToDatabase();
                    svcSave.SaveCMToDatabase(con, id, fileMergeResult.FileContent, mergeByFQN, mergeBy, fileMergeResult.FileName, isPreview);
                }
                con.Close();
            }
        }
    }
}
