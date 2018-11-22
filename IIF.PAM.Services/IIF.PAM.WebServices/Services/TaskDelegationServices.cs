using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

using SourceCode.Workflow.Client;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class TaskDelegationServices : BaseServices
    {
        public void StartK2OOF()
        {
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();
                string querySelect = "SELECT";
                querySelect = querySelect + " [Src].[Id], [Src].[FromDate], [Src].[ToDate]";
                querySelect = querySelect + " , [Src].[FromFQN], [Src].[ToFQN]";
                querySelect = querySelect + " , [Src].[IsActive], [Src].[IsCanceled], [Src].[IsExpired]";
                querySelect = querySelect + " , [Src].[IsStartedInK2], [Src].[IsEndedInK2]";
                querySelect = querySelect + " FROM [dbo].[Vw_TaskDelegation_NeedK2Started] AS [Src]";

                List<TaskDelegation> listTaskDelegation = new List<TaskDelegation>();

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = querySelect;

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_Id = dr.GetOrdinal("Id");
                        int indexOf_FromDate = dr.GetOrdinal("FromDate");
                        int indexOf_ToDate = dr.GetOrdinal("ToDate");
                        int indexOf_FromFQN = dr.GetOrdinal("FromFQN");
                        int indexOf_ToFQN = dr.GetOrdinal("ToFQN");
                        int indexOf_IsActive = dr.GetOrdinal("IsActive");
                        int indexOf_IsCanceled = dr.GetOrdinal("IsCanceled");
                        int indexOf_IsExpired = dr.GetOrdinal("IsExpired");
                        int indexOf_IsStartedInK2 = dr.GetOrdinal("IsStartedInK2");
                        int indexOf_IsEndedInK2 = dr.GetOrdinal("IsEndedInK2");

                        while (dr.Read())
                        {
                            TaskDelegation data = new TaskDelegation();
                            data.Id = dr.GetInt64(indexOf_Id);
                            data.FromDate = dr.GetDateTime(indexOf_FromDate);
                            data.ToDate = dr.GetDateTime(indexOf_ToDate);
                            data.FromFQN = dr.GetString(indexOf_FromFQN);
                            data.ToFQN = dr.GetString(indexOf_ToFQN);
                            data.IsActive = dr.GetBoolean(indexOf_IsActive);
                            data.IsCanceled = dr.GetBoolean(indexOf_IsCanceled);
                            data.IsExpired = dr.GetBoolean(indexOf_IsExpired);
                            data.IsStartedInK2 = dr.GetBoolean(indexOf_IsStartedInK2);
                            data.IsEndedInK2 = dr.GetBoolean(indexOf_IsEndedInK2);

                            listTaskDelegation.Add(data);
                        }
                    }
                }

                Exception innerEx = null;

                foreach (TaskDelegation taskDelegation in listTaskDelegation)
                {
                    try
                    {
                        using (Connection k2Con = new Connection())
                        {
                            k2Con.Open(this.AppConfig.K2Server);
                            k2Con.ImpersonateUser(taskDelegation.FromFQN);

                            WorklistCriteria worklistCriteria = new WorklistCriteria();
                            worklistCriteria.Platform = "ASP";
                            Destinations worktypeDestinations = new Destinations();
                            worktypeDestinations.Add(new Destination(taskDelegation.ToFQN, DestinationType.User));
                            WorkType workType = new WorkType("TaskDelegationWork_" + taskDelegation.Id.ToString(), worklistCriteria, worktypeDestinations);

                            WorklistShare worklistShare = new WorklistShare();
                            worklistShare.ShareType = ShareType.OOF;

                            worklistShare.StartDate = taskDelegation.FromDate;
                            worklistShare.EndDate = taskDelegation.ToDate;

                            worklistShare.WorkTypes.Add(workType);
                            k2Con.ShareWorkList(worklistShare);

                            k2Con.SetUserStatus(UserStatuses.OOF);

                            string queryUpdate = "UPDATE [dbo].[TaskDelegation] SET";
                            queryUpdate = queryUpdate + " [IsActive] = 1";
                            queryUpdate = queryUpdate + ", [IsExpired] = 0";
                            queryUpdate = queryUpdate + ", [IsStartedInK2] = 1";
                            queryUpdate = queryUpdate + " WHERE [Id] = @Id";

                            using (SqlCommand cmd = con.CreateCommand())
                            {
                                cmd.CommandType = CommandType.Text;
                                cmd.CommandText = queryUpdate;
                                cmd.Parameters.Add(this.NewSqlParameter("Id", SqlDbType.BigInt, taskDelegation.Id));

                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        innerEx = ex;
                    }
                }

                if (innerEx != null)
                {
                    throw new Exception("An error occured. " + innerEx.Message, innerEx);
                }

                con.Close();
            }
        }

        public void EndK2OOF()
        {
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();
                string querySelect = "SELECT";
                querySelect = querySelect + " [Src].[Id], [Src].[FromDate], [Src].[ToDate]";
                querySelect = querySelect + " , [Src].[FromFQN], [Src].[ToFQN]";
                querySelect = querySelect + " , [Src].[IsActive], [Src].[IsCanceled], [Src].[IsExpired]";
                querySelect = querySelect + " , [Src].[IsStartedInK2], [Src].[IsEndedInK2]";
                querySelect = querySelect + " FROM [dbo].[Vw_TaskDelegation_NeedK2Ended] AS [Src]";

                List<TaskDelegation> listTaskDelegation = new List<TaskDelegation>();

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = querySelect;

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_Id = dr.GetOrdinal("Id");
                        int indexOf_FromDate = dr.GetOrdinal("FromDate");
                        int indexOf_ToDate = dr.GetOrdinal("ToDate");
                        int indexOf_FromFQN = dr.GetOrdinal("FromFQN");
                        int indexOf_ToFQN = dr.GetOrdinal("ToFQN");
                        int indexOf_IsActive = dr.GetOrdinal("IsActive");
                        int indexOf_IsCanceled = dr.GetOrdinal("IsCanceled");
                        int indexOf_IsExpired = dr.GetOrdinal("IsExpired");
                        int indexOf_IsStartedInK2 = dr.GetOrdinal("IsStartedInK2");
                        int indexOf_IsEndedInK2 = dr.GetOrdinal("IsEndedInK2");

                        while (dr.Read())
                        {
                            TaskDelegation data = new TaskDelegation();
                            data.Id = dr.GetInt64(indexOf_Id);
                            data.FromDate = dr.GetDateTime(indexOf_FromDate);
                            data.ToDate = dr.GetDateTime(indexOf_ToDate);
                            data.FromFQN = dr.GetString(indexOf_FromFQN);
                            data.ToFQN = dr.GetString(indexOf_ToFQN);
                            data.IsActive = dr.GetBoolean(indexOf_IsActive);
                            data.IsCanceled = dr.GetBoolean(indexOf_IsCanceled);
                            data.IsExpired = dr.GetBoolean(indexOf_IsExpired);
                            data.IsStartedInK2 = dr.GetBoolean(indexOf_IsStartedInK2);
                            data.IsEndedInK2 = dr.GetBoolean(indexOf_IsEndedInK2);

                            listTaskDelegation.Add(data);
                        }
                    }
                }

                Exception innerEx = null;

                foreach (TaskDelegation taskDelegation in listTaskDelegation)
                {
                    try
                    {
                        using (Connection k2Con = new Connection())
                        {
                            k2Con.Open(this.AppConfig.K2Server);
                            k2Con.ImpersonateUser(taskDelegation.FromFQN);

                            WorklistShares worklistShares = k2Con.GetCurrentSharingSettings(ShareType.OOF);
                            foreach (WorklistShare worklistShare in worklistShares)
                            {
                                WorkTypes workTypes = worklistShare.WorkTypes;
                                bool needUnShare = false;
                                foreach (WorkType workType in workTypes)
                                {
                                    if (workType.Name == "TaskDelegationWork_" + taskDelegation.Id.ToString())
                                    {
                                        needUnShare = true;
                                        break;
                                    }
                                }

                                if (needUnShare)
                                {
                                    k2Con.UnShare(worklistShare);
                                }
                            }

                            k2Con.SetUserStatus(UserStatuses.Available);

                            string queryUpdate = "UPDATE [dbo].[TaskDelegation] SET";
                            queryUpdate = queryUpdate + " [IsActive] = 0";
                            if (!taskDelegation.IsCanceled)
                            {
                                queryUpdate = queryUpdate + ", [IsExpired] = 1";
                            }
                            queryUpdate = queryUpdate + ", [IsEndedInK2] = 1";
                            queryUpdate = queryUpdate + " WHERE [Id] = @Id";

                            using (SqlCommand cmd = con.CreateCommand())
                            {
                                cmd.CommandType = CommandType.Text;
                                cmd.CommandText = queryUpdate;
                                cmd.Parameters.Add(this.NewSqlParameter("Id", SqlDbType.BigInt, taskDelegation.Id));

                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        innerEx = ex;
                    }
                }

                string queryUpdateSkipped = "[dbo].[TaskDelegation_UpdateSkipped_SP]";

                try
                {
                    using (SqlCommand cmd = con.CreateCommand())
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = queryUpdateSkipped;
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    innerEx = ex;
                }

                if (innerEx != null)
                {
                    throw new Exception("An error occured. " + innerEx.Message, innerEx);
                }

                con.Close();
            }

        }
    }
}