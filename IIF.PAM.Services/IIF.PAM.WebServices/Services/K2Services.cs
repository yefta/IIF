using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

using SourceCode.Hosting.Client.BaseAPI;
using SourceCode.SmartObjects.Client;
using SourceCode.Workflow.Client;

using IIF.PAM.Utilities;

using IIF.PAM.WebServices;
using IIF.PAM.WebServices.Models;
using IIF.PAM.WebServices.Services;

namespace IIF.PAM.WebServices.Services
{
    public class K2Services : BaseServices
    {
        public SmartObjectClientServer NewSmartObjectClientServer()
        {
            SCConnectionStringBuilder hostServerConnectionString = new SCConnectionStringBuilder();
            hostServerConnectionString.Host = this.AppConfig.K2Server;
            hostServerConnectionString.Port = 5555;
            hostServerConnectionString.IsPrimaryLogin = true;
            hostServerConnectionString.Integrated = true;
            SmartObjectClientServer result = new SmartObjectClientServer();
            result.CreateConnection();
            //open the connection to the K2 server
            result.Connection.Open(hostServerConnectionString.ToString());
            //return the SOClientServer object
            return result;
        }

        public List<IIFWorklistItem> ListIIFWorklistItem(IIFWorklistItem_Filter filter)
        {
            List<IIFWorklistItem> result = new List<IIFWorklistItem>();

            Dictionary<int, List<IIFWorklistItem>> dictWfItem = new Dictionary<int, List<IIFWorklistItem>>();

            using (Connection k2Con = new Connection())
            {
                k2Con.Open(this.AppConfig.K2Server);
                if (filter != null)
                {
                    if (!string.IsNullOrEmpty(filter.Destination))
                    {
                        k2Con.ImpersonateUser(filter.Destination);
                    }
                }

                WorklistCriteria worklistCriteria = new WorklistCriteria();
                worklistCriteria.Platform = "ASP";
                worklistCriteria.AddFilterField(WCField.ProcessFullName, WCCompare.Equal, "IIF\\PAM");
                worklistCriteria.AddFilterField(WCLogical.Or, WCField.ProcessFullName, WCCompare.Equal, "IIF\\CM");
                worklistCriteria.AddFilterField(WCLogical.Or, WCField.WorklistItemOwner, "Me", WCCompare.Equal, WCWorklistItemOwner.Me);
                worklistCriteria.AddFilterField(WCLogical.Or, WCField.WorklistItemOwner, "Other", WCCompare.Equal, WCWorklistItemOwner.Other);

                Worklist worklist = k2Con.OpenWorklist(worklistCriteria);

                foreach (WorklistItem item in worklist)
                {
                    IIFWorklistItem newData = new IIFWorklistItem();
                    newData.K2ProcessId = item.ProcessInstance.ID;
                    newData.SN = item.SerialNumber;
                    newData.K2CurrentActivityName = item.ActivityInstanceDestination.DisplayName;
                    if (!dictWfItem.ContainsKey(newData.K2ProcessId))
                    {
                        dictWfItem.Add(newData.K2ProcessId, new List<IIFWorklistItem>());
                    }
                    newData.SharedUserFQN = item.AllocatedUser;
                    dictWfItem[newData.K2ProcessId].Add(newData);
                }
            }

            if (dictWfItem.Count > 0)
            {
                string conStringIIF = this.AppConfig.IIFConnectionString;
                using (SqlConnection con = new SqlConnection(conStringIIF))
                {
                    con.Open();

                    string queryInValue = string.Empty;
                    foreach (int k2ProcessId in dictWfItem.Keys)
                    {
                        if (string.IsNullOrEmpty(queryInValue))
                        {
                            queryInValue = k2ProcessId.ToString();
                        }
                        else
                        {
                            queryInValue = queryInValue + ", " + k2ProcessId.ToString();
                        }
                    }

                    string query = "SELECT [K2ProcessId], [Id]";
                    query = query + ", [MDocTypeId], [MDocTypeName], [MProductTypeId], [MProductTypeName]";
                    query = query + ", [ProjectCode], [CustomerName], [SubmitDate], [CMNumber]";
                    query = query + ", [IsInRevise]";
                    query = query + ", [MWorkflowStatusId], [MWorkflowStatusName]";
                    query = query + ", [ModifiedBy], [ModifiedOn]";
                    query = query + " FROM [dbo].[Vw_SubmissionList]";
                    query = query + " WHERE [K2ProcessId] IN (";
                    query = query + queryInValue;
                    query = query + " )";

                    using (SqlCommand cmd = con.CreateCommand())
                    {
                        cmd.CommandType = CommandType.Text;

                        #region Filter
                        if (filter != null)
                        {
                            if (filter.SubmitDate_FROM.HasValue)
                            {
                                query = query + " AND CONVERT(DATE, [SubmitDate]) >= @SubmitDate_FROM";
                                cmd.Parameters.Add(this.NewSqlParameter("SubmitDate_FROM", SqlDbType.Date, filter.SubmitDate_FROM));
                            }

                            if (filter.SubmitDate_TO.HasValue)
                            {
                                query = query + " AND CONVERT(DATE, [SubmitDate]) <= @SubmitDate_TO";
                                cmd.Parameters.Add(this.NewSqlParameter("SubmitDate_TO", SqlDbType.Date, filter.SubmitDate_TO));
                            }

                            if (!string.IsNullOrEmpty(filter.ProjectCode_LIKE))
                            {
                                query = query + " AND [ProjectCode] LIKE '%' + @ProjectCode_LIKE + '%'";
                                cmd.Parameters.Add(this.NewSqlParameter("ProjectCode_LIKE", SqlDbType.VarChar, filter.ProjectCode_LIKE));
                            }

                            if (!string.IsNullOrEmpty(filter.CustomerName_LIKE))
                            {
                                query = query + " AND [CustomerName] LIKE '%' + @CustomerName_LIKE + '%'";
                                cmd.Parameters.Add(this.NewSqlParameter("CustomerName_LIKE", SqlDbType.VarChar, filter.CustomerName_LIKE));
                            }

                            if (filter.ProductTypeId.HasValue)
                            {
                                query = query + " AND [MProductTypeId] = @ProductTypeId";
                                cmd.Parameters.Add(this.NewSqlParameter("ProductTypeId", SqlDbType.Int, filter.ProductTypeId));
                            }

                            if (filter.MDocTypeId.HasValue)
                            {
                                query = query + " AND [MDocTypeId] = @MDocTypeId";
                                cmd.Parameters.Add(this.NewSqlParameter("MDocTypeId", SqlDbType.VarChar, filter.MDocTypeId.Value));
                            }

                            if (!string.IsNullOrEmpty(filter.CMNumber_LIKE))
                            {
                                query = query + " AND [CMNumber] LIKE '%' + @CMNumber_LIKE + '%'";
                                cmd.Parameters.Add(this.NewSqlParameter("CMNumber_LIKE", SqlDbType.VarChar, filter.CMNumber_LIKE));
                            }
                        }
                        #endregion

                        cmd.CommandText = query;

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            int indexOf_K2ProcessId = dr.GetOrdinal("K2ProcessId");
                            int indexOf_Id = dr.GetOrdinal("Id");
                            int indexOf_MDocTypeId = dr.GetOrdinal("MDocTypeId");
                            int indexOf_MDocTypeName = dr.GetOrdinal("MDocTypeName");
                            int indexOf_MProductTypeId = dr.GetOrdinal("MProductTypeId");
                            int indexOf_MProductTypeName = dr.GetOrdinal("MProductTypeName");
                            int indexOf_ProjectCode = dr.GetOrdinal("ProjectCode");
                            int indexOf_CustomerName = dr.GetOrdinal("CustomerName");
                            int indexOf_SubmitDate = dr.GetOrdinal("SubmitDate");
                            int indexOf_CMNumber = dr.GetOrdinal("CMNumber");
                            int indexOf_IsInRevise = dr.GetOrdinal("IsInRevise");
                            int indexOf_MWorkflowStatusId = dr.GetOrdinal("MWorkflowStatusId");
                            int indexOf_MWorkflowStatusName = dr.GetOrdinal("MWorkflowStatusName");
                            int indexOf_ModifiedBy = dr.GetOrdinal("ModifiedBy");
                            int indexOf_ModifiedOn = dr.GetOrdinal("ModifiedOn");

                            while (dr.Read())
                            {
                                int k2ProcessId = dr.GetInt32(indexOf_K2ProcessId);

                                if (dictWfItem.ContainsKey(k2ProcessId))
                                {
                                    foreach (IIFWorklistItem data in dictWfItem[k2ProcessId])
                                    {
                                        Logger.Error("DocumentId: " + dr.GetInt64(indexOf_Id));
                                        Logger.Error("MDocTypeId: " + dr.GetInt32(indexOf_MDocTypeId));
                                        Logger.Error("MDocTypeName: " + dr.GetNullableString(indexOf_MDocTypeName));
                                        Logger.Error("ProductTypeId: " + dr.GetInt32(indexOf_MProductTypeId));
                                        Logger.Error("ProductTypeName: " + dr.GetNullableString(indexOf_MProductTypeName));
                                        Logger.Error("ProjectCode: " + dr.GetNullableString(indexOf_ProjectCode));
                                        Logger.Error("CustomerName: " + dr.GetNullableString(indexOf_CustomerName));
                                        //Logger.Error("SubmitDate: " + dr.GetDateTime(indexOf_SubmitDate));
                                        Logger.Error("CMNumber: " + dr.GetNullableString(indexOf_CMNumber));
                                        Logger.Error("IsInRevise: " + dr.GetBoolean(indexOf_IsInRevise));
                                        Logger.Error("WorkflowStatusId: " + dr.GetInt32(indexOf_MWorkflowStatusId));
                                        Logger.Error("WorkflowStatusName: " + dr.GetNullableString(indexOf_MWorkflowStatusName));
                                        Logger.Error("ModifiedBy: " + dr.GetNullableString(indexOf_ModifiedBy));
                                        //Logger.Error("ModifiedOn: " + dr.GetDateTime(indexOf_ModifiedOn));

                                        data.DocumentId = dr.GetInt64(indexOf_Id);
                                        data.MDocTypeId = dr.GetInt32(indexOf_MDocTypeId);
                                        data.MDocTypeName = dr.GetNullableString(indexOf_MDocTypeName);
                                        data.ProductTypeId = dr.GetInt32(indexOf_MProductTypeId);
                                        data.ProductTypeName = dr.GetNullableString(indexOf_MProductTypeName);
                                        data.ProjectCode = dr.GetNullableString(indexOf_ProjectCode);
                                        data.CustomerName = dr.GetNullableString(indexOf_CustomerName);
                                        try
                                        {
                                            data.SubmitDate = dr.GetDateTime(indexOf_SubmitDate);
                                        }
                                        catch
                                        {
                                        }

                                        data.CMNumber = dr.GetNullableString(indexOf_CMNumber);
                                        data.IsInRevise = dr.GetBoolean(indexOf_IsInRevise);
                                        data.WorkflowStatusId = dr.GetInt32(indexOf_MWorkflowStatusId);
                                        data.WorkflowStatusName = dr.GetNullableString(indexOf_MWorkflowStatusName);
                                        data.ModifiedBy = dr.GetNullableString(indexOf_ModifiedBy);

                                        try
                                        {
                                            data.ModifiedOn = dr.GetDateTime(indexOf_ModifiedOn);
                                        }
                                        catch
                                        {
                                        }

                                        if ((data.K2CurrentActivityName.ToUpper() == "Submit MoM BoD".ToUpper()) || (data.K2CurrentActivityName.ToUpper() == "Submit MoM BoC".ToUpper()))
                                        {
                                            data.TaskListStatus = "Responded (Risk Team)";
                                        }
                                        else
                                        {
                                            data.TaskListStatus = data.WorkflowStatusName;
                                        }

                                        result.Add(data);
                                    }
                                }
                            }
                        }
                    }
                    con.Close();
                }
            }

            return result;
        }


        public void RetryWorkflow()
        {
            SmartObjectClientServer soServer = this.NewSmartObjectClientServer();
            try
            {
                using (soServer.Connection)
                {
                    SmartObject soError = soServer.GetSmartObject("com_K2_System_Workflow_SmartObject_ErrorLog");
                    //set method we want to execute.
                    soError.MethodToExecute = "GetErrorLogs";
                    soError.Properties["ErrorProfileName"].Value = "All";

                    //get the list of SmartObjects returned by the method
                    SmartObjectList soListError = soServer.ExecuteList(soError);
                    //iterate over the collection

                    List<K2ErrorLog> listK2ErrorLog = new List<K2ErrorLog>();
                    string[] workflowNames = this.AppConfig.WorkflowNames.Split(';');
                    foreach (SmartObject soItem in soListError.SmartObjectsList)
                    {
                        for (int i = 0; i < workflowNames.Length; i++)
                        {
                            if (soItem.Properties["ProcessName"].Value == workflowNames[i])
                            {
                                string soDescription = soItem.Properties["Description"].Value.ToString();
                                if (soDescription.Contains("was deadlocked on lock resources with another process")
                                    || soDescription.Contains("SQL")
                                    || soDescription.Contains("IIF"))
                                {
                                    listK2ErrorLog.Add(new K2ErrorLog(soItem));
                                }
                            }
                        }
                        //ambil hanya transaksi deadlocked saja 
                    }

                    foreach (K2ErrorLog k2ErrorLog in listK2ErrorLog)
                    {
                        SmartObject errorSO = soServer.GetSmartObject("com_K2_System_Workflow_SmartObject_ErrorLog");
                        //set method we want to execute.
                        errorSO.MethodToExecute = "RetryError";

                        errorSO.Properties["Id"].Value = k2ErrorLog.Id.ToString();
                        errorSO.Properties["ProcInstId"].Value = k2ErrorLog.ProcInstID.ToString();
                        errorSO.Properties["TypeId"].Value = k2ErrorLog.TypeID.ToString();
                        errorSO.Properties["ObjectId"].Value = k2ErrorLog.ObjectID.ToString();
                        errorSO.Properties["UserName"].Value = "System";

                        this.Logger.Info("K2 Retry ProcInstId = " + k2ErrorLog.ProcInstID.ToString());

                        //get the list of SmartObjects returned by the method
                        Thread.Sleep(1000); //delay 1 detik
                        soServer.ExecuteScalar(errorSO);
                    }
                }
            }
            finally
            {
                soServer.DeleteConnection();
            }            
        }

    }
}