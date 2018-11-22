using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class ReminderServices : BaseServices
    {
        public void ReminderInsertOutboxEmail()
        {
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();

                string querySelect = "SELECT";
                querySelect = querySelect + " [SourceType], [SourceId]";
                querySelect = querySelect + ", [MDocTypeId], [DocumentId], [UserFQN]";
                querySelect = querySelect + ", [Reminder_Subject], [Reminder_Body]";
                querySelect = querySelect + ", [Reminder_IDEmailTemplate], [Reminder_IDReference]";
                querySelect = querySelect + " FROM [dbo].[Vw_PendingReminder]";

                List<Vw_PendingReminder> listReminder = new List<Vw_PendingReminder>();

                using (SqlCommand cmd = con.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = querySelect;

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_SourceType = dr.GetOrdinal("SourceType");
                        int indexOf_SourceId = dr.GetOrdinal("SourceId");
                        int indexOf_MDocTypeId = dr.GetOrdinal("MDocTypeId");
                        int indexOf_DocumentId = dr.GetOrdinal("DocumentId");
                        int indexOf_UserFQN = dr.GetOrdinal("UserFQN");                        
                        int indexOf_Reminder_Subject = dr.GetOrdinal("Reminder_Subject");
                        int indexOf_Reminder_Body = dr.GetOrdinal("Reminder_Body");
                        int indexOf_Reminder_IDEmailTemplate = dr.GetOrdinal("Reminder_IDEmailTemplate");
                        int indexOf_Reminder_IDReference = dr.GetOrdinal("Reminder_IDReference");

                        while (dr.Read())
                        {
                            Vw_PendingReminder reminder = new Vw_PendingReminder();
                            reminder.SourceType = dr.GetInt32(indexOf_SourceType);
                            reminder.SourceId = dr.GetInt64(indexOf_SourceId);
                            reminder.MDocTypeId = dr.GetInt32(indexOf_MDocTypeId);
                            reminder.DocumentId = dr.GetInt64(indexOf_DocumentId);
                            reminder.UserFQN = dr.GetString(indexOf_UserFQN);
                            reminder.Reminder_From = this.AppConfig.SMTPFromEmail;
                            reminder.Reminder_Subject = dr.GetString(indexOf_Reminder_Subject);
                            reminder.Reminder_Body = dr.GetString(indexOf_Reminder_Body);
                            reminder.Reminder_IDEmailTemplate = dr.GetString(indexOf_Reminder_IDEmailTemplate);
                            reminder.Reminder_IDReference = dr.GetString(indexOf_Reminder_IDReference);

                            listReminder.Add(reminder);
                        }
                    }
                }

                foreach (Vw_PendingReminder reminder in listReminder)
                {
                    try
                    {
                        UserADServices svcUserAD = new UserADServices();
                        svcUserAD.AppConfig = this.AppConfig;
                        UserAD userAD = svcUserAD.GetByFQN(reminder.UserFQN);
                        if (userAD != null)
                        {
                            if (!string.IsNullOrEmpty(userAD.Email))
                            {
                                StringBuilder sbQueryInsert = new StringBuilder();
                                sbQueryInsert.Append("INSERT INTO [dbo].[OutboxEmail] (");
                                sbQueryInsert.Append("[From], [To], [Cc], [Bcc], [Subject], [Body]");
                                sbQueryInsert.Append(", [Status], [IDEmailTemplate], [CreatedDate]");
                                sbQueryInsert.Append(", [IDReference], [SendDate]");
                                sbQueryInsert.Append(") VALUES (");
                                sbQueryInsert.Append("@From, @To, NULL, NULL, @Subject, @Body");
                                sbQueryInsert.Append(", NULL, @IDEmailTemplate, GETDATE()");
                                sbQueryInsert.Append(", @IDReference, NULL");
                                sbQueryInsert.Append(");");

                                sbQueryInsert.Append("UPDATE [dbo].[WorkflowHistory] SET");
                                sbQueryInsert.Append("[LastReminderDate] = GETDATE()");
                                sbQueryInsert.Append("WHERE [Id] = @SourceId");

                                using (SqlCommand cmd = con.CreateCommand())
                                {
                                    cmd.CommandType = CommandType.Text;
                                    cmd.CommandText = sbQueryInsert.ToString();

                                    cmd.Parameters.Add(this.NewSqlParameter("From", SqlDbType.VarChar, reminder.Reminder_From));
                                    cmd.Parameters.Add(this.NewSqlParameter("To", SqlDbType.VarChar, userAD.Email));
                                    cmd.Parameters.Add(this.NewSqlParameter("Subject", SqlDbType.VarChar, reminder.Reminder_Subject));
                                    cmd.Parameters.Add(this.NewSqlParameter("Body", SqlDbType.VarChar, reminder.Reminder_Body));
                                    cmd.Parameters.Add(this.NewSqlParameter("IDEmailTemplate", SqlDbType.VarChar, reminder.Reminder_IDEmailTemplate));
                                    cmd.Parameters.Add(this.NewSqlParameter("IDReference", SqlDbType.VarChar, reminder.Reminder_IDReference));

                                    cmd.Parameters.Add(this.NewSqlParameter("SourceId", SqlDbType.BigInt, reminder.SourceId));

                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex.Message, ex);
                    }
                }
                con.Close();
            }

        }
    }
}