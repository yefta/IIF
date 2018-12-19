using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;

using IIF.PAM.Utilities;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class OutboxEmailServices: BaseServices
    {
        public void InsertPAMGroupEmail(PAMGroupEmailParameter param)
        {
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();

                using (SqlTransaction tran = con.BeginTransaction())
                {
                    string querySelect = "[dbo].[PAM_ListApproval_SP]";

                    List<string> listUserFQN = new List<string>();

                    using (SqlCommand cmd = con.CreateCommand())
                    {
                        cmd.Transaction = tran;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = querySelect;
                        cmd.Parameters.Add(this.NewSqlParameter("PAMId", SqlDbType.BigInt, param.PAMId));
                        cmd.Parameters.Add(this.NewSqlParameter("MRoleId", SqlDbType.Int, param.MRoleId));

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            int indexOf_UserFQN = dr.GetOrdinal("UserFQN");

                            while (dr.Read())
                            {
                                string userFQN = dr.GetString(indexOf_UserFQN);
                                listUserFQN.Add(userFQN);
                            }
                        }
                    }

                    this.InsertGroupEmail(con, tran, param, listUserFQN, "PAM-" + param.PAMId.ToString());

                    tran.Commit();
                }
                con.Close();
            }

        }
        
        public void InsertCMGroupEmail(CMGroupEmailParameter param)
        {
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();

                using (SqlTransaction tran = con.BeginTransaction())
                {
                    string querySelect = "[dbo].[CM_ListApproval_SP]";

                    List<string> listUserFQN = new List<string>();

                    using (SqlCommand cmd = con.CreateCommand())
                    {
                        cmd.Transaction = tran;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = querySelect;
                        cmd.Parameters.Add(this.NewSqlParameter("CMId", SqlDbType.BigInt, param.CMId));
                        cmd.Parameters.Add(this.NewSqlParameter("MRoleId", SqlDbType.Int, param.MRoleId));

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            int indexOf_UserFQN = dr.GetOrdinal("UserFQN");

                            while (dr.Read())
                            {
                                string userFQN = dr.GetString(indexOf_UserFQN);
                                listUserFQN.Add(userFQN);
                            }
                        }
                    }

                    this.InsertGroupEmail(con, tran, param, listUserFQN, "CM-" + param.CMId.ToString());

                    tran.Commit();
                }
                con.Close();
            }

        }

        public void InsertGroupEmail(SqlConnection con, SqlTransaction tran, BaseGroupEmailParameter param, List<string> listUserFQN, string IDReference)
        {
            StringBuilder sbQueryInsert = new StringBuilder();
            sbQueryInsert.Append("INSERT INTO [dbo].[OutboxEmail] (");
            sbQueryInsert.Append("[From], [To], [Cc], [Bcc], [Subject], [Body]");
            sbQueryInsert.Append(", [Status], [IDEmailTemplate], [CreatedDate]");
            sbQueryInsert.Append(", [IDReference], [SendDate]");
            sbQueryInsert.Append(") VALUES ");

            int userIndex = 0;
            int userHasEmail = 0;
            using (SqlCommand cmd = con.CreateCommand())
            {
                foreach (string userFQN in listUserFQN)
                {
                    UserADServices svcUserAD = new UserADServices();
                    svcUserAD.AppConfig = this.AppConfig;
                    UserAD userAD = svcUserAD.GetByFQN(userFQN);
                    if (userAD != null)
                    {
                        if (!string.IsNullOrEmpty(userAD.Email))
                        {
                            if (userIndex > 0)
                            {
                                sbQueryInsert.Append(",");
                            }
                            sbQueryInsert.Append("(");
                            sbQueryInsert.Append("@From, @To_" + userIndex.ToString() + ", NULL, NULL, @Subject, @Body");
                            sbQueryInsert.Append(", '0', @IDEmailTemplate, GETDATE()");
                            sbQueryInsert.Append(", @IDReference, NULL");
                            sbQueryInsert.Append(")");

                            cmd.Parameters.Add(this.NewSqlParameter("To_" + userIndex.ToString(), SqlDbType.VarChar, userAD.Email));

                            userIndex++;
                            userHasEmail++;
                        }
                    }
                }

                if (userHasEmail > 0)
                {

                    cmd.Transaction = tran;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = sbQueryInsert.ToString();

                    cmd.Parameters.Add(this.NewSqlParameter("From", SqlDbType.VarChar, param.From));
                    cmd.Parameters.Add(this.NewSqlParameter("Subject", SqlDbType.VarChar, param.Subject));
                    cmd.Parameters.Add(this.NewSqlParameter("Body", SqlDbType.VarChar, param.Body));
                    cmd.Parameters.Add(this.NewSqlParameter("IDEmailTemplate", SqlDbType.VarChar, param.IDEmailTemplate));
                    cmd.Parameters.Add(this.NewSqlParameter("IDReference", SqlDbType.VarChar, "CM-" + IDReference));

                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void SendEmailOutbox()
        {
            string conStringIIF = this.AppConfig.IIFConnectionString;
            using (SqlConnection con = new SqlConnection(conStringIIF))
            {
                con.Open();

                List<OutboxEmail> listOutboxEmail = new List<OutboxEmail>();

                string querySelect = "SELECT";
                querySelect = querySelect + " [Src].[Id], [Src].[From], [Src].[To]";
                querySelect = querySelect + ", [Src].[Cc], [Src].[Bcc]";
                querySelect = querySelect + ", [Src].[Subject], [Src].[Body]";
                querySelect = querySelect + ", [Src].[Status], [Src].[IDEmailTemplate], [Src].[CreatedDate]";
                querySelect = querySelect + ", [Src].[IDReference], [Src].[SendDate]";
                querySelect = querySelect + " FROM [dbo].[Vw_OutboxEmail_NeedSend] AS [Src]";

                using (SqlCommand cmd = con.CreateCommand())
                {                    
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = querySelect;

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        int indexOf_Id = dr.GetOrdinal("Id");
                        int indexOf_From = dr.GetOrdinal("From");
                        int indexOf_To = dr.GetOrdinal("To");
                        int indexOf_Cc = dr.GetOrdinal("Cc");
                        int indexOf_Bcc = dr.GetOrdinal("Bcc");
                        int indexOf_Subject = dr.GetOrdinal("Subject");
                        int indexOf_Body = dr.GetOrdinal("Body");
                        int indexOf_Status = dr.GetOrdinal("Status");
                        int indexOf_IDEmailTemplate = dr.GetOrdinal("IDEmailTemplate");
                        int indexOf_CreatedDate = dr.GetOrdinal("CreatedDate");
                        int indexOf_IDReference = dr.GetOrdinal("IDReference");
                        int indexOf_SendDate = dr.GetOrdinal("SendDate");

                        while (dr.Read())
                        {
                            OutboxEmail data = new OutboxEmail();
                            data.Id = dr.GetGuid(indexOf_Id);
                            data.From = dr.GetNullableString(indexOf_From);
                            data.To = dr.GetNullableString(indexOf_To);
                            data.Cc = dr.GetNullableString(indexOf_Cc);
                            data.Bcc = dr.GetNullableString(indexOf_Bcc);
                            data.Subject = dr.GetNullableString(indexOf_Subject);
                            data.Body = dr.GetNullableString(indexOf_Body);
                            data.Status = dr.GetNullableString(indexOf_Status);
                            data.IDEmailTemplate = dr.GetNullableString(indexOf_IDEmailTemplate);
                            data.CreatedDate = dr.GetNullableDateTime(indexOf_CreatedDate);
                            data.IDReference = dr.GetNullableString(indexOf_IDReference);
                            data.SendDate = dr.GetNullableDateTime(indexOf_SendDate);
                            
                            listOutboxEmail.Add(data);
                        }
                    }
                }

                string queryUpdate = "[dbo].[OutboxEmail_UpdateStatusSend_SP]";

                foreach (OutboxEmail outboxEmail in listOutboxEmail)
                {
                    try
                    {
                        this.SendOneEmail(outboxEmail);

                        using (SqlCommand cmd = con.CreateCommand())
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.CommandText = queryUpdate;

                            cmd.Parameters.Add(this.NewSqlParameter("Id", SqlDbType.UniqueIdentifier, outboxEmail.Id));
                            cmd.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex.Message, ex);
                    }
                }
            }
        }

        public void SendOneEmail(OutboxEmail data)
        {
            string mailFrom = data.From;
            string mailTo = data.To;
            string mailCC = data.Cc;
            string mailBCC = data.Bcc;
            string mailSubject = data.Subject;
            string mailBody = data.Body;

            string smtpHost = this.AppConfig.SMTPHost;
            int smtpPort = this.AppConfig.SMTPPort;
            bool smtpEnableSSL = this.AppConfig.SMTPEnableSSL;
            string smtpCredentialName = this.AppConfig.SMTPCredentialName;
            string smtpCredentialPassword = this.AppConfig.SMTPCredentialPassword;
            bool canSend = true;
            if (string.IsNullOrEmpty(mailTo))
            {
                canSend = false;
            }
            if (string.IsNullOrEmpty(mailFrom))
            {
                canSend = false;
            }

            if (canSend)
            {
                MailMessage emailMessage = new MailMessage();

                emailMessage.From = new MailAddress(mailFrom, this.AppConfig.SMTPFromName);                
                IEnumerable<string> mailToSplitted = mailTo.Split(new string[] {";", ","},StringSplitOptions.RemoveEmptyEntries).Distinct();
                foreach (string oneMailTo in mailToSplitted)
                {
                    try
                    {
                        emailMessage.To.Add(oneMailTo.Trim());
                    }
                    catch (FormatException exFormat)
                    {
                        throw new FormatException("'" + oneMailTo.Trim() + "' is not a valid email address.", exFormat);
                    }
                }

                if (!string.IsNullOrEmpty(mailCC))
                {
                    IEnumerable<string> mailCCSplitted = mailCC.Split(new string[] { ";", "," }, StringSplitOptions.RemoveEmptyEntries).Distinct();
                    foreach (string oneMailCC in mailCCSplitted)
                    {
                        try
                        {
                            emailMessage.To.Add(oneMailCC.Trim());
                        }
                        catch (FormatException exFormat)
                        {
                            throw new FormatException("'" + oneMailCC.Trim() + "' is not a valid email address.", exFormat);
                        }
                    }
                }

                if (!string.IsNullOrEmpty(mailBCC))
                {
                    IEnumerable<string> mailBCCSplitted = mailBCC.Split(new string[] { ";", "," }, StringSplitOptions.RemoveEmptyEntries).Distinct();
                    foreach (string oneMailBCC in mailBCCSplitted)
                    {
                        try
                        {
                            emailMessage.To.Add(oneMailBCC.Trim());
                        }
                        catch (FormatException exFormat)
                        {
                            throw new FormatException("'" + oneMailBCC.Trim() + "' is not a valid email address.", exFormat);
                        }
                    }
                }
                
                emailMessage.Subject = mailSubject;
                emailMessage.IsBodyHtml = true;
                emailMessage.Body = mailBody;

                SmtpClient smtpClient = new SmtpClient();
                smtpClient.UseDefaultCredentials = true;
                smtpClient.Host = smtpHost;
                smtpClient.Port = smtpPort;
                smtpClient.EnableSsl = smtpEnableSSL;

                if (smtpCredentialName != string.Empty && smtpCredentialPassword != string.Empty)
                {
                    NetworkCredential credential = new NetworkCredential(smtpCredentialName, smtpCredentialPassword);
                    smtpClient.Credentials = credential;
                }
                smtpClient.Send(emailMessage);
            }
        }
    }
}