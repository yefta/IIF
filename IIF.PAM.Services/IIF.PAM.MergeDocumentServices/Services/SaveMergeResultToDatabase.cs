using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

using IIF.PAM.MergeDocumentServices.Helper;

namespace IIF.PAM.MergeDocumentServices.Services
{
    public class SaveMergeResultToDatabase : BaseServices
    {
        public void SavePAMToDatabase(SqlConnection con, long pamId, byte[] file, string mergeByFQN, string mergeBy, string fileName, bool isPreview = false)
        {
            DBHelper db = new DBHelper();
            
            string fileAsString = Convert.ToBase64String(file);
            fileAsString = "<file><name>" + fileName + "</name><content>" + fileAsString + "</content></file>";

            db.ExecNonQuery(con, "[dbo].[PAM_MergedDocumentResult_Save_SP]", CommandType.StoredProcedure,
                new List<SqlParameter>
                {
                    this.NewSqlParameter("@PAMId", SqlDbType.BigInt, pamId),                    
                    this.NewSqlParameter("@Attachment", SqlDbType.VarChar, fileAsString),
                    this.NewSqlParameter("@MergeByFQN", SqlDbType.VarChar, mergeByFQN),                    
                    this.NewSqlParameter("@MergeBy", SqlDbType.VarChar, mergeBy),
					this.NewSqlParameter("@IsPreview", SqlDbType.Bit, isPreview)
				});
        }

        public void SaveCMToDatabase(SqlConnection con, long pamId, byte[] file, string mergeByFQN, string mergeBy, string fileName, bool isPreview = false)
        {
            DBHelper db = new DBHelper();

            string fileAsString = Convert.ToBase64String(file);
            fileAsString = "<file><name>" + fileName + "</name><content>" + fileAsString + "</content></file>";

            db.ExecNonQuery(con, "[dbo].[CM_MergedDocumentResult_Save_SP]", CommandType.StoredProcedure,
                new List<SqlParameter>
                {
                    this.NewSqlParameter("@CMId", SqlDbType.BigInt, pamId),
                    this.NewSqlParameter("@Attachment", SqlDbType.VarChar, fileAsString),
                    this.NewSqlParameter("@MergeByFQN", SqlDbType.VarChar, mergeByFQN),
                    this.NewSqlParameter("@MergeBy", SqlDbType.VarChar, mergeBy),
					this.NewSqlParameter("@IsPreview", SqlDbType.Bit, isPreview)
				});
        }
    }
}
