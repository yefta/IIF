using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace IIF.PAM.MergeDocumentServices.Helper
{
    public class DBHelper
    {        
        public object ExecScalar(SqlConnection con, string commandText, CommandType commandType, List<SqlParameter> parameters = null)
        {
            if (con == null)
            {
                throw new Exception("Connection cannot be null");
            }
            object obj = new object();
            using (SqlCommand cmd = con.CreateCommand())
            {
                cmd.CommandType = commandType;
                cmd.CommandText = commandText;
                if (parameters != null && parameters.Count > 0)
                {
                    foreach (SqlParameter parameter in parameters)
                    {
                        cmd.Parameters.Add(parameter);
                    }
                }
                obj = cmd.ExecuteScalar();
            }
            return obj;
        }

        public DataTable ExecToDataTable(SqlConnection con, string commandText, CommandType commandType, List<SqlParameter> parameters = null)
        {
            if (con == null)
            {
                throw new Exception("Connection cannot be null");
            }
            DataTable result = new DataTable();

            using (SqlCommand cmd = con.CreateCommand())
            {
                cmd.CommandType = commandType;
                cmd.CommandText = commandText;
                if (parameters != null && parameters.Count > 0)
                {
                    foreach (SqlParameter parameter in parameters)
                    {
                        cmd.Parameters.Add(parameter);
                    }
                }
                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    DataTable schemaTable = dr.GetSchemaTable();                    
                    List<DataColumn> dataColumnList = new List<DataColumn>();
                    if (schemaTable != null)
                    {
                        foreach (DataColumn column in schemaTable.Rows.Cast<DataRow>().Select(dRow => new
                        {
                            dRow = dRow,
                            colName = dRow["ColumnName"].ToString()
                        }).Select(_param0 => new DataColumn(_param0.colName, (Type)_param0.dRow["DataType"])
                        {
                            Unique = (bool)_param0.dRow["IsUnique"],
                            AllowDBNull = (bool)_param0.dRow["AllowDBNull"],
                            AutoIncrement = (bool)_param0.dRow["IsAutoIncrement"]
                        }))
                        {
                            dataColumnList.Add(column);
                            result.Columns.Add(column);
                        }
                    }
                    while (dr.Read())
                    {
                        DataRow row = result.NewRow();
                        for (int index = 0; index < dataColumnList.Count; ++index)
                        {
                            row[dataColumnList[index]] = dr[index];
                        }
                        result.Rows.Add(row);
                    }
                }
            }
            
            return result;
        }

        public SqlDataReader ExecToDataReader(SqlConnection con, string commandText, CommandType commandType, List<SqlParameter> parameters = null)
        {
            if (con == null)
            {
                throw new Exception("Connection cannot be null");
            }

            SqlDataReader result;
            using (SqlCommand cmd = con.CreateCommand())
            {
                cmd.CommandType = commandType;
                cmd.CommandText = commandText;
                if (parameters != null && parameters.Count > 0)
                {
                    foreach (SqlParameter parameter in parameters)
                    {
                        cmd.Parameters.Add(parameter);
                    }
                }
                result = cmd.ExecuteReader();
            }
            return result;
        }

        public List<T> ExecToModel<T>(SqlConnection con, string commandText, CommandType commandType, List<SqlParameter> parameters = null) where T : class
        {
            if (con == null)
            {
                throw new Exception("Connection cannot be null");
            }

            List<T> result;
            using (SqlCommand cmd = con.CreateCommand())
            {
                cmd.CommandType = commandType;
                cmd.CommandText = commandText;
                if (parameters != null && parameters.Count > 0)
                {
                    foreach (SqlParameter parameter in parameters)
                    {
                        cmd.Parameters.Add(parameter);
                    }
                }
                using (SqlDataReader sqlDataReader = cmd.ExecuteReader())
                {
                    result = ModelMapping.DataReaderMapToList<T>((IDataReader)sqlDataReader);
                }
            } return result;
        }

        public void ExecNonQuery(SqlConnection con, string commandText, CommandType commandType, List<SqlParameter> parameters = null)
        {
            if (con == null)
            {
                throw new Exception("Connection cannot be null");
            }

            using (SqlCommand cmd = con.CreateCommand())
            {
                cmd.CommandType = commandType;
                cmd.CommandText = commandText;
                if (parameters != null && parameters.Count > 0)
                {
                    foreach (SqlParameter parameter in parameters)
                    {
                        cmd.Parameters.Add(parameter);
                    }
                }
                cmd.ExecuteNonQuery();
            }
        }

        public DataSet FindData(SqlConnection con, string strTable, long pamId)
        {
            DataSet ds = null;
            string selectStr = "SELECT * FROM " + strTable + " WHERE PAMId = " + pamId.ToString();
            SqlDataAdapter da = new SqlDataAdapter(selectStr, con);
            ds = new DataSet();

            da.Fill(ds, strTable);

            return ds;
        }
    }
}
