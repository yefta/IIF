using System;
using System.Data.SqlClient;

namespace IIF.PAM.Utilities
{
    public static class DataReaderExtensions
    {
        public static string GetNullableString(this SqlDataReader dr, int index)
        {
            string result = null;
            if (!dr.IsDBNull(index))
            {
                result = dr.GetString(index);
            }
            return result;
        }

        public static DateTime? GetNullableDateTime(this SqlDataReader dr, int index)
        {
            DateTime? result = null;
            if (!dr.IsDBNull(index))
            {
                result = dr.GetDateTime(index);
            }
            return result;
        }
    }
}
