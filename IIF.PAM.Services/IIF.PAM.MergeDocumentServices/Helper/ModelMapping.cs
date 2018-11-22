using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;

namespace IIF.PAM.MergeDocumentServices.Helper
{
    public class ModelMapping
    {
        private static T MapToMyModel<T>(object reader) where T : class
        {
            SqlDataReader sqlDataReader = (SqlDataReader)reader;
            T instance = Activator.CreateInstance<T>();
            foreach (PropertyInfo property in instance.GetType().GetProperties())
            {
                MappingAttribute[] customAttributes = (MappingAttribute[])property.GetCustomAttributes(typeof(MappingAttribute), true);
                if (customAttributes.Length > 0 && customAttributes[0].ColumnName != null)
                    property.SetValue((object)instance, Convert.ChangeType(sqlDataReader[customAttributes[0].ColumnName], property.PropertyType), (object[])null);
            }
            return instance;
        }

        public static T MapToModel<T>(SqlDataReader reader) where T : class
        {
            return ModelMapping.MapToMyModel<T>((object)reader);
        }

        public static T MapToModel<T>(OleDbDataAdapter reader) where T : class
        {
            return ModelMapping.MapToMyModel<T>((object)reader);
        }

        public static List<T> DataReaderMapToList<T>(IDataReader dr)
        {
            List<T> objList = new List<T>();
            while (dr.Read())
            {
                T instance = Activator.CreateInstance<T>();
                foreach (PropertyInfo propertyInfo in ((IEnumerable<PropertyInfo>)instance.GetType().GetProperties()).Where<PropertyInfo>((Func<PropertyInfo, bool>)(prop => !object.Equals(dr[prop.Name], (object)DBNull.Value))))
                    propertyInfo.SetValue((object)instance, dr[propertyInfo.Name] is Guid ? (object)((Guid)dr[propertyInfo.Name]).ToString("D") : dr[propertyInfo.Name], (object[])null);
                objList.Add(instance);
            }
            return objList;
        }
    }

    [AttributeUsage(AttributeTargets.Property, Inherited = true)]
    [Serializable]
    public class MappingAttribute : Attribute
    {
        public string ColumnName;
    }
}
