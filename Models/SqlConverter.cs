
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class SqlConverter
    {
        //public static T ConvertToObject<T>(this SqlDataReader rd) where T : class, new() // possibly a fast way to read in sql data to a model if varaibles are the same between each
        //{
        //    Type type = typeof(T);
        //    var accessor = TypeAccessor.Create(type);
        //    var members = accessor.GetMembers();
        //    var t = new T();

        //    for (int i = 0; i < rd.FieldCount; i++)
        //    {
        //        if (!rd.IsDBNull(i))
        //        {
        //            string fieldName = rd.GetName(i);

        //            if (members.Any(m => string.Equals(m.Name, fieldName, StringComparison.OrdinalIgnoreCase)))
        //            {
        //                accessor[t, fieldName] = rd.GetValue(i);
        //            }
        //        }
        //    }

        //    return t;
        //}

        //public static T ConvertFromDBVal<T>(object obj)
        //{
        //    if (obj == null || obj == DBNull.Value)
        //    {
        //        return default(T); // returns the default value for the type
        //    }
        //    else
        //    {
        //        return (T)obj;
        //    }
        //}

    
}
}
