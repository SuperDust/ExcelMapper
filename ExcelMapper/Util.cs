using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMapper
{
    public static class Util
    {
        /// <summary>
        ///     获取属性值
        /// </summary>
        /// <param name="obj">对象</param>
        /// <param name="field">属性/字段</param>
        public static object GetValue<T>(this T obj, string field)
        {
            var index = field.IndexOf(".", StringComparison.CurrentCulture);
            var fieldName = index == -1 ? field : index > field.Length ? field : field[..index];
            var prop = obj.GetType().GetProperty(fieldName);
            if (fieldName.Length == field.Length && prop != null)
            {
                return prop.GetValue(obj);
            }
            if (prop == null)
            {
                return null;
            }
            var propValue = prop.GetValue(obj);
            var len = field.Length - (fieldName.Length + 1);
            return propValue?.GetValue(len > field.Length ? field : field.Substring(field.Length - len, len));
        }

        /// <summary>
        ///     获取T中存在的attribute的属性及attribute值
        /// </summary>
        public static Dictionary<PropertyInfo, T> AttrValues<T>(this Type type) where T : Attribute
        {
            var props =
                type.GetProperties().Where(item => item.CustomAttributes.Any(attr => typeof(T) == attr.AttributeType));
            return props.ToDictionary(item => item, item => item.GetCustomAttribute<T>());
        }
    }
}