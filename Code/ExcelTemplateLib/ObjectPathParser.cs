using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTemplateLib
{
    public class ObjectPathParser
    {
        public static object GetDeepPropertyValue(object instance, string path)
        {
            var pp = path.Split('.');
            Type t = instance.GetType();
            foreach (var prop in pp)
            {
                var propOrFieldName = prop.Trim();
                // 处理数组
                if (propOrFieldName.EndsWith("]"))
                {
                    // 去掉"]"
                    propOrFieldName = propOrFieldName.Remove(propOrFieldName.Length - 1, 1);
                    // 找到"["的位置
                    var splitPosition = propOrFieldName.IndexOf("[");
                    var propName = propOrFieldName.Substring(0, splitPosition);
                    var arrayIndex = propOrFieldName.Substring(splitPosition + 1);

                    instance = GetObjectFromType(instance, ref t, propName);
                    if (instance is IList)
                    {
                        var list = instance as IList;
                        var index = int.Parse(arrayIndex);
                        if (list.Count <= index)
                            return null;
                        instance = list[index];
                        if (instance == null) return null;
                        t = instance.GetType();
                    }
                }
                else
                {
                    instance = GetObjectFromType(instance, ref t, propOrFieldName);
                    if (instance == null) return null;
                }
            }

            return instance;
        }

        private static object GetObjectFromType(object parentObj, ref Type t, string propOrFieldName)
        {
            object instance = null;
            PropertyInfo propInfo = t.GetProperty(propOrFieldName);
            if (propInfo != null)
            {
                instance = propInfo.GetValue(parentObj, null);
                t = propInfo.PropertyType;
            }
            else
            {
                FieldInfo fieldInfo = t.GetField(propOrFieldName);
                if (fieldInfo != null)
                {
                    instance = fieldInfo.GetValue(parentObj);
                    t = fieldInfo.FieldType;
                }
                else
                {
                    instance = null;
                }
            }

            return instance;
        }
    }
}
