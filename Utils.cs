using System.Data;

namespace xlstool
{
    public enum CodeType
    {
        Json,
        Sql,
        CSharp,
        Lua,
        Go,
        TypeScript,
    }

    public class Utils
    {
        public const string KV = "K:V";

        public static bool IsKV
        {
            get;set;
        }
        public static string GetStructName(string tableName)
        {
            string[] flags = tableName.Split('#');
            string structName = "StructName";
            if (flags != null && flags.Length == 2)
            {
                structName = flags[1];
            }

            return structName;
        }

        public static DataTable GetDataTable(DataSet book, string tableName)
        {
            for (int i = 0; i < book.Tables.Count; i++)
            {
                if (book.Tables[i].TableName == tableName)
                    return book.Tables[i];
            }

            return null;
        }

        /// <summary>
        /// 判断是否是数组类型
        /// </summary>
        /// <param name="fieldType"></param>
        /// <returns></returns>
        public static bool IsArray(string fieldType)
        {
            string content = fieldType.Trim();
            return content.Contains("[]");
        }

        /// <summary>
        /// 数组类型
        /// </summary>
        /// <param name="json">1;2;3</param>
        /// <returns></returns>

        public static string[] GetArrayItems(string json)
        {
            char[] chs = json.ToCharArray();
            if (chs == null || chs.Length <=2)
            {
                return null;//数组为空
            }


            string content = json;
            content = content.Trim();
            string[] items = content.Split(';');

            return items;
        }

        /// <summary>
        /// 获取表格数组定义的类型
        /// </summary>
        /// <param name="fieldType">表格类型定义</param>
        /// <returns></returns>
       public static string GetArrayItemType(string fieldType)
       {
            string content = fieldType;
            content = content.Trim();
            content = content.Trim(new char []{ '[', ']'});
            return content;
       }

       public static string ConvertFieldType(CodeType codeType, string fieldType)
       {
            string targetType = fieldType;
            switch (codeType)
            {
                //PostgreSQL
                case CodeType.Sql:
                    if (fieldType == "int")
                        targetType = "INTEGER";
                    else if (fieldType == "int64" || fieldType == "double")
                        targetType = "Numeric";
                    else if (fieldType == "string")
                        targetType = "VARCHAR(128)";
                    else if (fieldType == "bool")
                        targetType = "BOOLEAN";
                    else if (IsArray(fieldType))
                        targetType = "VARCHAR(128)";
                    break;
                case CodeType.Go:
                    if (fieldType == "double")
                        targetType = "float64";
                    break;
                case CodeType.TypeScript:
                    if (fieldType == "int" || fieldType == "int64" || fieldType == "double")
                        targetType = "number";
                    else if (fieldType == "bool")
                        targetType = "boolean";
                    break;
            }

            return targetType;
       }


        public static object GetDefaultValue(CodeType codeType, string fieldType)
        {
            switch (codeType)
            {
                case CodeType.Json:
                default:
                    switch (fieldType)
                    {
                        case "int":
                            return default(int);
                        case "int64":
                            return default(System.Int64);
                        case "double":
                            return default(double);
                        case "string":
                            return "";
                        case "bool":
                            return default(bool);
                        default:
                            break;
                    }
                    return null;
            }
        }
    }

}
