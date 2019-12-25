using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Collections;
using System.Globalization;
using Newtonsoft.Json;

namespace xlstool
{
    /// <summary>
    /// 将DataTable对象，转换成Lua Table表，并保存到文件中
    /// </summary>
    class LuaExporter : IExporter
    {
        string mCode;

        public string code
        {
            get
            {
                return this.mCode;
            }
        }

        /// <summary>
        /// 构造函数：完成内部数据创建
        /// </summary>
        /// <param name="sheet">ExcelReader创建的一个表单</param>
        /// <param name="headerRows">表单中的那几行是表头</param>
        public LuaExporter(string excelName, DataTable sheet, int headerRows, bool lowcase)
        {
            
            if (sheet.Columns.Count <= 0)
                return;
            if (sheet.Rows.Count <= 0)
                return;
           
            convertTable(excelName, sheet, headerRows, lowcase);
        }

        private void convertTable(string excelName, DataTable sheet, int headerRows, bool lowcase)
        {
            string structName = Utils.GetStructName(sheet.TableName);

            object import = null;
            if (Utils.IsKV)
            {
                int firstDataRow = 2;
                import = convertRowData(sheet, null, lowcase, firstDataRow);
            }
            else
            {
                Dictionary<string, object> importData = new Dictionary<string, object>();
                int firstDataRow = headerRows;
                for (int i = firstDataRow; i < sheet.Rows.Count; i++)
                {
                    DataRow row = sheet.Rows[i];
                    string ID = row[0].ToString();
                    if (ID.Length <= 0)
                        ID = string.Format("row_{0}", i);

                    importData[ID] = convertRowData(sheet, row, lowcase, firstDataRow);
                }
                import = importData;
            }
            

            

            StringBuilder builder = new StringBuilder(BUILDER_CAPACITY);
            builder.AppendLine("--[[");
            builder.AppendLine("\tAuto Generated Code By xlstool");
            builder.AppendFormat("\tGenerate From {0}.xlsx | {1}", excelName, sheet.TableName);
            builder.AppendLine();
            builder.AppendLine("--]]");

            builder.AppendFormat("local {0} = ", structName);
            builder.AppendLine();
            Encode(import, builder);
            builder.AppendLine();
            builder.AppendFormat("return {0}", structName);

            mCode = builder.ToString();
        }
       
        /// <summary>
        /// 把一行数据转换成一个对象，每一列是一个属性
        /// </summary>
        private object convertRowData(DataTable sheet, DataRow row, bool lowcase, int firstDataRow)
        {
            var rowData = new Dictionary<string, object>();
            int col = 0;
            string fieldName;
            string fieldType;
            if (Utils.IsKV)
            {
                for (int i = 2; i < sheet.Rows.Count; i++)
                {
                    fieldName = sheet.Rows[i][0].ToString();
                    fieldType = sheet.Rows[i][1].ToString();
                    // 表头自动转换成小写
                    if (lowcase)
                        fieldName = fieldName.ToLower();

                    if (string.IsNullOrEmpty(fieldName))
                        fieldName = string.Format("col_{0}", col);

                    object value = sheet.Rows[i][2];
                    if (value.GetType() == typeof(System.DBNull))
                    {
                        throw new Exception(string.Format("cell row:{0} col:{1} is null", i, 2));
                    }
                    else if (value.GetType() == typeof(double))
                    { // 去掉数值字段的“.0”
                        double num = (double)value;
                        if ((int)num == num)
                            value = (int)num;
                    }

                    if (Utils.IsArray(fieldType))
                    {
                        string content = "[";
                        string[] items = Utils.GetArrayItems(value.ToString());
                        if (items != null)
                        {
                            for (int j = 0; j < items.Length; i++)
                            {
                                content += items[j];
                                if (j != items.Length - 1)
                                    content += ",";
                            }
                        }

                        content += "]";
                        object[] arr = JsonConvert.DeserializeObject<object[]>(content);
                        rowData[fieldName] = arr;
                    }
                    else
                        rowData[fieldName] = value;

                }
            }
            else
            {
                foreach (DataColumn column in sheet.Columns)
                {
                    object value = row[column];

                    if (value.GetType() == typeof(System.DBNull))
                    {
                        value = getColumnDefault(sheet, column, firstDataRow);
                    }
                    else if (value.GetType() == typeof(double))
                    { // 去掉数值字段的“.0”
                        double num = (double)value;
                        if ((int)num == num)
                            value = (int)num;
                    }

                    fieldName = sheet.Rows[0][column].ToString();
                    // 表头自动转换成小写
                    if (lowcase)
                        fieldName = fieldName.ToLower();

                    if (string.IsNullOrEmpty(fieldName))
                        fieldName = string.Format("col_{0}", col);

                    fieldType = sheet.Rows[1][column].ToString();
                    if (Utils.IsArray(fieldType))
                    {
                        string content = "[";
                        string[] items = Utils.GetArrayItems(value.ToString());
                        if (items != null)
                        {
                            for (int i = 0; i < items.Length; i++)
                            {
                                content += items[i];
                                if (i != items.Length - 1)
                                    content += ",";
                            }
                        }
                        content += "]";
                        object[] arr = JsonConvert.DeserializeObject<object[]>(content);
                        rowData[fieldName] = arr;
                    }
                    else
                        rowData[fieldName] = value;
                    col++;
                }
            }

            return rowData;
        }

        /// <summary>
        /// 对于表格中的空值，找到一列中的非空值，并构造一个同类型的默认值
        /// </summary>
        private object getColumnDefault(DataTable sheet, DataColumn column, int firstDataRow)
        {
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                object value = sheet.Rows[i][column];
                Type valueType = value.GetType();
                if (valueType != typeof(System.DBNull))
                {
                    if (valueType.IsValueType)
                        return Activator.CreateInstance(valueType);
                    break;
                }
            }
            return "";
        }

        private const int BUILDER_CAPACITY = 2000;

        private  string Encode(object lua, StringBuilder builder)
        {
            bool success = SerializeValue(lua, builder);
            return (success ? builder.ToString() : null);
        }

        private bool SerializeValue(object value, StringBuilder builder)
        {
            bool success = true;

            if (value is string)
            {
                success = SerializeString((string)value, builder);
            }
            else if (value is IDictionary)
            {
                success = SerializeObject((IDictionary)value, builder);
                if (success)
                {
                    builder.AppendLine();
                }
            }
            else if (value is IList)
            {
                success = SerializeArray(value as IList, builder);
            }
            else if ((value is Boolean) && ((Boolean)value == true))
            {
                builder.Append("true");
            }
            else if ((value is Boolean) && ((Boolean)value == false))
            {
                builder.Append("false");
            }
            else if (value is ValueType)
            {
                success = SerializeNumber(Convert.ToDouble(value), builder);
            }
            else if (value == null)
            {
                builder.Append("nil");
            }
            else
            {
                success = false;
            }

            return success;
        }

        private bool SerializeString(string aString, StringBuilder builder, bool key = false)
        {
            if (!key)
                builder.Append("\"");

            /*
            char[] charArray = aString.ToCharArray();
            for (int i = 0; i < charArray.Length; i++)
            {
                char c = charArray[i];
                if (c == '"')
                {
                    builder.Append("\\\"");
                }
                else if (c == '\\')
                {
                    builder.Append("\\\\");
                }
                else if (c == '\b')
                {
                    builder.Append("\\b");
                }
                else if (c == '\f')
                {
                    builder.Append("\\f");
                }
                else if (c == '\n')
                {
                    builder.Append("\\n");
                }
                else if (c == '\r')
                {
                    builder.Append("\\r");
                }
                else if (c == '\t')
                {
                    builder.Append("\\t");
                }
                else
                {
                    int codepoint = Convert.ToInt32(c);
                    if ((codepoint >= 32) && (codepoint <= 126))
                    {
                        builder.Append(c);
                    }
                    else
                    {
                        builder.Append("\\u" + Convert.ToString(codepoint, 16).PadLeft(4, '0'));
                    }
                }
            }
            */
            builder.Append(aString);
            if (!key)
                builder.Append("\"");
            return true;
        }

        private bool SerializeObject(IDictionary anObject, StringBuilder builder)
        {
            builder.Append("{");
            IDictionaryEnumerator e = anObject.GetEnumerator();

            bool first = true;
            while (e.MoveNext())
            {
                string key = e.Key.ToString();
                object value = e.Value;

                if (!first)
                {
                    builder.Append(", ");
                }

                builder.Append("[");
                SerializeString(key, builder, false);
                builder.Append("]");
                
                builder.Append(" = ");
                if (!SerializeValue(value, builder))
                {
                    return false;
                }

                first = false;
            }

            builder.Append("}");
            
            return true;
        }

        private bool SerializeArray(IList anArray, StringBuilder builder)
        {
            builder.Append("{");
            bool first = true;
            for (int i = 0; i < anArray.Count; i++)
            {
                object value = anArray[i];
                if (!first)
                {
                    builder.Append(", ");
                }
                if (!SerializeValue(value, builder))
                {
                    return false;
                }
                first = false;
            }
            builder.Append("}");
            return true;
        }

        private bool SerializeNumber(Double number, StringBuilder builder)
        {
            builder.Append(Convert.ToString(number, CultureInfo.InvariantCulture));
            return true;
        }



        /// <summary>
        /// 将内部数据转换成Lua table，并保存至文件
        /// </summary>
        /// <param name="filePath">存盘文件</param>
        /// <param name="encoding">编码格式</param>
        public void SaveToFile(string filePath, Encoding encoding)
        {
            //-- 保存文件
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(mCode);
            }
        }
    }
}
