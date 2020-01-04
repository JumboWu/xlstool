using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;

namespace xlstool {
    /// <summary>
    /// 根据表头，生成C#类定义数据结构
    /// 表头使用三行定义：字段名称、字段类型、注释
    /// </summary>
    class CSDefineGenerator : IExporter {
        struct FieldDef {
            public string name;
            public string type;
            public string comment;
        }

        string mCode;

        public string code {
            get {
                return this.mCode;
            }
        }

        public CSDefineGenerator(string excelName, DataTable sheet, bool lowcase) {
            string structName = Utils.GetStructName(sheet.TableName);
            List<FieldDef> m_fieldList = new List<FieldDef>();
            
            if (sheet.Rows.Count < 3)
                return;
            if (Utils.IsKV)
            {
                if (sheet.Columns.Count < 4)
                    return;

                for(int i = 2; i < sheet.Rows.Count; i++)
                {
                    FieldDef field;
                    field.name = sheet.Rows[i][0].ToString();
                    field.type = sheet.Rows[i][1].ToString();
                    if (string.IsNullOrEmpty(field.name) || string.IsNullOrEmpty(field.type))
                        continue;

                    if (Utils.IsArray(field.type))
                    {
                        field.type = Utils.GetArrayItemType(field.type) + "[]";
                    }
                    field.comment = sheet.Rows[i][4].ToString();

                    if (lowcase)
                        field.name = field.name.ToLower();

                    m_fieldList.Add(field);
                }
            }
            else
            {
                DataRow nameRow = sheet.Rows[0];
                DataRow typeRow = sheet.Rows[1];
                //sheet.Rows[2] platform 相关
                DataRow commentRow = sheet.Rows[3];

                foreach (DataColumn column in sheet.Columns)
                {
                    FieldDef field;
                    field.name = nameRow[column].ToString();
                    field.type = typeRow[column].ToString();
                    if (string.IsNullOrEmpty(field.name) || string.IsNullOrEmpty(field.type))
                        continue;

                    if (Utils.IsArray(field.type))
                    {
                        field.type = Utils.GetArrayItemType(field.type) + "[]";
                    }

                    field.comment = commentRow[column].ToString();

                    if (lowcase)
                        field.name = field.name.ToLower();

                    m_fieldList.Add(field);
                }
            }


            //-- 创建代码字符串
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("// Auto Generated Code By xlstool");
            sb.AppendFormat("// Generate From {0}.xlsx | {1}", excelName, sheet.TableName);
            sb.AppendLine();
            sb.AppendFormat("public class {0}\r\n{{", structName);
            sb.AppendLine();

            foreach (FieldDef field in m_fieldList) {
                sb.AppendFormat("\tpublic {0} {1}; // {2}", field.type, field.name, field.comment);
                sb.AppendLine();
            }

            sb.Append('}');
            sb.AppendLine();
            sb.AppendLine("// End of Auto Generated Code");

            mCode = sb.ToString();
        }

        /// <summary>
        /// 保存CS结构定义
        /// </summary>
        /// <param name="filePath">存盘文件</param>
        /// <param name="encoding">编码格式</param>
        public void SaveToFile(string filePath, Encoding encoding) {
            //-- 保存文件
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write)) {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(mCode);
            }
        }
    }
}
