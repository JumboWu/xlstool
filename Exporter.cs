using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace xlstool
{
    class Exporter : IExporter
    {
        string mContent;
        /// <summary>
        /// 初始化内部数据 支持tsv csv导出
        /// </summary>
        /// <param name="sheet">Excel读取的一个表单</param>
        /// <param name="format">导出格式tsv/csv</param>
        public Exporter(DataTable sheet, string format = "tsv")
        {
            if (format == "tsv")
            {
                convertBySplit(sheet, '\t');
            }
            else
            {
                convertBySplit(sheet, ',');
            }
        }

        private void convertBySplit(DataTable sheet, char split = ',')
        {
            StringBuilder builder = new StringBuilder();
            var rows = sheet.Rows;
            var cols = sheet.Columns;

            for (int i = 0; i < rows.Count; i++)
            {
                for (int j = 0; j < cols.Count; j++)
                {
                    builder.Append(rows[i][j]);
                    if (j != cols.Count)
                        builder.Append(split);
                }
                if (i != rows.Count)
                    builder.AppendLine();
            }

            mContent = builder.ToString();
        }

        /// <summary>
        /// 转换成CSV/TSV 保存
        /// </summary>
        /// <param name="filePath">存盘文件</param>
        /// <param name="encoding">编码格式</param>
        public void SaveToFile(string filePath, Encoding encoding)
        {
            //-- 保存文件
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter writer = new StreamWriter(file, encoding))
                {
                    writer.Write(mContent);
                }
            }
        }

      
    }
}
