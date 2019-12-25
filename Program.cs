using System;
using System.IO;
using System.Data;
using System.Text;
using System.Windows.Forms;

using Excel;

namespace xlstool {
    /// <summary>
    /// 应用程序
    /// </summary>
    sealed partial class Program {
        /// <summary>
        /// 应用程序入口
        /// </summary>
        /// <param name="args">命令行参数</param>
        [STAThread]
        static void Main(string[] args) {
            if (args.Length <= 0) {
                //-- GUI MODE ----------------------------------------------------------
                Console.WriteLine("Launch xlstool GUI Mode...");
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new GUI.MainForm());
            }
            else {
                //-- COMMAND LINE MODE -------------------------------------------------
                
                //-- 分析命令行参数
                var options = new Options();
         
                var parser = new CommandLine.Parser(with => with.HelpWriter = Console.Error);

                if (parser.ParseArgumentsStrict(args, options, () => Environment.Exit(-1))) {
                    //-- 执行导出操作
                    try {
                        DateTime startTime = DateTime.Now;
                        Run(options);
                        //-- 程序计时
                        DateTime endTime = DateTime.Now;
                        TimeSpan dur = endTime - startTime;
                        Console.WriteLine(
                            string.Format("[{0}]：\tConversion complete in [{1}ms].",
                            Path.GetFileName(options.ExcelPath),
                            dur.TotalMilliseconds)
                            );
                    }
                    catch (Exception exp) {
                        Console.WriteLine("Error: " + exp.Message);
                    }
                }
            }// end of else
        }

        /// <summary>
        /// 根据命令行参数，执行Excel数据导出工作
        /// </summary>
        /// <param name="options">命令行参数</param>
        private static void Run(Options options) {
            string excelPath = options.ExcelPath;
            string excelName = Path.GetFileNameWithoutExtension(options.ExcelPath);
            int header = options.HeaderRows;

            // 加载Excel文件
            //using (FileStream excelFile = File.Open(excelPath, FileMode.Open, FileAccess.Read))
            using (FileStream excelFile = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Reading from a OpenXml Excel file (2007 format; *.xlsx)
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelFile);

                // The result of each spreadsheet will be created in the result.Tables
                excelReader.IsFirstRowAsColumnNames = false;
                DataSet book = excelReader.AsDataSet();

                // 数据检测
                if (book.Tables.Count < 1) {
                    throw new Exception("Excel file is empty: " + excelPath);
                }

                // 取得数据
                DataTable sheet = Utils.GetDataTable(book, options.TableName);
                if (sheet == null)
                {
                    sheet = book.Tables[0];
                    options.TableName = sheet.TableName;
                }
                    

                if (sheet.Rows.Count <= 0) {
                    throw new Exception("Excel Sheet is empty: " + excelPath);
                }

                //-- 确定编码
                Encoding cd = new UTF8Encoding(false);
                if (options.Encoding != "utf8-nobom") {
                    foreach (EncodingInfo ei in Encoding.GetEncodings()) {
                        Encoding e = ei.GetEncoding();
                        if (e.HeaderName == options.Encoding) {
                            cd = e;
                            break;
                        }
                    }
                }


                if (options.Code != "tsv" && options.Code != "csv")
                {
                    //不满足的清掉
                    Utils.IsKV = (sheet.Rows[0][0].ToString() == Utils.KV);
                    if (!Utils.IsKV)
                    {
                        DataRow _platRow = sheet.Rows[2];
                        string _plat;
                        for (int col = sheet.Columns.Count - 1; col >= 0; col--)
                        {
                            _plat = _platRow[col].ToString();
                            if (_plat == string.Empty || _plat == options.Platform)
                                continue;

                            //标记这个不需要导出
                            sheet.Columns.RemoveAt(col);
                        }
                    }
                    else
                    {
                        string _plat;
                        for (int row = sheet.Rows.Count - 1; row >= 2; row--)
                        {
                            _plat = sheet.Rows[row][3].ToString();
                            if (_plat == string.Empty || _plat == options.Platform)
                                continue;

                            //标记这个不需要导出
                            sheet.Rows.RemoveAt(row);
                        }
                    }
                }
              

                IExporter exporter = null;
                switch (options.Code)
                {
                    case "json":
                        //-- 导出JSON文件
                        exporter = new JsonExporter(sheet, header, options.Lowcase, options.ExportArray);
                        break;
                    case "sql":
                        //-- 导出SQL文件
                        exporter = new SQLExporter(sheet, header, options.Lowcase);
                        break;
                    case "cs":
                        //-- 生成C#定义文件
                        exporter = new CSDefineGenerator(excelName, sheet, options.Lowcase);
                        break;
                    case "go":
                        //-- 生成Go文件
                        exporter = new GoDefineGenerator(excelName, sheet, options.Lowcase);
                        break;
                    case "lua":
                        //-- 生成Lua文件
                        exporter = new LuaExporter(excelName, sheet, header, options.Lowcase);
                        break;
                    case "ts":
                        exporter = new TSDefineGenerator(excelName, sheet, options.Lowcase);
                        break;
                        //--导出csv、tsv
                    case "tsv":
                    case "csv":
                        exporter = new Exporter(sheet, options.Code);
                        break;
                }

                if (exporter != null)
                    exporter.SaveToFile(options.ExportPath, cd);


            }
        }

    }
}
