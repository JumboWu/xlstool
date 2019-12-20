using Excel;
using System;
using System.Data;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace xlstool.GUI {

    /// <summary>
    /// 为GUI模式提供的整体数据管理
    /// </summary>
    class DataManager {
      
        // 数据导入设置
        private Program.Options mOptions;
        private Encoding mEncoding;

        // 导出数据
        private JsonExporter mJson;
        private SQLExporter mSQL;
        private CSDefineGenerator mCSharp;
        private LuaExporter mLua;
        private GoDefineGenerator mGo;

        public Program.Options Options { get { return mOptions; } }

        private List<string> mSheetList;
        public List<string> SheetList
        {
            get { return mSheetList; }
        }

        /// <summary>
        /// 导出的Json文本
        /// </summary>
        public string JsonContext {
            get {
                if (mJson != null)
                    return mJson.context;
                else
                    return "";
            }
        }

        /// <summary>
        /// 导出的SQL文本
        /// </summary>
        public string SQLContext {
            get {
                if (mSQL != null)
                    return mSQL.structSQL + mSQL.contentSQL;
                else
                    return "";
            }
        }

        /// <summary>
        /// 导出的C#代码
        /// </summary>
        public string CSharpCode {
            get {
                if (mCSharp != null)
                    return mCSharp.code;
                else
                    return "";
            }
        }

        /// <summary>
        /// 导出的Lua代码
        /// </summary>
        public string LuaCode
        {
            get
            {
                if (mLua != null)
                    return mLua.code;
                else
                    return "";
            }
        }

        /// <summary>
        /// 导出的Go代码
        /// </summary>
        public string GoCode
        {
            get
            {
                if (mGo != null)
                    return mGo.code;
                else
                    return "";
            }
        }

        /// <summary>
        /// 保存Json
        /// </summary>
        /// <param name="filePath">保存路径</param>
        public void saveJson(string filePath) {
            if (mJson != null) {
                mJson.SaveToFile(filePath, mEncoding);
            }
        }

        /// <summary>
        /// 保存SQL
        /// </summary>
        /// <param name="filePath">保存路径</param>
        public void saveSQL(string filePath) {
            if (mSQL != null) {
                mSQL.SaveToFile(filePath, mEncoding);
            }
        }

        /// <summary>
        /// 保存C#代码
        /// </summary>
        /// <param name="filePath">保存路径</param>
        public void saveCS(string filePath) {
            if (mCSharp != null) {
                mCSharp.SaveToFile(filePath, mEncoding);
            }
        }
        /// <summary>
        /// 保存Lua代码
        /// </summary>
        /// <param name="filePath">保存路径</param>
        public void saveLua(string filePath)
        {
            if (mLua != null)
            {
                mLua.SaveToFile(filePath, mEncoding);
            }
        }

        /// <summary>
        /// 保存Go代码
        /// </summary>
        /// <param name="filePath">保存路径</param>
        public void saveGo(string filePath)
        {
            if (mGo != null)
            {
                mGo.SaveToFile(filePath, mEncoding);
            }
        }


        /// <summary>
        /// 加载Excel文件
        /// </summary>
        /// <param name="options">导入设置</param>
        public void loadExcel(Program.Options options) {
            mOptions = options;
            string excelPath = options.ExcelPath;
            string excelName = Path.GetFileNameWithoutExtension(excelPath);
            int header = options.HeaderRows;

            // 加载Excel文件 
            //using (FileStream excelFile = File.Open(excelPath, FileMode.Open, FileAccess.Read)) //这种方式会导致，文件被其他进程打开，不能重新打开问题，不便于工作
            using (FileStream excelFile = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Reading from a OpenXml Excel file (2007 format; *.xlsx)
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelFile);

                // The result of each spreadsheet will be created in the result.Tables
                excelReader.IsFirstRowAsColumnNames = false;
                DataSet book = excelReader.AsDataSet();

                // 数据检测
                if (book.Tables.Count < 1)
                {
                    throw new Exception("Excel file is empty: " + excelPath);
                }

                mSheetList = new List<string>(book.Tables.Count);
                for (int i = 0; i < book.Tables.Count; i++)
                    mSheetList.Add(book.Tables[i].TableName);
                // 取得数据
                DataTable sheet = Utils.GetDataTable(book, options.TableName);
                if (sheet == null)
                {
                    sheet = book.Tables[0];
                    options.TableName = sheet.TableName;
                }

                if (sheet.Rows.Count <= 0)
                {
                    throw new Exception("Excel Sheet is empty: " + excelPath);
                }

                //-- 确定编码
                Encoding cd = new UTF8Encoding(false);
                if (options.Encoding != "utf8-nobom")
                {
                    foreach (EncodingInfo ei in Encoding.GetEncodings())
                    {
                        Encoding e = ei.GetEncoding();
                        if (e.HeaderName == options.Encoding)
                        {
                            cd = e;
                            break;
                        }
                    }
                }
                mEncoding = cd;

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
               

                //-- 导出JSON
                mJson = new JsonExporter(sheet, header, options.Lowcase, options.ExportArray);

                //-- 导出SQL
                mSQL = new SQLExporter(sheet, header, options.Lowcase);

                //-- 生成C#定义代码
                mCSharp = new CSDefineGenerator(excelName, sheet, options.Lowcase);

                //--生成Lua
                mLua = new LuaExporter(excelName, sheet, header, options.Lowcase);

                //--生成Go
                mGo = new GoDefineGenerator(excelName, sheet, options.Lowcase);
            }
        }
     
    }
}
