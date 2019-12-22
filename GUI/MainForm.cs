using FastColoredTextBoxNS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace xlstool.GUI {

    /// <summary>
    /// 主窗口
    /// </summary>
    public partial class MainForm : Form {
        // Excel导入数据管理
        private DataManager mDataMgr;
        private string mCurrentXlsx;

        // 支持语法高亮的文本框
        private FastColoredTextBox mJsonTextBox;
        private FastColoredTextBox mSQLTextBox;
        private FastColoredTextBox mCodeTextBox;
        private FastColoredTextBox mLuaTextBox;
        private FastColoredTextBox mGoTextBox;
        private FastColoredTextBox mTSTextBox;

        // 文本框的样式
        private TextStyle mBrownStyle = new TextStyle(Brushes.Brown, null, FontStyle.Regular);
        private TextStyle mMagentaStyle = new TextStyle(Brushes.Magenta, null, FontStyle.Regular);
        private TextStyle mGreenStyle = new TextStyle(Brushes.Green, null, FontStyle.Regular);

        // 导出数据相关的按钮，方便整体Enable/Disable
        private List<ToolStripButton> mExportButtonList;

        /// <summary>
        /// 构造函数，初始化控件初值；创建文本框
        /// </summary>
        public MainForm() {
            InitializeComponent();

            //-- syntax highlight text box
            mJsonTextBox = createTextBoxInTab(this.tabPageJSON);
            mJsonTextBox.Language = Language.Custom;
            mJsonTextBox.TextChanged += new EventHandler<TextChangedEventArgs>(this.jsonTextChanged);

            mSQLTextBox = createTextBoxInTab(this.tabPageSQL);
            mSQLTextBox.Language = Language.SQL;

            mCodeTextBox = createTextBoxInTab(this.tabPageCode);
            mCodeTextBox.Language = Language.CSharp;

            mLuaTextBox = createTextBoxInTab(this.tabPageLua);
            mLuaTextBox.Language = Language.Lua;

            mGoTextBox = createTextBoxInTab(this.tabPageGo);
            mGoTextBox.Language = Language.Custom;

            mTSTextBox = createTextBoxInTab(this.tabPageTS);
            mTSTextBox.Language = Language.JS;

            //-- componet init states
            this.comboBoxType.SelectedIndex = 0;
            this.comboBoxLowcase.SelectedIndex = 1;
            this.comboBoxHeader.SelectedIndex = 2;
            this.comboBoxHeader.Enabled = false;//Excel定义号结构后,不需要操作了

            this.comboBoxSheet.Items.Clear();
            this.comboBoxEncoding.Items.Clear();
            this.comboBoxEncoding.Items.Add("utf8-nobom");
            foreach (EncodingInfo ei in Encoding.GetEncodings()) {
                Encoding e = ei.GetEncoding();
                this.comboBoxEncoding.Items.Add(e.HeaderName);
            }
            this.comboBoxEncoding.SelectedIndex = 0;
            this.comboBoxEncoding.Enabled = false;

            //-- button list
            mExportButtonList = new List<ToolStripButton>();
            mExportButtonList.Add(this.btnCopyJSON);
            mExportButtonList.Add(this.btnSaveJson);
            mExportButtonList.Add(this.btnSaveSQL);
            mExportButtonList.Add(this.btnSaveCSharp);
            mExportButtonList.Add(this.btnSaveLua);
            mExportButtonList.Add(this.btnSaveGo);
            mExportButtonList.Add(this.btnSaveTS);
            enableExportButtons(false);

            //-- data manager
            mDataMgr = new DataManager();
            this.btnReimport.Enabled = false;
        }

        /// <summary>
        /// 设置导出相关的按钮是否可用
        /// </summary>
        /// <param name="enable">是否可用</param>
        private void enableExportButtons(bool enable) {
            foreach (var btn in mExportButtonList)
                btn.Enabled = enable;
        }

        /// <summary>
        /// 在一个TabPage中创建Text Box
        /// </summary>
        /// <param name="tab">TabPage容器控件</param>
        /// <returns>新建的Text Box控件</returns>
        private FastColoredTextBox createTextBoxInTab(TabPage tab) {
            FastColoredTextBox textBox = new FastColoredTextBox();
            textBox.Dock = DockStyle.Fill;
            textBox.Font = new Font("Microsoft YaHei", 11F);
            tab.Controls.Add(textBox);
            return textBox;
        }

        /// <summary>
        /// 设置Json文本高亮格式
        /// </summary>
        private void jsonTextChanged(object sender, TextChangedEventArgs e) {
            e.ChangedRange.ClearStyle(mBrownStyle, mMagentaStyle, mGreenStyle);
            //allow to collapse brackets block
            e.ChangedRange.SetFoldingMarkers("{", "}");
            //string highlighting
            e.ChangedRange.SetStyle(mBrownStyle, @"""""|@""""|''|@"".*?""|(?<!@)(?<range>"".*?[^\\]"")|'.*?[^\\]'");
            //number highlighting
            e.ChangedRange.SetStyle(mGreenStyle, @"\b\d+[\.]?\d*([eE]\-?\d+)?[lLdDfF]?\b|\b0x[a-fA-F\d]+\b");
        }

        /// <summary>
        /// 使用BackgroundWorker加载Excel文件，使用UI中的Options设置
        /// </summary>
        /// <param name="path">Excel文件路径</param>
        private void loadExcelAsync(string path, bool currentXlsx = false) {

            mCurrentXlsx = path;

            //-- update ui
            this.btnReimport.Enabled = true;
            this.labelExcelFile.Text = path;
            enableExportButtons(false);

            this.statusLabel.IsLink = false;
            this.statusLabel.Text = "Loading Excel ...";

            //-- load options from ui
            Program.Options options = mDataMgr.Options;
            if (!currentXlsx)
                options = new Program.Options();

            options.ExcelPath = path;
            options.ExportArray = this.comboBoxType.SelectedIndex == 0;
            options.Encoding = this.comboBoxEncoding.SelectedText;
            options.Lowcase = this.comboBoxLowcase.SelectedIndex == 0;
            options.HeaderRows = int.Parse(this.comboBoxHeader.Text);
            options.Platform = this.comboBoxPlat.SelectedItem as string;
            

            //-- start import
            this.backgroundWorker.RunWorkerAsync(options);
        }

        /// <summary>
        /// 接受Excel拖放事件
        /// </summary>
        private void panelExcelDropBox_DragDrop(object sender, DragEventArgs e) {
            string[] dropData = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (dropData != null) {
                this.loadExcelAsync(dropData[0], false);
            }
        }

        /// <summary>
        /// 判断拖放对象是否是一个.xlsx文件
        /// </summary>
        private void panelExcelDropBox_DragEnter(object sender, DragEventArgs e) {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                string[] dropData = (string[])e.Data.GetData(DataFormats.FileDrop, false);
                if (dropData != null && dropData.Length > 0) {
                    string szPath = dropData[0];
                    string szExt = System.IO.Path.GetExtension(szPath);
                    szExt = szExt.ToLower();
                    if (szExt == ".xlsx") {
                        e.Effect = DragDropEffects.All;
                        return;
                    }
                }
            }//end of if(file)
            e.Effect = DragDropEffects.None;
        }

        /// <summary>
        /// 执行实际的Excel加载
        /// </summary>
        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e) {
            lock (this.mDataMgr) {
                this.mDataMgr.loadExcel((Program.Options)e.Argument);
            }
        }

        /// <summary>
        /// Excel加载完成
        /// </summary>
        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {

            // 判断错误信息
            if (e.Error != null) {
                showStatus(e.Error.Message, Color.Red);
                return;
            }

            // 更新UI
            lock (this.mDataMgr) {
                this.statusLabel.IsLink = false;
                this.statusLabel.Text = "Load completed.";

                mJsonTextBox.Text = mDataMgr.JsonContext;
                mSQLTextBox.Text = mDataMgr.SQLContext;
                mCodeTextBox.Text = mDataMgr.CSharpCode;
                mLuaTextBox.Text = mDataMgr.LuaCode;
                mGoTextBox.Text = mDataMgr.GoCode;
                mTSTextBox.Text = mDataMgr.TSCode;

                List<string> sheetList = mDataMgr.SheetList;
                if(sheetList != null)
                {
                    this.comboBoxSheet.Items.Clear();
                    for (int i = 0; i < sheetList.Count; i++)
                    {
                        if (!sheetList[i].Contains("#"))
                            continue;
                        this.comboBoxSheet.Items.Add(sheetList[i]);
                        if (sheetList[i] == mDataMgr.Options.TableName)
                            this.comboBoxSheet.SelectedIndex = i;
                    }
                }

                enableExportButtons(true);
            }
        }

        /// <summary>
        /// 工具栏按钮：Import Excel
        /// </summary>
        private void btnImportExcel_Click(object sender, EventArgs e) {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.RestoreDirectory = true;
            dlg.Filter = "Excel File(*.xlsx)|*.xlsx";
            if (dlg.ShowDialog() == DialogResult.OK) {
                this.loadExcelAsync(dlg.FileName, false);
            }
        }

        /// <summary>
        /// 点击状态栏链接
        /// </summary>
        private void statusLabel_Click(object sender, EventArgs e) {
            if (this.statusLabel.IsLink) {
                System.Diagnostics.Process.Start(this.statusLabel.Text);
            }
        }

        /// <summary>
        /// 保存导出文件
        /// </summary>
        private void saveToFile(CodeType type, string filter, string fileName = "") {

            try {
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.RestoreDirectory = true;
                dlg.Filter = filter;
                dlg.FileName = fileName;
                if (dlg.ShowDialog() == DialogResult.OK) {
                    lock (mDataMgr) {
                        switch (type) {
                            case CodeType.Json:
                                mDataMgr.saveJson(dlg.FileName);
                                break;
                            case CodeType.Sql:
                                mDataMgr.saveSQL(dlg.FileName);
                                break;
                            case CodeType.CSharp:
                                mDataMgr.saveCS(dlg.FileName);
                                break;
                            case CodeType.Lua:
                                mDataMgr.saveLua(dlg.FileName);
                                break;
                            case CodeType.Go:
                                mDataMgr.saveGo(dlg.FileName);
                                break;
                            case CodeType.TypeScript:
                                mDataMgr.saveTS(dlg.FileName);
                                break;
                        }
                    }
                    showStatus(string.Format("{0} saved!", dlg.FileName), Color.Black);
                }// end of if
            }
            catch (Exception ex) {
                showStatus(ex.Message, Color.Red);
            }
        }

        /// <summary>
        /// 工具栏按钮：Save Json
        /// </summary>
        private void btnSaveJson_Click(object sender, EventArgs e) {
            saveToFile(CodeType.Json, "Json File(*.json)|*.json", Utils.GetStructName(mDataMgr.Options.TableName));
        }

        /// <summary>
        /// 工具栏按钮：Save SQL
        /// </summary>
        private void btnSaveSQL_Click(object sender, EventArgs e) {
            saveToFile(CodeType.Sql, "SQL File(*.sql)|*.sql", Utils.GetStructName(mDataMgr.Options.TableName));
        }

        /// <summary>
        /// 工具栏按钮：Save C#
        /// </summary>
        private void btnSaveCSharp_Click(object sender, EventArgs e) {
            saveToFile(CodeType.CSharp, "C# Code File(*.cs)|*.cs", Utils.GetStructName(mDataMgr.Options.TableName));
        }

        /// <summary>
        /// 工具栏按钮：Copy Json
        /// </summary>
        private void btnCopyJSON_Click(object sender, EventArgs e) {
            lock (mDataMgr) {
                Clipboard.SetText(mDataMgr.JsonContext);
                showStatus("Json text copyed to clipboard.", Color.Black);
            }
        }

        /// <summary>
        /// 设置状态栏信息
        /// </summary>
        /// <param name="szMessage">信息文字</param>
        /// <param name="color">信息颜色</param>
        private void showStatus(string szMessage, Color color) {
            this.statusLabel.Text = szMessage;
            this.statusLabel.ForeColor = color;
            this.statusLabel.IsLink = false;
        }

        /// <summary>
        /// 配置项变更之后，手动重新导入xlsx文件
        /// </summary>
        private void btnReimport_Click(object sender, EventArgs e) {
            if (!string.IsNullOrEmpty(mCurrentXlsx)) {
                loadExcelAsync(mCurrentXlsx, true);
            }
        }

        /// <summary>
        /// 切换sheet,自动导入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBoxSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (mDataMgr == null || mDataMgr.Options == null)
                return;

            int index = this.comboBoxSheet.SelectedIndex;
            if (mDataMgr.Options.TableName != mDataMgr.SheetList[index])
            {
                mDataMgr.Options.TableName = mDataMgr.SheetList[index];
                if (!string.IsNullOrEmpty(mCurrentXlsx))
                {
                    loadExcelAsync(mCurrentXlsx, true);
                }
            }
                
        }

		/// <summary>
        /// 切换plat，自动导入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBoxPlat_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (mDataMgr == null || mDataMgr.Options == null)
                return;

            string plat = this.comboBoxPlat.SelectedItem as string;
            if (mDataMgr.Options.Platform != plat)
            {
                mDataMgr.Options.Platform = plat;
                if (!string.IsNullOrEmpty(mCurrentXlsx))
                {
                    loadExcelAsync(mCurrentXlsx, true);
                }
            }
            
        }
		
        private void btnSaveLua_Click(object sender, EventArgs e)
        {
            saveToFile(CodeType.Lua, "Lua Code File(*.lua)|*.lua", Utils.GetStructName(mDataMgr.Options.TableName));
        }

        private void btnSaveGo_Click(object sender, EventArgs e)
        {
            saveToFile(CodeType.Go, "Go Code File(*.go)|*.go", Utils.GetStructName(mDataMgr.Options.TableName));
        }

        private void btnSaveTS_Click(object sender, EventArgs e)
        {
            saveToFile(CodeType.TypeScript, "TypeScript Code File(*.ts)|*.ts", Utils.GetStructName(mDataMgr.Options.TableName));
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://jumbowu.github.io");
        }
    }
}
