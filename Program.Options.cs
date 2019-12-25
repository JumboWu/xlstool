using System;
using CommandLine;
using CommandLine.Text;

namespace xlstool {
    partial class Program {
        /// <summary>
        /// 命令行参数定义
        /// </summary>
        internal sealed class Options {
            public Options() {
                this.TableName = "Sheet1#StructName";
                this.HeaderRows = 4;
                this.Encoding = "utf8-nobom";
                this.Lowcase = false;
                this.ExportArray = false;
                this.Platform = "client";//client server and empty is both two
            }

            [Option('i', "excel", Required = true, HelpText = "input excel file path.")]
            public string ExcelPath
            {
                get;
                set;
            }

            [Option('t', "table", Required = true, HelpText = "input excel table sheet name.")]
            public string TableName
            {
                get;
                set;
            }

            
            [Option('c', "code", Required = false, HelpText = "json sql cs lua go ts tsv csv")]
            public string Code {
                get;
                set;
            }

            [Option('h', "header", Required = false, DefaultValue = 4, HelpText = "number lines in sheet as header.")]
            public int HeaderRows {
                get;
                set;
            }

            [Option("encoding", Required = false, DefaultValue = "utf8-nobom", HelpText = "export file encoding.")]
            public string Encoding {
                get;
                set;
            }

            [Option("lowcase", Required = false, DefaultValue = false, HelpText = "convert filed name to lowcase.")]
            public bool Lowcase {
                get;
                set;
            }

            [Option("array", Required = false, DefaultValue = false, HelpText = "export as array, otherwise as dict object.")]
            public bool ExportArray {
                get;
                set;
            }

            [Option('p', "plat", Required = false, DefaultValue = "client",  HelpText = "export only client or server or both two. ")]
            public string Platform{
                get;
                set;
            }

            [Option('o', "out", Required = false, DefaultValue = "", HelpText = "export file to path. ")]
            public string ExportPath
            {
                get;
                set;
            }

            [HelpOption('h', "help")]
            public string GetUsage()
            {
                return HelpText.AutoBuild(this, current => HelpText.DefaultParsingErrorsHandler(this, current));
            }
        }
    }
}
