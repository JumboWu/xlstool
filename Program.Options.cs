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

            [Option("table", Required = true, HelpText = "input excel table sheet name.")]
            public string TableName
            {
                get;
                set;
            }

            [Option("excel", Required = true, HelpText = "input excel file path.")]
            public string ExcelPath {
                get;
                set;
            }

            [Option("json", Required = false, HelpText = "export json file path.")]
            public string JsonPath {
                get;
                set;
            }

            [Option("sql", Required = false, HelpText = "export SQL file path.")]
            public string SQLPath {
                get;
                set;
            }

            [Option("cs", Required = false, HelpText = "export C# data struct code file path.")]
            public string CSharpPath {
                get;
                set;
            }

            [Option("lua", Required = false, HelpText = "export Lua code file path.")]
            public string LuaPath
            {
                get;
                set;
            }

            [Option("go", Required = false, HelpText = "export Go data struct code file path.")]
            public string GoPath
            {
                get;
                set;
            }

            [Option("ts", Required = false, HelpText = "export typescript data struct code file path.")]
            public string TSPath
            {
                get;
                set;
            }


            [Option("header", Required = true, HelpText = "number lines in sheet as header.")]
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
            [Option("plat", Required = false, DefaultValue = "client",  HelpText = "export only client or server or both two. ")]
            public string Platform{
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
