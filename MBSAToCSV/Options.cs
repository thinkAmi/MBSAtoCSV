using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CommandLine;
using CommandLine.Text;

namespace MBSAToCSV
{
    class Options
    {
        [Option('d', "Download", DefaultValue=false, HelpText="cab2ファイルをMSサイトよりダウンロード")]
        public bool Download { get; set; }


        [Option('f', "MBSAFileName", DefaultValue = "", HelpText = "MBSAのレポートとcsv化した時のファイル名(拡張子不要)")]
        public string FileName { get; set; }


        [HelpOption('h', "help", HelpText="ヘルプの表示")]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this, (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }
}
