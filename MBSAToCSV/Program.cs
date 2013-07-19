using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace MBSAToCSV
{
    class Program
    {
        /// <summary>
        /// エントリポイント
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            var options = new Options();

            if (!CommandLine.Parser.Default.ParseArguments(args, options))
            {
                Console.WriteLine("コマンドラインの引数が不正です。");
                return;
            }


            if (options.Download)
            {
                Console.WriteLine("cab2ファイルのダウンロードを開始します。");
                if (Download())
                {
                    Console.WriteLine("cab2ファイルのダウンロードが完了しました。");
                }
                else
                {
                    Console.WriteLine("ダウンロードに失敗したため、処理を終了します。");
                    return;
                }
            }


            var fileName = string.IsNullOrEmpty(options.FileName) ? @"mbsareport.xml" : options.FileName + ".xml";

            Console.WriteLine("MBSAのレポート出力を開始します。");
            if (RunMBSA(fileName))
            {
                Console.WriteLine("MBSAのレポートを出力しました。");
            }
            else
            {
                Console.WriteLine("MBSAのレポート出力に失敗しましたので、処理を終了します。");
            }


            Console.WriteLine("CSVファイル出力を開始します。");
            if (ToCSV(fileName))
            {
                Console.WriteLine("CSVファイルを出力しました。");
            }
            else
            {
                Console.WriteLine("CSVファイル出力に失敗しました。");
            }
        }


        /// <summary>
        /// Microsoftのサイトより、wsusscn2.cabファイルをダウンロードする
        /// </summary>
        /// <returns>true - ダウンロードに成功</returns>
        static bool Download()
        {
            using (var wc = new System.Net.WebClient())
            {
                wc.DownloadFile("http://go.microsoft.com/fwlink/?LinkId=76054",
                                System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + @"\" + "wsusscn2.cab");
            }

            return true;
        }


        /// <summary>
        /// コマンドプロンプトによるMBSAの実行
        /// </summary>
        /// <param name="fileName">レポートファイルのファイル名(拡張子：xml)</param>
        /// <returns>true - MBSAの実行に成功</returns>
        static bool RunMBSA(string fileName)
        {
            using (var p = new System.Diagnostics.Process())
            {
                p.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec");
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.Arguments = @"/c mbsacli.exe /xmlout /unicode /nd /nvc /wi /catalog wsusscn2.cab > " + fileName;
                p.Start();
                p.WaitForExit();
            }

            return true;
        }


        /// <summary>
        /// MBSAから出力されたXMLファイルを、CSV化
        /// </summary>
        /// <param name="fileName">レポートファイルのファイル名(拡張子：xml)</param>
        /// <returns>true - CSV化に成功</returns>
        static bool ToCSV(string fileName)
        {
            var xmlFilePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + @"\" + fileName;

            var csv = System.Xml.Linq.XElement.Load(xmlFilePath).Elements("Check").Elements("Detail").Elements("UpdateData")
                             .Where(a => a.Attribute("IsInstalled").Value == "false")
                             .Select(b => string.Format("{0},{1},{2},{3}{4}",
                                                         (string)b.Attribute("ID"),
                                                         "KB" + (string)b.Attribute("KBID"),
                                                         (string)b.Element("Title"),
                                                         (string)b.Element("References").Element("InformationURL"),
                                                         Environment.NewLine
                                                         ))
                             .Aggregate(new StringBuilder(), (sb, s) => sb.Append(s), sb => sb.ToString());


            string csvPath = xmlFilePath.Substring(0, xmlFilePath.Length - 3) + "csv";
            System.Text.Encoding enc = System.Text.Encoding.GetEncoding("shift_jis");

            using (System.IO.StreamWriter sw = new System.IO.StreamWriter(csvPath, false, enc))
            {
                // ヘッダーの先頭にID列を表記したい場合はダブルコーテーションで囲む
                // See: http://pasofaq.jp/office/excel/sylk.htm
                string header = "\"ID\",KB番号,タイトル,情報へのURL";
                sw.WriteLine(header);

                sw.WriteLine(csv);
            }


            return true;
        }
    }
}
