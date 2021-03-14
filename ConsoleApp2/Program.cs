using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using ExcelDataReader;

namespace ConsoleApp2
{
    class Program
    {
        private static void ReadExcelDataSet(DataSet ds, StreamWriter sw)
        { // ds(Excelのデータ) ⇒ sw(ファイルへ出力)
            foreach (DataTable tbl in ds.Tables)
            {
                sw.WriteLine($"TABLE {tbl.TableName}");
                foreach (DataRow row in tbl.Rows)
                {
                    for (var i = 0; i < tbl.Columns.Count; i++)
                    {
                        sw.Write($"{row[i]} ");
                    }
                    sw.WriteLine("");
                }
            }
        }

        // ds(Excelのデータ) ⇒ まとめにあるように、「日付」を見て「仕入金額」を月別に合計して ⇒ sw(ファイルへ出力)
        // 1. シート毎(foreach DataTable)に、(ただし「まとめ」シートは除く)
        // 2. 先頭行から、「日付」と「仕入金額」の列を探し出して、dayとyenに列番号を入れる。
        // 3. 「日付」は、テキストで 「yyyy/mm/dd」形式なので、前の7文字が共通な文字列は同じ月と考える。
        // 4. 月(day)と仕入金額(yen)の構造体の配列を作り、仕入金額(yen)を月毎に合計していく。
        // 5. シート毎に、各月(day)と仕入金額(yen)の合計を出力する。
        // Gitのmessage"自動決算/仕入帳/まとめの代わりを作る/月別の仕入金額合計を作る/SumPurchasing()
        // 月別の仕入金額合計を作るの次は、別のシートを使い、完成マーク毎の仕入金額合計を作る。
        private static void SumPurchasing(DataSet ds, StreamWriter sw) // Purchasing=仕入
        {
            foreach (DataTable tbl in ds.Tables)
            {
                sw.WriteLine($"TABLE {tbl.TableName}");

                // 2. 先頭行から、「日付」と「仕入金額」の列を探し出して、dayColumnとyenColumnに列番号を入れる。

                var TopRow = tbl.Rows[0];
                var dayColumn = -1;
                var yenColumn = -1;
                for (var i = 0; i < tbl.Columns.Count; i++)
                {
                    if(TopRow[i].Equals("日付"))
                    {
                        dayColumn = i;
                        System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(1):dayColumn =<{tbl.TableName}><{dayColumn}>");
					}
                    if(TopRow[i].Equals("仕入金額"))
                    {
                        yenColumn = i;
                        System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(2):yenColumn =<{tbl.TableName}><{yenColumn}>");
					}
                }
                System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(3):dayColumn =<{tbl.TableName}><{dayColumn}>");
                System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(4):yenColumn =<{tbl.TableName}><{yenColumn}>");

                foreach (DataRow row in tbl.Rows)
                {
                    for (var i = 0; i < tbl.Columns.Count; i++)
                    {
                        sw.Write($"{row[i]} ");
                    }
                    sw.WriteLine("");
                }
            }
        }
        private static void ExcelDataRead(StreamWriter sw)
        { // ExcelDataReaderというライブラリを使ってExcelファイルを読む。
            try
            {
                string path = @"C:\develop\Excel";
                System.Collections.Generic.IEnumerable<string> files = Directory.EnumerateFiles(path, "*.xlsx");
                foreach (string filenames in files)
                {
                    if (filenames.Contains('~'))
                    {
                        sw.WriteLine($"ExcelDataRead(1):<{filenames}> エクセルのテンポラリーファイルです。");
                    }
                    else
                    {
                        sw.WriteLine($"ExcelDataRead(2):<{filenames}>");
                        using FileStream stream = File.Open(filenames, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                        DataSet ds = reader.AsDataSet();
                        SumPurchasing(ds, sw);
                    }
                }
            }
            catch (Exception e)
            {
                sw.WriteLine($"ExcelDataRead(3):ERROR<{e.Message}>");
            }
        }
        private static void TextWrite()
        { // テキスト出力を行う。
            string pathname = @"C:\develop" + @"\" + DateTime.Now.ToString("yyyy年MM月dd日HHmmss") + ".txt";
            System.Diagnostics.Debug.WriteLine($"DEBUG:TextWrite(1):pathname =<{pathname}>");
            Console.WriteLine($"TextWrite(1) pathname =<{pathname}>");
            using StreamWriter sw = new StreamWriter(pathname);
            ExcelDataRead(sw);
            sw.Close();
        }
        private static void CopySourceFile()
        {
            string pathnameSource = @"C:\Users\hiroy\source\repos\ConsoleApp1\ConsoleApp1\Program.cs";
            string filename = DateTime.Now.ToString("yyyy年MM月dd日HHmmss") + ".txt";
            string pathnameDestination1 = @"C:\develop\source\" + filename;
            string pathnameDestination2 = @"\Users\hiroy\OneDrive\backup\CSharp\" + filename;
            File.Copy(pathnameSource, pathnameDestination1);
            File.Copy(pathnameSource, pathnameDestination2);
            System.Diagnostics.Debug.WriteLine($"DEBUG:CopySourceFile(1):{pathnameDestination1}");
            System.Diagnostics.Debug.WriteLine($"DEBUG:CopySourceFile(2):{pathnameDestination2}");
            Console.WriteLine(pathnameDestination1);
            Console.WriteLine(pathnameDestination2);
        }
        static void Main(string[] args)
        {
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            System.Diagnostics.Debug.WriteLine($"DEBUG:Main(1):Test debug write");
            Console.WriteLine("Hello World!");
            CopySourceFile();
            TextWrite();
			Console.WriteLine($"{sw.ElapsedMilliseconds}msec");
            System.Diagnostics.Debug.WriteLine($"DEBUG:Main(2):time={sw.ElapsedMilliseconds}msec");
            sw.Stop();
            Console.ReadKey();
        }
    }
}
