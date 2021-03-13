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
        {
            foreach (DataTable tbl in ds.Tables)
            {
                sw.WriteLine("TABLE {0}", tbl.TableName);
                foreach (DataRow row in tbl.Rows)
                {
                    for (var i = 0; i < tbl.Columns.Count; i++)
                    {
                        sw.Write("{0} ", row[i]);
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
                        sw.WriteLine("message=0<{0}> エクセルのテンポラリーファイルです。", filenames);
                    }
                    else
                    {
                        sw.WriteLine("message=1<{0}>", filenames);
                        using FileStream stream = File.Open(filenames, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                        DataSet ds = reader.AsDataSet();
                        ReadExcelDataSet(ds, sw);
                    }
                }
            }
            catch (Exception e)
            {
                sw.WriteLine("ERROR in ExcelDataRead()<{0}>", e.Message);
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
            Console.WriteLine(pathnameDestination1);
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
