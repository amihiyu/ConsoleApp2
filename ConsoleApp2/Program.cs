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

        // {日付(年＋月):string,仕入金額合計:int}の配列 // 文字列がいいか、年月に特定の配列箇所を割り当てるか？
        // 高速になるので、年月に特定の配列箇所を割り当てる。2000年1月～2039年12月で40x12=480配列。

        const int START_YEAR = 2000;
        const int END_YEAR = 2039;
        const int MONTHS = (END_YEAR - START_YEAR + 1)*12;      // 2000年1月～2039年12月で40x12=480配列。
        const int SHEETS = 256;                                 // シートの数の最大値。仮に256個。

		private static int String2MONTHS(string s){
            int length = s.Length;
            if(length < 9){
                System.Diagnostics.Debug.WriteLine($"DEBUG:StringMONTHS(1):ERROR:too small length=<{length}>,s<{s}>");
                return -1;
			}
            DateTime dt = DateTime.Parse(s);
            if(dt.Year < START_YEAR){
                System.Diagnostics.Debug.WriteLine($"DEBUG:StringMONTHS(2):ERROR:dt.Year is not in range=<{START_YEAR}<={dt.Year}<={END_YEAR}>");
                return -1;
			}
            int months = (dt.Year - 2000) * 12 + dt.Month;
            return months;
		}
        private class Purchasing
        {
            int nSheet;                                     // 現在のシート数
            string[] sheet = new string[SHEETS];            // シート名
			int[,] sumYen = new int[SHEETS,MONTHS];         // 仕入金額合計(月別の配列)

			public Purchasing(){ // 初期化
                nSheet = 0;
			}
            public void NewSheet(string sheet){ // 新たなシートを登録
                if(0 <= nSheet & nSheet < SHEETS){
                    this.sheet[nSheet] = sheet;
                    nSheet++;
				}else{ 
                    System.Diagnostics.Debug.WriteLine($"DEBUG:ERROR:Purchasing(1) 配列範囲外 nSheet=<{nSheet}>");
                }
			}
            public void AddSum(int ym, int yen){ // 現在のシートの年月の所の仕入金額を増額
                if(0 <= nSheet && nSheet < SHEETS && 0 <= ym && ym < MONTHS){
                    this.sumYen[nSheet - 1,ym] += yen;
				}else{
                    if(!(0 <= nSheet & nSheet < SHEETS)){
                        System.Diagnostics.Debug.WriteLine($"DEBUG:ERROR:Purchasing(2) 配列範囲外 nSheet=<{nSheet}>");
					}
                    if(!(0 <= ym & ym < MONTHS)){
                        System.Diagnostics.Debug.WriteLine($"DEBUG:ERROR:Purchasing(3) 配列範囲外 ym=<{ym}>");
					}
				}
			}
            public string GetNowSheetName(){
                if(!(1 <= nSheet & nSheet <= SHEETS)){
                    System.Diagnostics.Debug.WriteLine($"DEBUG:ERROR:Purchasing(4) 配列範囲外 nSheet=<{nSheet}>");
                    return null;
                }else{
                    return sheet[nSheet-1];
                }
                
			}
            public void DisplayNowSheet(){
                if(!(0 <= nSheet & nSheet < SHEETS)){
                    System.Diagnostics.Debug.WriteLine($"DEBUG:ERROR:DisplyNowSheet(1) 配列範囲外 nSheet=<{nSheet}>");
                }else{
                    string s = GetNowSheetName();
                    if(s != null){
                        System.Diagnostics.Debug.WriteLine($"DEBUG:DisplyNowSheet(2) シート数<{nSheet}>,シート名=<{s}>");
                        for(int ym = 0; ym < MONTHS; ym++){
                            int sy = sumYen[nSheet-1, ym];
                            if(sy != 0){
                                System.Diagnostics.Debug.WriteLine($"DEBUG:DisplyNowSheet(3) nSheet=<{nSheet}> SheetName=<{s}> ym=<{ym / 12},{ym % 12}>, sy=<{sy}>");
							}
						}
					}
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
        private static void SumPurchasing(Purchasing purchase, DataSet ds, StreamWriter sw) // Purchasing=仕入
        {
            foreach (DataTable tbl in ds.Tables)
            {
                sw.WriteLine($"TABLE {tbl.TableName}");

                // 2. 先頭行から、「日付」と「仕入金額」の列を探し出して、dayColumnとyenColumnに列番号を入れる。

                var TopRow = tbl.Rows[0];
                int dayColumn = -1;
                int yenColumn = -1;
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

                // 3. 「日付」は、テキストで 「yyyy/mm/dd」形式なので、前の7文字が共通な文字列は同じ月と考える。

                if(dayColumn >= 0 && yenColumn >= 0){
//                    System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(5):");
                    purchase.NewSheet(tbl.TableName);
  //                  System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(6):tbl.Rows.Count=<{tbl.Rows.Count}>");
                    for(int j = 1; j < tbl.Rows.Count; j++){
    //                    System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(6-1):j=<{j}>");
                        DataRow row = tbl.Rows[j];
      //                  System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(6-2):row=<{row}>");
                        if(row[dayColumn] == null){
        //                     System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(7):tbl.Rows.Count=<{tbl.Rows.Count}>");
						}else{ 
          //                   System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(7-1):tbl.Rows.Count=<{tbl.Rows.Count}>");
                            string dayString = row[dayColumn].ToString();
            //                 System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(7-2):tbl.Rows.Count=<{tbl.Rows.Count}>");
                            if(dayString.Length < 10){
              //                  System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(8):tbl.Rows.Count=<{tbl.Rows.Count}>");
						    }else{
                //                System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(8-1):tbl.Rows.Count=<{tbl.Rows.Count}>");
                                int ym = String2MONTHS(dayString);
                  //              System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(8-2):tbl.Rows.Count=<{tbl.Rows.Count}>");
                                if(row[yenColumn] == null){
                    //                 System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(9):tbl.Rows.Count=<{tbl.Rows.Count}>");
						        }else{
                      //               System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(9-1):tbl.Rows.Count=<{tbl.Rows.Count}>");
                                    string yenString = row[yenColumn].ToString();
                        //             System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(9-2):tbl.Rows.Count=<{tbl.Rows.Count}>");
                                    if(yenString == null){
                          //              System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(10):tbl.Rows.Count=<{tbl.Rows.Count}>");
						            }else{ 
                            //            System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(10-1):yenString=<{yenString}>");
                                        try{
                                            int yen;
                                            double dyen;
                                            if(int.TryParse(yenString, out yen)){
                                                if(ym > 0 & yen != 0){
                                                    purchase.AddSum(ym, yen);
				    		                    }
											}else if(double.TryParse(yenString, out dyen)){
                                                yen = Convert.ToInt32(dyen);
                                                if(ym > 0 & yen != 0){
                                                    purchase.AddSum(ym, yen);
				    		                    }
											}else{

											}
                                        }catch(Exception e){
                                            System.Diagnostics.Debug.WriteLine($"DEBUG:SumPurchasing(11):yenString=<{yenString}> e=<{e}>");
										}
                                    }
                                }
                                // dayColumnやyenColumnに情報が無い場合、例外がスローされるみたい。
                            }
                        }
                    }
                    purchase.DisplayNowSheet();
				}
            }
        }
        private static void AnalyzeExcelData(StreamWriter sw)
        { // ExcelDataReaderというライブラリを使ってExcelファイルを読む。
            Purchasing purchase = new Purchasing();
            try
            {
                string path = @"C:\develop\Excel";
                System.Collections.Generic.IEnumerable<string> files = Directory.EnumerateFiles(path, "*.xlsx");
                foreach (string filenames in files)
                {
                    if (filenames.Contains('~'))
                    {
                        sw.WriteLine($"AnalyzeExcelData(1):<{filenames}> エクセルのテンポラリーファイルです。");
                    }
                    else
                    {
                        sw.WriteLine($"ExcelDataRead(2):<{filenames}>");
                        using FileStream stream = File.Open(filenames, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                        DataSet ds = reader.AsDataSet();
                        System.Diagnostics.Debug.WriteLine($"DEBUG:AnalyzeExcelData(1): befor SumPurchasing()");
                        SumPurchasing(purchase, ds, sw);
                        System.Diagnostics.Debug.WriteLine($"DEBUG:AnalyzeExcelData(2): after SumPurchasing()");
                    }
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine($"DEBUG:AnalyzeExcelData(3):e=<{e}>");
                sw.WriteLine($"AnalyzeExcelData(3):ERROR<{e.Message}>");
            }
        }
        private static void TextWrite()
        { // テキスト出力を行う。
            string pathname = @"C:\develop" + @"\" + DateTime.Now.ToString("yyyy年MM月dd日HHmmss") + ".txt";
            System.Diagnostics.Debug.WriteLine($"DEBUG:TextWrite(1):pathname =<{pathname}>");
            Console.WriteLine($"TextWrite(1) pathname =<{pathname}>");
            using StreamWriter sw = new StreamWriter(pathname);
            AnalyzeExcelData(sw);
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
