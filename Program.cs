namespace PerformanceAppraisal
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Performance Appraisal System Starting...");

            try
            {
                // 建立 Excel 讀取器實例
                var excelReader = new ExcelReader();

                // 建立 Excel 寫入器實例
                var excelWriter = new ExcelWriter();

                // 示範用法
                Console.WriteLine("Excel Reader and Writer initialized successfully!");

                // 示範讀取不同路徑的 Excel 檔案
                DemonstrateExcelReading(excelReader);

                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
        }

        private static void DemonstrateExcelReading(ExcelReader reader)
        {
            Console.WriteLine("\n=== Excel 讀取示範 ===");

            // 方法 1: 使用 Const 中定義的路徑讀取檔案
            try
            {
                Console.WriteLine("\n1. 檢查來源檔案存在狀態:");
                Console.WriteLine($"   檔案 A: {Const.SourceAPath}");
                Console.WriteLine($"   檔案 B: {Const.SourceBPath}");
                Console.WriteLine($"   檔案 C: {Const.SourceCPath}");

                // 讀取檔案 A (如果存在)
                if (System.IO.File.Exists(Const.SourceAPath))
                {
                    Console.WriteLine("\n2. 讀取檔案 A:");
                    
                    // 先查看有哪些工作表
                    var worksheetNames = reader.GetWorksheetNames(Const.SourceAPath);
                    
                    // 讀取第一個工作表
                    if (worksheetNames.Count > 0)
                    {
                        var data = reader.ReadWorksheet(Const.SourceAPath, worksheetNames[0]);
                        Console.WriteLine($"   讀取到 {data.Count} 筆記錄");
                        
                        // 顯示前幾筆資料作為範例
                        DisplaySampleData(data, 3);
                    }

                    // 或者讀取所有工作表
                    Console.WriteLine("\n3. 讀取所有工作表:");
                    var allWorksheets = reader.ReadAllWorksheets(Const.SourceAPath);
                    foreach (var worksheet in allWorksheets)
                    {
                        Console.WriteLine($"   工作表 '{worksheet.Key}': {worksheet.Value.Count} 筆記錄");
                    }
                }
                else
                {
                    Console.WriteLine($"\n檔案 A 不存在: {Const.SourceAPath}");
                    Console.WriteLine("請確認檔案路徑是否正確或建立測試檔案");
                }

                // 方法 2: 使用自定義路徑讀取檔案
                Console.WriteLine("\n4. 您也可以指定任何 Excel 檔案路徑:");
                Console.WriteLine("   var data = reader.ReadExcelFile(@\"C:\\your\\custom\\path.xlsx\");");
                Console.WriteLine("   var specificSheet = reader.ReadWorksheet(@\"C:\\your\\path.xlsx\", \"Sheet1\");");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"讀取示範時發生錯誤: {ex.Message}");
            }
        }

        private static void DisplaySampleData(List<Dictionary<string, object>> data, int maxRows)
        {
            if (data.Count == 0)
            {
                Console.WriteLine("   沒有資料可顯示");
                return;
            }

            Console.WriteLine("   範例資料:");
            var sampleCount = Math.Min(maxRows, data.Count);
            
            // 顯示欄位名稱
            var firstRow = data[0];
            Console.WriteLine($"   欄位: {string.Join(", ", firstRow.Keys)}");
            
            // 顯示前幾筆資料
            for (int i = 0; i < sampleCount; i++)
            {
                var values = data[i].Values.Select(v => v?.ToString() ?? "").ToArray();
                Console.WriteLine($"   資料 {i + 1}: {string.Join(" | ", values)}");
            }
            
            if (data.Count > maxRows)
            {
                Console.WriteLine($"   ... 還有 {data.Count - maxRows} 筆資料");
            }
        }
    }
}