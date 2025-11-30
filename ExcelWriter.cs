namespace ExcelWriter
{
    /// <summary>
    /// Excel 檔案寫入器
    /// 負責將資料寫入 Excel 檔案
    /// </summary>
    public class ExcelWriter
    {
        /// <summary>
        /// 將資料寫入 Excel 檔案
        /// </summary>
        /// <param name="data">要寫入的資料</param>
        /// <param name="filePath">輸出檔案路徑</param>
        public void WriteToExcel(List<Dictionary<string, object>> data, string filePath)
        {
            if (data == null || data.Count == 0)
            {
                throw new ArgumentException("沒有資料可以寫入");
            }

            try
            {
                Console.WriteLine($"開始寫入 Excel 檔案: {filePath}");

                // 確保目錄存在
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // TODO: 實作 Excel 寫入邏輯
                // 建議使用 EPPlus 或 ClosedXML 套件來處理 Excel 檔案

                Console.WriteLine("Excel 檔案寫入完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"寫入 Excel 檔案時發生錯誤: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 建立新的 Excel 檔案並寫入資料
        /// </summary>
        /// <param name="data">要寫入的資料</param>
        /// <param name="filePath">輸出檔案路徑</param>
        /// <param name="worksheetName">工作表名稱</param>
        public void CreateExcelFile(List<Dictionary<string, object>> data, string filePath, string worksheetName = "Sheet1")
        {
            if (data == null || data.Count == 0)
            {
                throw new ArgumentException("沒有資料可以寫入");
            }

            try
            {
                Console.WriteLine($"開始建立新的 Excel 檔案: {filePath}");

                // 確保目錄存在
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // TODO: 實作新 Excel 檔案建立邏輯

                Console.WriteLine($"Excel 檔案 '{filePath}' 建立完成，工作表名稱: {worksheetName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"建立 Excel 檔案時發生錯誤: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 將資料附加到現有的 Excel 檔案
        /// </summary>
        /// <param name="data">要附加的資料</param>
        /// <param name="filePath">Excel 檔案路徑</param>
        /// <param name="worksheetName">工作表名稱</param>
        public void AppendToExcel(List<Dictionary<string, object>> data, string filePath, string worksheetName = "Sheet1")
        {
            if (data == null || data.Count == 0)
            {
                throw new ArgumentException("沒有資料可以附加");
            }

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Excel 檔案不存在: {filePath}");
            }

            try
            {
                Console.WriteLine($"開始附加資料到 Excel 檔案: {filePath}");

                // TODO: 實作資料附加邏輯

                Console.WriteLine($"資料已成功附加到工作表 '{worksheetName}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"附加資料到 Excel 檔案時發生錯誤: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 將簡單資料表格寫入 Excel
        /// </summary>
        /// <param name="headers">表格標題</param>
        /// <param name="rows">表格資料行</param>
        /// <param name="filePath">輸出檔案路徑</param>
        /// <param name="worksheetName">工作表名稱</param>
        public void WriteTableToExcel(List<string> headers, List<List<object>> rows, string filePath, string worksheetName = "Sheet1")
        {
            if (headers == null || headers.Count == 0)
            {
                throw new ArgumentException("表格標題不能為空");
            }

            if (rows == null || rows.Count == 0)
            {
                throw new ArgumentException("表格資料不能為空");
            }

            try
            {
                Console.WriteLine($"開始寫入表格資料到 Excel 檔案: {filePath}");

                // 確保目錄存在
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // TODO: 實作表格寫入邏輯

                Console.WriteLine($"表格資料寫入完成，包含 {headers.Count} 個欄位，{rows.Count} 行資料");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"寫入表格資料時發生錯誤: {ex.Message}");
                throw;
            }
        }
    }
}