namespace ExcelReader
{
    /// <summary>
    /// Excel 檔案讀取器
    /// 負責從 Excel 檔案中讀取資料
    /// </summary>
    public class ExcelReader
    {
        /// <summary>
        /// 讀取 Excel 檔案
        /// </summary>
        /// <param name="filePath">Excel 檔案路徑</param>
        /// <returns>讀取的資料</returns>
        public List<Dictionary<string, object>> ReadExcelFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Excel 檔案不存在: {filePath}");
            }

            var data = new List<Dictionary<string, object>>();

            try
            {
                Console.WriteLine($"開始讀取 Excel 檔案: {filePath}");

                // TODO: 實作 Excel 讀取邏輯
                // 建議使用 EPPlus 或 ClosedXML 套件來處理 Excel 檔案

                Console.WriteLine("Excel 檔案讀取完成");
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"讀取 Excel 檔案時發生錯誤: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 驗證 Excel 檔案格式
        /// </summary>
        /// <param name="filePath">Excel 檔案路徑</param>
        /// <returns>檔案是否有效</returns>
        public bool ValidateExcelFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    return false;
                }

                var extension = Path.GetExtension(filePath).ToLower();
                return extension == ".xlsx" || extension == ".xls";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"驗證 Excel 檔案時發生錯誤: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 讀取指定工作表的資料
        /// </summary>
        /// <param name="filePath">Excel 檔案路徑</param>
        /// <param name="worksheetName">工作表名稱</param>
        /// <returns>工作表資料</returns>
        public List<Dictionary<string, object>> ReadWorksheet(string filePath, string worksheetName)
        {
            if (!ValidateExcelFile(filePath))
            {
                throw new ArgumentException("無效的 Excel 檔案");
            }

            var data = new List<Dictionary<string, object>>();

            try
            {
                Console.WriteLine($"開始讀取工作表: {worksheetName}");

                // TODO: 實作特定工作表讀取邏輯

                Console.WriteLine($"工作表 '{worksheetName}' 讀取完成");
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"讀取工作表時發生錯誤: {ex.Message}");
                throw;
            }
        }
    }
}