using ClosedXML.Excel;

namespace PerformanceAppraisal
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

            try
            {
                Console.WriteLine($"開始讀取 Excel 檔案: {filePath}");

                using (var workbook = new XLWorkbook(filePath))
                {
                    // 讀取第一個工作表
                    var worksheet = workbook.Worksheet(1);
                    var data = ReadWorksheetData(worksheet);

                    Console.WriteLine($"Excel 檔案讀取完成，共 {data.Count} 筆記錄");
                    return data;
                }
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

            try
            {
                Console.WriteLine($"開始讀取工作表: {worksheetName}");

                using (var workbook = new XLWorkbook(filePath))
                {
                    IXLWorksheet worksheet = GetWorksheet(workbook, worksheetName);
                    var data = ReadWorksheetData(worksheet);

                    Console.WriteLine($"工作表 '{worksheet.Name}' 讀取完成，共 {data.Count} 筆記錄");
                    return data;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"讀取工作表時發生錯誤: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 獲取 Excel 檔案中所有工作表的名稱
        /// </summary>
        /// <param name="filePath">Excel 檔案路徑</param>
        /// <returns>工作表名稱列表</returns>
        public List<string> GetWorksheetNames(string filePath)
        {
            if (!ValidateExcelFile(filePath))
            {
                throw new ArgumentException("無效的 Excel 檔案");
            }

            var worksheetNames = new List<string>();

            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        worksheetNames.Add(worksheet.Name);
                    }
                }

                Console.WriteLine($"檔案包含 {worksheetNames.Count} 個工作表: {string.Join(", ", worksheetNames)}");
                return worksheetNames;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"獲取工作表名稱時發生錯誤: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 讀取 Excel 檔案並返回每個工作表的資料
        /// </summary>
        /// <param name="filePath">Excel 檔案路徑</param>
        /// <returns>包含所有工作表資料的字典</returns>
        public Dictionary<string, List<Dictionary<string, object>>> ReadAllWorksheets(string filePath)
        {
            if (!ValidateExcelFile(filePath))
            {
                throw new ArgumentException("無效的 Excel 檔案");
            }

            var allData = new Dictionary<string, List<Dictionary<string, object>>>();

            try
            {
                Console.WriteLine($"開始讀取所有工作表: {filePath}");

                using (var workbook = new XLWorkbook(filePath))
                {
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        Console.WriteLine($"正在讀取工作表: {worksheet.Name}");
                        var worksheetData = ReadWorksheetData(worksheet);
                        allData[worksheet.Name] = worksheetData;
                    }
                }

                Console.WriteLine($"所有工作表讀取完成，共 {allData.Count} 個工作表");
                return allData;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"讀取所有工作表時發生錯誤: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 從工作表對象讀取資料的輔助方法
        /// </summary>
        /// <param name="worksheet">工作表對象</param>
        /// <returns>工作表資料</returns>
        private List<Dictionary<string, object>> ReadWorksheetData(IXLWorksheet worksheet)
        {
            var data = new List<Dictionary<string, object>>();
            var range = worksheet.RangeUsed();

            if (range == null)
            {
                Console.WriteLine($"工作表 '{worksheet.Name}' 沒有資料");
                return data;
            }

            // 讀取標題行
            var headers = ReadHeaders(range);
            Console.WriteLine($"找到 {headers.Count} 個欄位: {string.Join(", ", headers)}");

            // 讀取資料行 (從第二行開始)
            for (int row = 2; row <= range.RowCount(); row++)
            {
                var rowData = ReadRow(range, row, headers);
                if (rowData != null)
                {
                    data.Add(rowData);
                }
            }

            return data;
        }

        /// <summary>
        /// 根據名稱或索引找到工作表
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="worksheetName">工作表名稱</param>
        /// <returns>工作表對象</returns>
        private IXLWorksheet GetWorksheet(XLWorkbook workbook, string worksheetName)
        {
            try
            {
                return workbook.Worksheet(worksheetName);
            }
            catch
            {
                Console.WriteLine($"找不到工作表 '{worksheetName}'，使用第一個工作表");
                return workbook.Worksheet(1);
            }
        }

        /// <summary>
        /// 讀取標題行
        /// </summary>
        /// <param name="range">資料範圍</param>
        /// <returns>標題列表</returns>
        private List<string> ReadHeaders(IXLRange range)
        {
            var headers = new List<string>();
            var headerRow = range.Row(1);

            for (int col = 1; col <= range.ColumnCount(); col++)
            {
                var headerValue = headerRow.Cell(col).Value.ToString();
                headers.Add(string.IsNullOrWhiteSpace(headerValue) ? $"Column{col}" : headerValue);
            }

            return headers;
        }

        /// <summary>
        /// 讀取單一資料行
        /// </summary>
        /// <param name="range">資料範圍</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="headers">標題列表</param>
        /// <returns>行資料，空行時返回 null</returns>
        private Dictionary<string, object>? ReadRow(IXLRange range, int rowIndex, List<string> headers)
        {
            var rowData = new Dictionary<string, object>();
            bool hasData = false;

            for (int col = 1; col <= range.ColumnCount(); col++)
            {
                var cellValue = range.Cell(rowIndex, col).Value;
                var header = headers[col - 1];

                // 處理不同的資料類型
                object value = ConvertCellValue(cellValue);
                rowData[header] = value;

                // 檢查是否有非空資料
                if (!string.IsNullOrWhiteSpace(value?.ToString()))
                {
                    hasData = true;
                }
            }

            // 只返回有資料的行
            return hasData ? rowData : null;
        }

        /// <summary>
        /// 轉換儲存格值為適當的 .NET 類型
        /// </summary>
        /// <param name="cellValue">儲存格值</param>
        /// <returns>轉換後的值</returns>
        private object ConvertCellValue(XLCellValue cellValue)
        {
            return cellValue.Type switch
            {
                XLDataType.Text => cellValue.GetText(),
                XLDataType.Number => cellValue.GetNumber(),
                XLDataType.DateTime => cellValue.GetDateTime(),
                XLDataType.Boolean => cellValue.GetBoolean(),
                XLDataType.TimeSpan => cellValue.GetTimeSpan(),
                _ => cellValue.ToString()
            };
        }
    }
}