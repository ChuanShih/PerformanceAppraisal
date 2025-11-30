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

                // TODO: 在此添加您的業務邏輯

                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
        }
    }
}