namespace PerformanceAppraisal
{
    public static class Const
    {
        // 基礎資料夾路徑
        public static readonly string DocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        public static readonly string SourceFolderName = "個人績效考核記錄";

        // 動態生成的檔案路徑 - 適應不同使用者
        // ? "C:\Users\william\Documents\個人績效考核記錄\個人績效考核紀錄A.xlsx"
        public static readonly string SourceFolderPath = Path.Combine(DocumentsPath, SourceFolderName);
        public static readonly string SourceAPath = Path.Combine(SourceFolderPath, "個人績效考核紀錄A.xlsx");
        public static readonly string SourceBPath = Path.Combine(SourceFolderPath, "個人績效考核紀錄B.xlsx");
        public static readonly string SourceCPath = Path.Combine(SourceFolderPath, "個人績效考核紀錄C.xlsx");

        // 主檔案路徑 - 使用相對路徑或設定檔
        public static readonly string MainFilePath = Path.Combine(DocumentsPath, "績效考核.xlsx");

        // 檔案名稱常數
        public const string SourceAFileName = "個人績效考核紀錄A.xlsx";
        public const string SourceBFileName = "個人績效考核紀錄B.xlsx";
        public const string SourceCFileName = "個人績效考核紀錄C.xlsx";
        public const string MainFileName = "績效考核.xlsx";
    }
}