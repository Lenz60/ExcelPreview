namespace Backend.ViewModel
{
    public class ExcelFileVM
    {
        public byte[] FileContent { get; set; }
        public string FileName { get; set; }
        public string ContentType { get; set; } = "application/vnd.ms-excel";
        public long FileSize => FileContent?.Length ?? 0;
        public string? TempFilePath { get; set; } // Optional: to track temp file location
    }
}
