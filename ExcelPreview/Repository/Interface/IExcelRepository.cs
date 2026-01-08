using Backend.Models;
using Backend.ViewModel;

namespace ExcelPreview.Repository.Interface
{
    public interface IExcelRepository
    {
        Task<ExcelFileVM> GenerateExcelFileAsync();
        ExcelFileVM GenerateExcelFileSync(List<ExcelData> data);
        Task<List<ExcelData>> GetAllExcelDataAsync();
        List<ExcelData> GetAllExcelData();

        // New method to generate Excel and return temp file path
        Task<string> GenerateExcelTempFileAsync();
        string GenerateExcelTempFileSync(List<ExcelData> data);

        // New PDF generation methods
        byte[] GenerateExcelAsPDF(List<ExcelData> data);
        string GenerateExcelAsPDFTempFile(List<ExcelData> data);
        Task<string> GenerateExcelAsPDFTempFileAsync();

        // Convert Excel to Pdf Only
        string ExcelToPdf();
    }
}