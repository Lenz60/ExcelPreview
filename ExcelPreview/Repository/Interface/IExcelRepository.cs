using Backend.Models;
using Backend.ViewModel;

namespace ExcelPreview.Repository.Interface
{
    public interface IExcelRepository
    {
        Task<ExcelFileVM> GenerateExcelFileAsync();
        Task<ExcelFileVM> GenerateExcelFileAsync(List<ExcelData> data);
        Task<List<ExcelData>> GetAllExcelDataAsync();
        List<ExcelData> GetAllExcelData(); // Keep your existing method
    }
}
