using Backend.Context;
using Backend.Models;
using Backend.ViewModel;
using Bytescout.Spreadsheet;
using Bytescout.Spreadsheet.Constants;
using ExcelPreview.Repository.Interface;
using Microsoft.AspNetCore.Hosting;
using Microsoft.EntityFrameworkCore;

namespace ExcelPreview.Repository
{
    public class ExcelRepository : IExcelRepository
    {
        private readonly ExcelPreviewContext _context;
        private readonly IWebHostEnvironment _webHostEnvironment;

        public ExcelRepository(ExcelPreviewContext context, IWebHostEnvironment webHostEnvironment)
        {
            _context = context;
            _webHostEnvironment = webHostEnvironment;
        }

        public async Task<List<ExcelData>> GetAllExcelDataAsync()
        {
            return await _context.ExcelDatas.ToListAsync();
        }

        public List<ExcelData> GetAllExcelData()
        {
            return _context.ExcelDatas.ToList();
        }

        public async Task<ExcelFileVM> GenerateExcelFileAsync()
        {
            var data = await GetAllExcelDataAsync();
            return await GenerateExcelFileAsync(data);
        }
        public async Task<ExcelFileVM> GenerateExcelFileAsync(List<ExcelData> data)
        {
            return await Task.Run(() =>
            {
                // Get the path to the existing Excel template
                var templatePath = Path.Combine(_webHostEnvironment.ContentRootPath, "Assets", "ExcelFile.xlsx");

                // Check if template file exists
                if (!File.Exists(templatePath))
                {
                    throw new FileNotFoundException($"Excel template not found at: {templatePath}");
                }

                // Load the existing Excel file
                var spreadsheet = new Spreadsheet();
                spreadsheet.LoadFromFile(templatePath);

                // Get the first worksheet
                var worksheet = spreadsheet.Workbook.Worksheets[0];

                // Iterate through all data records using foreach
                int currentRow = 3; // Start at row 4 (0-based indexing)

                foreach (var record in data)
                {
                    // Place Name in column C (column index 2)
                    worksheet.Cell(currentRow, 2).Value = record.Name;

                    // Parse Value to numeric (double) before placing in column D (column index 3)
                    if (double.TryParse(record.Value, out double numericValue))
                    {
                        worksheet.Cell(currentRow, 3).Value = numericValue;
                    }
                    else
                    {
                        // If parsing fails, set to 0 or keep original string
                        worksheet.Cell(currentRow, 3).Value = 0; // or record.Value for original string
                    }

                    currentRow++; // Move to next row for next record
                }

                // Generate temporary file path for the modified file
                var tempPath = Path.GetTempPath();
                var fileName = $"ExcelData_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                var tempFilePath = Path.Combine(tempPath, fileName);

                // Save the modified file to temporary location
                spreadsheet.SaveAsXLSX(tempFilePath);

                // Read the file content for response
                var fileContent = File.ReadAllBytes(tempFilePath);

                // Clean up temporary file
                File.Delete(tempFilePath);

                return new ExcelFileVM
                {
                    FileContent = fileContent,
                    FileName = fileName,
                    ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                };
            });
        }
    }
}
