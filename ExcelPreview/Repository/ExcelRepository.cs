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
            return GenerateExcelFileSync(data);
        }

        public ExcelFileVM GenerateExcelFileSync(List<ExcelData> data)
        {
            var templatePath = Path.Combine(_webHostEnvironment.ContentRootPath, "Assets", "ExcelFile.xlsx");

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"Excel template not found at: {templatePath}");
            }

            var tempDir = Path.GetTempPath();
            var tempTemplatePath = Path.Combine(tempDir, $"template_{Guid.NewGuid():N}.xlsx");
            var tempOutputPath = Path.Combine(tempDir, $"output_{Guid.NewGuid():N}.xlsx");

            try
            {
                File.Copy(templatePath, tempTemplatePath, true);

                var spreadsheet = new Spreadsheet();
                spreadsheet.LoadFromFile(tempTemplatePath);

                var worksheet = spreadsheet.Workbook.Worksheets[0];
                int currentRow = 3;

                foreach (var record in data)
                {
                    worksheet.Cell(currentRow, 1).Value = record.Id;
                    worksheet.Cell(currentRow, 2).Value = record.Name;

                    if (double.TryParse(record.Value, out double numericValue))
                    {
                        worksheet.Cell(currentRow, 3).Value = numericValue;
                    }
                    else
                    {
                        worksheet.Cell(currentRow, 3).Value = 0;
                    }

                    currentRow++;
                }

                spreadsheet.SaveAsXLSX(tempOutputPath);
                spreadsheet.Dispose();
                spreadsheet = null;

                // Force garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Thread.Sleep(200);

                var fileContent = File.ReadAllBytes(tempOutputPath);
                var fileName = $"ExcelData_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                return new ExcelFileVM
                {
                    FileContent = fileContent,
                    FileName = fileName,
                    ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                };
            }
            finally
            {
                try
                {
                    if (File.Exists(tempTemplatePath))
                        File.Delete(tempTemplatePath);

                    if (File.Exists(tempOutputPath))
                        File.Delete(tempOutputPath);
                }
                catch { }
            }
        }

        // NEW METHOD: Generate Excel and return temporary file path (async)
        public async Task<string> GenerateExcelTempFileAsync()
        {
            var data = await GetAllExcelDataAsync();
            return GenerateExcelTempFileSync(data);
        }

        // NEW METHOD: Generate Excel and return temporary file path (sync)
        public string GenerateExcelTempFileSync(List<ExcelData> data)
        {
            var templatePath = Path.Combine(_webHostEnvironment.ContentRootPath, "Assets", "ExcelFile.xlsx");

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"Excel template not found at: {templatePath}");
            }

            var tempDir = Path.GetTempPath();
            var tempTemplatePath = Path.Combine(tempDir, $"template_{Guid.NewGuid():N}.xlsx");
            var tempOutputPath = Path.Combine(tempDir, $"output_{Guid.NewGuid():N}.xlsx");

            try
            {
                File.Copy(templatePath, tempTemplatePath, true);

                var spreadsheet = new Spreadsheet();
                spreadsheet.LoadFromFile(tempTemplatePath);

                var worksheet = spreadsheet.Workbook.Worksheets[0];
                int currentRow = 3;

                foreach (var record in data)
                {
                    worksheet.Cell(currentRow, 1).Value = record.Id;
                    worksheet.Cell(currentRow, 2).Value = record.Name;

                    if (double.TryParse(record.Value, out double numericValue))
                    {
                        worksheet.Cell(currentRow, 3).Value = numericValue;
                    }
                    else
                    {
                        worksheet.Cell(currentRow, 3).Value = 0;
                    }

                    currentRow++;
                }

                spreadsheet.SaveAsXLSX(tempOutputPath);
                spreadsheet.Dispose();
                spreadsheet = null;

                // Force garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Thread.Sleep(200);

                // Return the temporary file path instead of reading the file content
                return tempOutputPath;
            }
            finally
            {
                try
                {
                    // Only delete the template file, keep the output file
                    if (File.Exists(tempTemplatePath))
                        File.Delete(tempTemplatePath);

                    // NOTE: tempOutputPath is NOT deleted here since we're returning it
                    // The caller is responsible for cleaning up this file when done
                }
                catch { }
            }
        }
    }
}