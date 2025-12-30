using Backend.Context;
using Backend.Models;
using Backend.ViewModel;
using ExcelPreview.Repository.Interface;
using Microsoft.AspNetCore.Hosting;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

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

            byte[] fileContent = null;
            ExcelPackage template = null;
            ExcelPackage package = null;
            FileStream templateStream = null;

            try
            {
                // Set EPPlus license context (required for EPPlus 5.0+)
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Use FileStream to load template with leaveOpen = true for better control
                templateStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                template = new ExcelPackage(templateStream);
                package = new ExcelPackage();

                // Get the first worksheet from template
                ExcelWorksheet worksheet = PrepareWorksheet(template, package);

                // Place data and perform calculations
                PopulateWorksheetAndCalculate(worksheet, package, data);

                // Get file content as byte array
                fileContent = package.GetAsByteArray();

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
                // Dispose resources in proper order
                try
                {
                    package?.Dispose();
                }
                catch { }

                try
                {
                    template?.Dispose();
                }
                catch { }

                try
                {
                    templateStream?.Dispose();
                }
                catch { }

                // Perform controlled garbage collection after disposal
                if (fileContent != null)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect(); // Second collection to clean up finalizer queue
                }
            }
        }

        private ExcelWorksheet PrepareWorksheet(ExcelPackage template, ExcelPackage package)
        {
            ExcelWorksheet worksheet;

            if (template.Workbook.Worksheets.Count > 0)
            {
                // Copy the first worksheet from template
                var templateWorksheet = template.Workbook.Worksheets[0];
                worksheet = package.Workbook.Worksheets.Add(templateWorksheet.Name, templateWorksheet);
            }
            else
            {
                // Create new worksheet if template doesn't have one
                worksheet = package.Workbook.Worksheets.Add("Sheet1");
            }

            return worksheet;
        }


        private void PopulateWorksheetAndCalculate(ExcelWorksheet worksheet, ExcelPackage package, List<ExcelData> data)
        {
            int currentRow = 4; // Starting from row 4

            foreach (var record in data)
            {
                // Place ID in column B (2)
                worksheet.Cells[currentRow, 2].Value = record.Id;

                // Place Name in column C (3)
                worksheet.Cells[currentRow, 3].Value = record.Name;

                // Place Value in column D (4), converting to numeric if possible
                if (double.TryParse(record.Value, out double numericValue))
                {
                    worksheet.Cells[currentRow, 4].Value = numericValue;
                }
                else
                {
                    worksheet.Cells[currentRow, 4].Value = 0;
                }

                currentRow++;
            }

            // Force calculation to ensure charts and formulas are updated
            package.Workbook.Calculate();

            // Allow calculation to complete
            Thread.Sleep(100);
        }

        // NEW METHOD: Generate Excel and return temporary file path (async)
        public async Task<string> GenerateExcelTempFileAsync()
        {
            var data = await GetAllExcelDataAsync();
            return GenerateExcelTempFileSync(data);
        }

        // NEW METHOD: Generate Excel and return temporary file path (sync) - Updated to use EPPlus with FileStream
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

            ExcelPackage template = null;
            ExcelPackage package = null;
            FileStream tempTemplateStream = null;
            FileStream outputStream = null;

            try
            {
                // Set EPPlus license context
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Copy template to temp location
                File.Copy(templatePath, tempTemplatePath, true);

                // Use FileStream to load template with proper stream management
                tempTemplateStream = new FileStream(tempTemplatePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                template = new ExcelPackage(tempTemplateStream);
                package = new ExcelPackage();

                // Get the first worksheet from template
                ExcelWorksheet worksheet = PrepareWorksheet(template, package);

                // Place data and perform calculations (with longer wait for temp file operations)
                PopulateWorksheetAndCalculateForTempFile(worksheet, package, data);

                // Create output stream with FileStream for better control
                outputStream = new FileStream(tempOutputPath, FileMode.Create, FileAccess.Write, FileShare.None);
                package.SaveAs(outputStream);

                // Ensure all data is written to disk
                outputStream.Flush();
                outputStream.Close();
                outputStream.Dispose();
                outputStream = null;

                // Allow file system to complete operations
                Thread.Sleep(300);

                return tempOutputPath;
            }
            finally
            {
                // Dispose resources in proper order
                try
                {
                    outputStream?.Dispose();
                }
                catch { }

                try
                {
                    package?.Dispose();
                }
                catch { }

                try
                {
                    template?.Dispose();
                }
                catch { }

                try
                {
                    tempTemplateStream?.Dispose();
                }
                catch { }

                // Clean up temporary template file
                try
                {
                    if (File.Exists(tempTemplatePath))
                        File.Delete(tempTemplatePath);
                }
                catch { }

                // Perform controlled garbage collection after all disposals
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect(); // Second collection to ensure cleanup

                // NOTE: tempOutputPath is NOT deleted here since we're returning it
                // The caller is responsible for cleaning up this file when done
            }
        }

        private void PopulateWorksheetAndCalculateForTempFile(ExcelWorksheet worksheet, ExcelPackage package, List<ExcelData> data)
        {
            int currentRow = 4; // Starting from row 4

            foreach (var record in data)
            {
                // Place ID in column B (2)
                worksheet.Cells[currentRow, 2].Value = record.Id;

                // Place Name in column C (3)
                worksheet.Cells[currentRow, 3].Value = record.Name;

                // Place Value in column D (4), converting to numeric if possible
                if (double.TryParse(record.Value, out double numericValue))
                {
                    worksheet.Cells[currentRow, 4].Value = numericValue;
                }
                else
                {
                    worksheet.Cells[currentRow, 4].Value = 0;
                }

                currentRow++;
            }

            // Force calculation for temp file operations
            package.Workbook.Calculate();

            // Allow calculation to complete - longer wait for temp file operations
            Thread.Sleep(200);
        }
    }
}