using Backend.Context;
using Backend.Models;
using Backend.ViewModel;
using Bytescout.Spreadsheet;
using Bytescout.Spreadsheet.Constants;
using ExcelPreview.Repository.Interface;
using GemBox.Spreadsheet;
using Microsoft.AspNetCore.Hosting;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Drawing.Printing;
using ExcelWorksheet = OfficeOpenXml.ExcelWorksheet;

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

        public byte[] GenerateExcelAsPDF(List<ExcelData> data)
        {
            var templatePath = Path.Combine(_webHostEnvironment.ContentRootPath, "Assets", "ExcelFile.xlsx");

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"Excel template not found at: {templatePath}");
            }

            byte[] pdfContent = null;
            ExcelPackage template = null;
            ExcelPackage package = null;
            FileStream templateStream = null;

            // Generate unique file names for temp files
            var tempExcelPath = Path.Combine(Path.GetTempPath(), $"temp_excel_{Guid.NewGuid():N}.xlsx");
            var tempPdfPath = Path.Combine(Path.GetTempPath(), $"temp_pdf_{Guid.NewGuid():N}.pdf");

            try
            {
                // Set EPPlus license context
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Use FileStream to load template with proper control
                templateStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                template = new ExcelPackage(templateStream);
                package = new ExcelPackage();

                // Reuse existing composable functions
                ExcelWorksheet worksheet = PrepareWorksheet(template, package);
                PopulateWorksheetAndCalculate(worksheet, package, data);

                // Save Excel file temporarily
                package.SaveAs(new FileInfo(tempExcelPath));

                // Wait for file to be completely written
                Thread.Sleep(200);

                // Convert to PDF (this will also clean up the temp Excel file)
                ConvertExcelToPDF(tempExcelPath, tempPdfPath);

                // Verify PDF was created and read content
                if (!File.Exists(tempPdfPath) || new FileInfo(tempPdfPath).Length == 0)
                {
                    throw new InvalidOperationException("PDF conversion failed - output file is missing or empty");
                }

                pdfContent = File.ReadAllBytes(tempPdfPath);

                return pdfContent;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error generating PDF: {ex.Message}", ex);
            }
            finally
            {
                // Dispose EPPlus resources in proper order
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

                // Clean up ALL temporary files
                var tempFilesToClean = new[] { tempExcelPath, tempPdfPath };

                foreach (var tempFile in tempFilesToClean)
                {
                    try
                    {
                        if (File.Exists(tempFile))
                        {
                            File.Delete(tempFile);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Failed to delete temporary file {tempFile}: {ex.Message}");
                    }
                }

                // Perform comprehensive garbage collection after all cleanup
                if (pdfContent != null)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect(); // Second collection to clean up finalizer queue

                    // Additional cleanup for large objects
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        private void ConvertExcelToPDF(string excelPath, string pdfPath)
        {
            GemBox.Spreadsheet.ExcelFile workbook = null;

            try
            {
                // Set license before any GemBox operations
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

                // Load the Excel file
                workbook = GemBox.Spreadsheet.ExcelFile.Load(excelPath);

                // Ensure formulas and charts are calculated
                workbook.Calculate();

                // Enhanced worksheet configuration for charts to prevent cropping
                if (workbook.Worksheets.Count > 0)
                {
                    var worksheet = workbook.Worksheets[0];

                    // Set optimal print settings for charts beside tables
                    worksheet.PrintOptions.PaperType = PaperType.A3; // Large paper for complex layouts
                    worksheet.PrintOptions.Portrait = false; // Wide orientation

                    // Key settings to prevent cropping:
                    worksheet.PrintOptions.FitWorksheetWidthToPages = 1; // Fit content to 1 page wide
                    worksheet.PrintOptions.FitWorksheetHeightToPages = 1; // Allow multiple pages tall if needed

                    // Minimal margins for maximum content area
                    worksheet.PrintOptions.LeftMargin = 0.1;
                    worksheet.PrintOptions.RightMargin = 0.1;
                    worksheet.PrintOptions.TopMargin = 0.1;
                    worksheet.PrintOptions.BottomMargin = 0.1;

                    // Scale settings - let GemBox auto-adjust
                    worksheet.PrintOptions.AutomaticPageBreakScalingFactor = 100; // Start with 100%

                    // Center content horizontally
                    worksheet.PrintOptions.HorizontalCentered = false;
                    worksheet.PrintOptions.VerticalCentered = false; // Don't center vertically to prevent cropping
                }

                // Simple PDF options - the real control is in worksheet.PrintOptions
                var pdfOptions = new PdfSaveOptions()
                {
                    SelectionType = SelectionType.EntireFile
                };

                // Save as PDF
                workbook.Save(pdfPath, pdfOptions);

                // Verify PDF was created successfully
                if (!File.Exists(pdfPath))
                {
                    throw new InvalidOperationException("PDF file was not created successfully");
                }

                // Allow file operations to complete
                Thread.Sleep(300);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to convert Excel to PDF: {ex.Message}", ex);
            }
            finally
            {
                // Clean up workbook reference
                workbook = null;

                // Clean up temporary Excel file immediately after PDF creation
                try
                {
                    if (File.Exists(excelPath))
                    {
                        File.Delete(excelPath);
                    }
                }
                catch (Exception ex)
                {
                    // Log but don't throw - cleanup failure shouldn't break the process
                    Console.WriteLine($"Warning: Failed to delete temporary Excel file {excelPath}: {ex.Message}");
                }

                // Perform controlled garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect(); // Second collection to clean up finalizer queue
            }
        }


        /// <summary>
        /// Auto-fits columns to their content width with EPPlus
        /// </summary>
        /// <param name="worksheet">EPPlus ExcelWorksheet</param>
        private void AutoFitColumns(ExcelWorksheet worksheet)
        {
            try
            {
                // Method 1: Auto-fit all columns that have data
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Method 2: Alternative - Auto-fit specific columns with min/max width
                // worksheet.Column(2).AutoFit(5, 50); // Column B: min 5, max 50
                // worksheet.Column(3).AutoFit(10, 100); // Column C: min 10, max 100  
                // worksheet.Column(4).AutoFit(8, 30); // Column D: min 8, max 30

                // Method 3: Alternative - Auto-fit with global min/max for all columns
                // worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns(5, 100);
            }
            catch (Exception ex)
            {
                // Log error but don't fail the entire operation
                Console.WriteLine($"Warning: Failed to auto-fit columns: {ex.Message}");

                // Fallback: Set reasonable default widths
                worksheet.Column(2).Width = 10; // ID column
                worksheet.Column(3).Width = 25; // Name column
                worksheet.Column(4).Width = 15; // Value column
            }
        }

        public async Task<string> GenerateExcelAsPDFTempFileAsync()
        {
            var data = await GetAllExcelDataAsync();
            return GenerateExcelAsPDFTempFile(data);
        }

        public string GenerateExcelAsPDFTempFile(List<ExcelData> data)
        {
            var templatePath = Path.Combine(_webHostEnvironment.ContentRootPath, "Assets", "ExcelFile.xlsx");

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"Excel template not found at: {templatePath}");
            }

            ExcelPackage template = null;
            ExcelPackage package = null;
            FileStream templateStream = null;

            var tempPdfPath = Path.Combine(Path.GetTempPath(), $"pdf_{Guid.NewGuid():N}.pdf");
            var tempExcelPath = Path.Combine(Path.GetTempPath(), $"temp_excel_{Guid.NewGuid():N}.xlsx");

            try
            {
                // Set EPPlus license context
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Use FileStream to load template
                templateStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                template = new ExcelPackage(templateStream);
                package = new ExcelPackage();

                // Reuse existing composable functions
                ExcelWorksheet worksheet = PrepareWorksheet(template, package);
                PopulateWorksheetAndCalculateForTempFile(worksheet, package, data);

                // Save Excel temporarily, then convert to PDF
                package.SaveAs(new FileInfo(tempExcelPath));
                Thread.Sleep(200);

                // Convert to PDF (this will clean up the temp Excel file)
                ConvertExcelToPDF(tempExcelPath, tempPdfPath);

                // Verify PDF was created
                if (!File.Exists(tempPdfPath) || new FileInfo(tempPdfPath).Length == 0)
                {
                    throw new InvalidOperationException("PDF generation failed");
                }

                return tempPdfPath;
            }
            finally
            {
                // Dispose EPPlus resources in proper order
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

                // Clean up temporary Excel file (PDF file should remain for caller)
                try
                {
                    if (File.Exists(tempExcelPath))
                    {
                        File.Delete(tempExcelPath);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Failed to delete temporary Excel file {tempExcelPath}: {ex.Message}");
                }

                // Controlled garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
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
        private ExcelWorksheet PrepareWorksheetConvertOnly(ExcelPackage template, ExcelPackage package)
        {
            ExcelWorksheet worksheet;

            if (template.Workbook.Worksheets.Count > 0)
            {
                // Copy the first worksheet from template
                var templateWorksheet = template.Workbook.Worksheets[1];
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

            // Auto-fit columns to content width - EPPlus method
            AutoFitColumns(worksheet);

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
                // Commented the function that deletes file in temp folder
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

            // Auto-fit columns to content width - EPPlus method
            AutoFitColumns(worksheet);

            // Force calculation for temp file operations
            package.Workbook.Calculate();

            // Allow calculation to complete - longer wait for temp file operations
            Thread.Sleep(200);
        }

        public string ExcelToPdf()
        {
            var templatePath = Path.Combine(_webHostEnvironment.ContentRootPath, "Assets", "psikogram.xlsx");

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"Excel template not found at: {templatePath}");
            }

            ExcelPackage template = null;
            ExcelPackage package = null;
            FileStream templateStream = null;

            var tempPdfPath = Path.Combine(Path.GetTempPath(), $"pdf_{Guid.NewGuid():N}.pdf");
            var tempExcelPath = Path.Combine(Path.GetTempPath(), $"temp_excel_{Guid.NewGuid():N}.xlsx");

            try
            {
                // Set EPPlus license context
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Use FileStream to load template
                templateStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                template = new ExcelPackage(templateStream);
                package = new ExcelPackage();

                // Reuse existing composable functions
                ExcelWorksheet worksheet = PrepareWorksheetConvertOnly(template, package);
                //PopulateWorksheetAndCalculateForTempFile(worksheet, package, data);

                // Save Excel temporarily, then convert to PDF
                package.SaveAs(new FileInfo(tempExcelPath));
                Thread.Sleep(200);

                // Convert to PDF (this will clean up the temp Excel file)
                ConvertExcelToPDF(tempExcelPath, tempPdfPath);

                // Verify PDF was created
                if (!File.Exists(tempPdfPath) || new FileInfo(tempPdfPath).Length == 0)
                {
                    throw new InvalidOperationException("PDF generation failed");
                }

                return tempPdfPath;
            }
            finally
            {
                // Dispose EPPlus resources in proper order
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

                // Clean up temporary Excel file (PDF file should remain for caller)
                try
                {
                    if (File.Exists(tempExcelPath))
                    {
                        File.Delete(tempExcelPath);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Failed to delete temporary Excel file {tempExcelPath}: {ex.Message}");
                }

                // Controlled garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }
    }
}