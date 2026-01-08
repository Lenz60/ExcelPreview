using Backend.Models;
using ExcelPreview.Repository.Interface;
using Microsoft.AspNetCore.Mvc;

namespace ExcelPreview.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly IExcelRepository _excelRepository;

        public ExcelController(IExcelRepository excelRepository)
        {
            _excelRepository = excelRepository;
        }

        [HttpGet("download")]
        public async Task<IActionResult> DownloadExcel()
        {
            try
            {
                var excelFile = await _excelRepository.GenerateExcelFileAsync();

                return File(
                    excelFile.FileContent,
                    excelFile.ContentType,
                    excelFile.FileName
                );
            }
            catch (Exception ex)
            {
                return BadRequest($"Error generating Excel file: {ex.Message}");
            }
        }

        [HttpGet("download-pdf")]
        public async Task<IActionResult> DownloadExcelAsPDF()
        {
            try
            {
                var data = await _excelRepository.GetAllExcelDataAsync();
                var pdfContent = _excelRepository.GenerateExcelAsPDF(data);

                var fileName = $"ExcelData_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";

                return File(
                    pdfContent,
                    "application/pdf",
                    fileName
                );
            }
            catch (Exception ex)
            {
                return BadRequest($"Error generating PDF file: {ex.Message}");
            }
        }

        [HttpGet("data")]
        public async Task<ActionResult<List<ExcelData>>> GetExcelData()
        {
            var data = await _excelRepository.GetAllExcelDataAsync();
            return Ok(data);
        }

        [HttpGet("temp-path")]
        public async Task<IActionResult> GetExcelTempPath()
        {
            try
            {
                var tempFilePath = await _excelRepository.GenerateExcelTempFileAsync();

                var fileName = Path.GetFileName(tempFilePath);

                // Set cookie on server when file is created
                var cookieOptions = new CookieOptions
                {
                    HttpOnly = false, // JavaScript can read this
                    Secure = true, // ✅ Only secure on HTTPS
                    SameSite = SameSiteMode.None,
                    Path = "/",
                    Expires = DateTimeOffset.Now.AddHours(2), // 2 hour expiry
                    Domain = null // ✅ No domain for localhost
                };

                Response.Cookies.Append("tempExcelFileName", fileName, cookieOptions);

                return Ok(new
                {
                    TempFilePath = tempFilePath,
                    Message = "Excel file generated successfully",
                    FileName = fileName
                });
            }
            catch (Exception ex)
            {
                return BadRequest($"Error generating Excel file: {ex.Message}");
            }
        }

        [HttpGet("preview-temp/{fileName}")]
        public IActionResult PreviewFromTemp(string fileName)
        {
            try
            {
                var tempDir = Path.GetTempPath();
                var filePath = Path.Combine(tempDir, fileName);

                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound("Temporary file not found or has been cleaned up.");
                }

                var fileContent = System.IO.File.ReadAllBytes(filePath);

                // Determine content type based on file extension
                var contentType = Path.GetExtension(fileName).ToLower() switch
                {
                    ".pdf" => "application/pdf",
                    ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ".xls" => "application/vnd.ms-excel",
                    _ => "application/octet-stream"
                };

                // return it for preview
                return File(fileContent, contentType);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error previewing file: {ex.Message}");
            }
        }

        //[HttpDelete("delete-excel-temp")]
        //public async Task<IActionResult> DeleteExcelTemp()
        //{
        //    try
        //    {
        //        // Get filename from cookie to delete the physical file
        //        string tempFileName = Request.Cookies["tempExcelFileName"];

        //        if (!string.IsNullOrEmpty(tempFileName))
        //        {
        //            var tempDir = Path.GetTempPath();
        //            var filePath = Path.Combine(tempDir, tempFileName);

        //            if (!System.IO.File.Exists(filePath))
        //            {
        //                return NotFound("Temporary file not found or has been cleaned up.");
        //            }

        //            var fileContent = System.IO.File.ReadAllBytes(filePath);

        //            // Clean up the temp file after reading
        //            try
        //            {
        //                System.IO.File.Delete(filePath);
        //            }
        //            catch
        //            {
        //                // Ignore cleanup errors
        //            }


        //        }

        //        // Clear the cookie
        //        var cookieOptions = new CookieOptions
        //        {
        //            HttpOnly = false,
        //            Secure = true,
        //            SameSite = SameSiteMode.None,
        //            Path = "/",
        //            Expires = DateTimeOffset.Now.AddDays(-1) // Expire immediately
        //        };

        //        Response.Cookies.Append("tempExcelFileName", "", cookieOptions);

        //        return Ok(new
        //        {
        //            Message = "Temp Excel data cleared successfully",
        //            FileName = tempFileName,
        //            Success = true
        //        });
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(new
        //        {
        //            Message = $"Error clearing temp data: {ex.Message}",
        //            Success = false
        //        });
        //    }
        //}

        [HttpGet("pdf-temp-path")]
        public async Task<IActionResult> GetPDFTempPath()
        {
            try
            {
                var tempFilePath = await _excelRepository.GenerateExcelAsPDFTempFileAsync();

                var fileName = Path.GetFileName(tempFilePath);

                // Set cookie on server when file is created
                var cookieOptions = new CookieOptions
                {
                    HttpOnly = false, // JavaScript can read this
                    Secure = true, // ✅ Only secure on HTTPS
                    SameSite = SameSiteMode.None,
                    Path = "/",
                    Expires = DateTimeOffset.Now.AddHours(2), // 2 hour expiry
                    Domain = null // ✅ No domain for localhost
                };

                Response.Cookies.Append("tempPdfFileName", fileName, cookieOptions);

                return Ok(new
                {
                    TempFilePath = tempFilePath,
                    Message = "PDF file generated successfully",
                    FileName = Path.GetFileName(tempFilePath)
                });
            }
            catch (Exception ex)
            {
                return BadRequest($"Error generating PDF file: {ex.Message}");
            }
        }

        //[HttpGet("download-temp/{fileName}")]
        //public IActionResult DownloadFromTemp(string fileName)
        //{
        //    try
        //    {
        //        var tempDir = Path.GetTempPath();
        //        var filePath = Path.Combine(tempDir, fileName);

        //        if (!System.IO.File.Exists(filePath))
        //        {
        //            return NotFound("Temporary file not found or has been cleaned up.");
        //        }

        //        var fileContent = System.IO.File.ReadAllBytes(filePath);

        //        // Determine content type based on file extension
        //        var contentType = Path.GetExtension(fileName).ToLower() switch
        //        {
        //            ".pdf" => "application/pdf",
        //            ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        //            ".xls" => "application/vnd.ms-excel",
        //            _ => "application/octet-stream"
        //        };



        //        //// Clean up the temp file after reading
        //        //try
        //        //{
        //        //    System.IO.File.Delete(filePath);
        //        //}
        //        //catch
        //        //{
        //        //    // Ignore cleanup errors
        //        //}

        //        return File(fileContent, contentType, fileName);
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest($"Error downloading file: {ex.Message}");
        //    }
        //}
        [HttpGet("download-temp")]
        public IActionResult DownloadFromTempNoParameter()
        {
            try
            {
                var fileName = Request.Cookies["tempExcelFileName"];
                var tempDir = Path.GetTempPath();
                var filePath = Path.Combine(tempDir, fileName);

                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound("Temporary file not found or has been cleaned up.");
                }

                var fileContent = System.IO.File.ReadAllBytes(filePath);

                // Determine content type based on file extension
                var contentType = Path.GetExtension(fileName).ToLower() switch
                {
                    ".pdf" => "application/pdf",
                    ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ".xls" => "application/vnd.ms-excel",
                    _ => "application/octet-stream"
                };

                // Clear the cookie
                var cookieOptions = new CookieOptions
                {
                    HttpOnly = false,
                    Secure = true,
                    SameSite = SameSiteMode.None,
                    Path = "/",
                    Expires = DateTimeOffset.Now.AddDays(-1) // Expire immediately
                };

                Response.Cookies.Append("tempExcelFileName", "", cookieOptions);

                // Clean up the temp file after reading
                try
                {
                    System.IO.File.Delete(filePath);
                }
                catch
                {
                    // Ignore cleanup errors
                }

                return File(fileContent, contentType, fileName);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error downloading file: {ex.Message}");
            }
        }

        [HttpGet("download-pdf-temp")]
        public IActionResult DownloadPDFFromTemp()
        {
            try
            {
                var fileName = Request.Cookies["tempPdfFileName"];
                var tempDir = Path.GetTempPath();
                var filePath = Path.Combine(tempDir, fileName);

                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound("Temporary PDF file not found or has been cleaned up.");
                }

                var fileContent = System.IO.File.ReadAllBytes(filePath);

                // Clear the cookie
                var cookieOptions = new CookieOptions
                {
                    HttpOnly = false,
                    Secure = true,
                    SameSite = SameSiteMode.None,
                    Path = "/",
                    Expires = DateTimeOffset.Now.AddDays(-1) // Expire immediately
                };

                Response.Cookies.Append("tempPdfFileName", "", cookieOptions);

                // Clean up the temp file after reading
                try
                {
                    System.IO.File.Delete(filePath);
                }
                catch
                {
                    // Ignore cleanup errors
                }

                return File(fileContent, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error downloading PDF file: {ex.Message}");
            }
        }

        // Keep your existing method
        [HttpGet]
        public IActionResult GetAllExcelData()
        {
            var data = _excelRepository.GetAllExcelData();
            return Ok(data);
        }

        // Additional endpoint to generate PDF with custom data
        [HttpPost("generate-pdf")]
        public IActionResult GeneratePDFWithData([FromBody] List<ExcelData> data)
        {
            try
            {
                if (data == null || !data.Any())
                {
                    return BadRequest("No data provided for PDF generation.");
                }

                var pdfContent = _excelRepository.GenerateExcelAsPDF(data);
                var fileName = $"CustomExcelData_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";

                return File(pdfContent, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error generating PDF with custom data: {ex.Message}");
            }
        }

        [HttpGet("preview-pdf-temp/{fileName}")]
        public IActionResult PreviewPDFFromTemp(string fileName)
        {
            try
            {
                var tempDir = Path.GetTempPath();
                var filePath = Path.Combine(tempDir, fileName);

                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound("Temporary PDF file not found or has been cleaned up.");
                }

                var fileContent = System.IO.File.ReadAllBytes(filePath);

                //// Clean up the temp file after reading
                //try
                //{
                //    System.IO.File.Delete(filePath);
                //}
                //catch
                //{
                //    // Ignore cleanup errors
                //}

                // ✅ Add headers to allow iframe embedding
                Response.Headers.Add("X-Frame-Options", "SAMEORIGIN"); // Allow same origin
                Response.Headers.Add("Content-Security-Policy", "frame-ancestors 'self' https://localhost:*"); // Allow localhost iframes
                Response.Headers.Add("Cache-Control", "no-cache, no-store, must-revalidate");
                Response.Headers.Add("Pragma", "no-cache");
                Response.Headers.Add("Expires", "0");

                // Return PDF for inline viewing (no filename = no download prompt)
                return File(fileContent, "application/pdf");
            }
            catch (Exception ex)
            {
                return BadRequest($"Error previewing PDF file: {ex.Message}");
            }
        }

        // Additional endpoint to generate Excel with custom data (for consistency)
        [HttpPost("generate-excel")]
        public IActionResult GenerateExcelWithData([FromBody] List<ExcelData> data)
        {
            try
            {
                if (data == null || !data.Any())
                {
                    return BadRequest("No data provided for Excel generation.");
                }

                var excelFile = _excelRepository.GenerateExcelFileSync(data);

                return File(
                    excelFile.FileContent,
                    excelFile.ContentType,
                    excelFile.FileName
                );
            }
            catch (Exception ex)
            {
                return BadRequest($"Error generating Excel with custom data: {ex.Message}");
            }
        }

        [HttpGet("convert-excel-pdf")]
        public IActionResult ConvertExcelToPdf()
        {
            try
            {
                var pdfFilePath = _excelRepository.ExcelToPdf();
                var fileName = Path.GetFileName(pdfFilePath);
                var fileContent = System.IO.File.ReadAllBytes(pdfFilePath);
                return File(
                    fileContent,
                    "application/pdf",
                    fileName
                );
            }
            catch (Exception ex)
            {
                return BadRequest($"Error converting Excel to PDF: {ex.Message}");
            }
        }
    }
}