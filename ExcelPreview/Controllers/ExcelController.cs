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

                return Ok(new
                {
                    TempFilePath = tempFilePath,
                    Message = "Excel file generated successfully",
                    FileName = Path.GetFileName(tempFilePath)
                });
            }
            catch (Exception ex)
            {
                return BadRequest($"Error generating Excel file: {ex.Message}");
            }
        }

        [HttpGet("download-temp/{fileName}")]
        public IActionResult DownloadFromTemp(string fileName)
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

                // Clean up the temp file after reading
                try
                {
                    System.IO.File.Delete(filePath);
                }
                catch
                {
                    // Ignore cleanup errors
                }

                return File(fileContent,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error downloading file: {ex.Message}");
            }
        }

        // Keep your existing method
        [HttpGet]
        public IActionResult GetAllExcelData()
        {
            var data = _excelRepository.GetAllExcelData();
            return Ok(data);
        }
    }
}