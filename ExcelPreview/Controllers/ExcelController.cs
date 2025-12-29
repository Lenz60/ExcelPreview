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

        // Keep your existing method
        [HttpGet]
        public IActionResult GetAllExcelData()
        {
            var data = _excelRepository.GetAllExcelData();
            return Ok(data);
        }
    }
}