using Microsoft.AspNetCore.Mvc;

namespace ExcelPreview.Controllers
{
    public class ExcelController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
