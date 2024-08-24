using Aspose_Word_Net8.Models;
using Microsoft.AspNetCore.Mvc;

namespace Aspose_Word_Net8.Controllers
{
    public class HomeController : Controller
    {
        private readonly IHttpContextAccessor _httpContextAccessor;
        private readonly IConfiguration _configuration;
        public HomeController(IHttpContextAccessor httpContextAccessor, IConfiguration configuration)
        {
            _httpContextAccessor = httpContextAccessor;
            _configuration = configuration;
        }

        public IActionResult Index()
        {
            string[] Field = new string[] { "Shomareh", "Tarikh", "Peyvast", "HtmlMatn", "HtmlEmza", "HtmlRoonevesht", "HtmlParaf" };
            object[] data;

            data = new object[] {
                "Shomareh",
                "ShamsiDateNow",
                "",
                "PreMatnName",
                "getEmzaFormatHtmlFA",
                "Roonevesht رونوشت",
                "ShowparaphsPrint"
            };

            string path = iAspose.BuildPrintLetter(Field, data);
            TempData["file"] = path;

            return View();
        }
    }
}
