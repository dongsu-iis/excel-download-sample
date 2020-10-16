using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using ClosedXML.Excel;
using excel_download_sample.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

namespace excel_download_sample.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult ExportExcel(string input)
        {
            // ファイルをwwwrootで生成する
            var fileName = Guid.NewGuid().ToString() + ".xlsx";
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), @"wwwroot\temp_excel", fileName);
            using XLWorkbook wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add(input);
            wb.SaveAs(filePath);

            // ファイル名だけ返す
            return Json(new ExcelContent
            {
                FileName = fileName
            });

        }


        public IActionResult Download(string fileName)
        {

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), @"wwwroot\temp_excel", fileName);
            byte[] data = System.IO.File.ReadAllBytes(filePath);

            // TODO:一時的に保存したファイルを消す処理を追加して

            return File(data, "application/vnd.ms-excel", fileName);
        }





        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

    }
}
