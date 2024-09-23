using ClosedXML.Excel;
using ExcelAutomationTask.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace ExcelAutomationTask.Controllers
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

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult ExcelReader()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ExcelReader(IFormFile file)
        {
            if (file != null && file.Length > 0)
            {
                var uploadDirectory = $"{Directory.GetCurrentDirectory()}\\wwwroot\\Uploads";

                if (!Directory.Exists(uploadDirectory))
                {
                    Directory.CreateDirectory(uploadDirectory);
                }

                var filePath = Path.Combine(uploadDirectory, file.FileName);

                using (var stream = new FileStream(filePath, FileMode.Create)) //var stream = new MemoryStream() if I Want it on the fly
                {
                    await file.CopyToAsync(stream);

                    using (var workbook = new XLWorkbook(stream))
                    {
                        var worksheet = workbook.Worksheet(1);
                        // Add new column and compute values
                        worksheet.Column(worksheet.LastColumnUsed().ColumnNumber() + 1).Cell(1).Value = "Total Value before Taxing";
                        for (int i = 2; i <= worksheet.LastRowUsed().RowNumber(); i++)
                        {
                            worksheet.Cell(i, worksheet.LastColumnUsed().ColumnNumber()).Value = ComputeTotalBeforeTaxing(worksheet.Row(i));
                        }

                        // Add new row for total
                        var totalRow = worksheet.LastRowUsed().RowNumber() + 1;
                        worksheet.Cell(totalRow, 1).Value = "Total";
                        worksheet.Cell(totalRow, worksheet.LastColumnUsed().ColumnNumber()).Value = worksheet.Column(worksheet.LastColumnUsed().ColumnNumber()).CellsUsed().Sum(cell => cell.GetValue<double>());

                        // Save modified file
                        using (var memoryStream = new MemoryStream())
                        {
                            workbook.SaveAs(memoryStream);
                            return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ModifiedSheet.xlsx");
                        }
                    }
                }
            }
            return View();
        }

        private double ComputeTotalBeforeTaxing(IXLRow row)
        {
            double TotalBeforeTaxing = 0,
                   TotalAfterTaxing = 0,
                   TaxingValue = 0;
            if (row == null)
                return 0;

            TotalAfterTaxing = row.Cell(7).GetDouble();

            TaxingValue = row.Cell(5).GetDouble();

            TotalBeforeTaxing = TotalAfterTaxing - TaxingValue;

            return TotalBeforeTaxing;
        }

    }
}
