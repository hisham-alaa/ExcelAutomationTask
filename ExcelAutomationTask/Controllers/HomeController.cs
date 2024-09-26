using ClosedXML.Excel;
using ExcelAutomationTask.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.IO;

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
                using (var memoryStream = new MemoryStream())
                {
                    await file.CopyToAsync(memoryStream);
                    memoryStream.Position = 0;

                    using (var workbook = new XLWorkbook(memoryStream))
                    {
                        var worksheet = workbook.Worksheet(1);

                        double SumOfTotalAfterTaxing = 0;

                        // Add new column and compute values
                        int lastColumn = worksheet.LastColumnUsed().ColumnNumber();

                        IXLStyle style = worksheet.Cell(1, lastColumn).Style;
                        worksheet.Cell(1, lastColumn + 1).Style = style;

                        worksheet.Cell(1, ++lastColumn).Value = "Total Value before Taxing";
                        IXLRow row;

                        for (int i = 2; i <= worksheet.LastRowUsed().RowNumber(); i++)
                        {
                            row = worksheet.Row(i);
                            worksheet.Cell(i, lastColumn).Style = style;
                            worksheet.Cell(i, lastColumn).Value = ComputeTotalBeforeTaxing(row);
                            SumOfTotalAfterTaxing += row.Cell(7).GetDouble();
                        }

                        // Add new row for total
                        var totalRow = worksheet.LastRowUsed().RowNumber() + 1;

                        worksheet.Cell(totalRow, 1).Value = "Total After Taxing";

                        worksheet.Cell(totalRow, 7).Value = SumOfTotalAfterTaxing;

                        using (var outputStream = new MemoryStream())
                        {
                            workbook.SaveAs(outputStream);
                            outputStream.Position = 0;
                            return File(outputStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"Modified_{file.FileName}");
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

            TaxingValue = row.Cell(8).GetDouble();

            TotalBeforeTaxing = TotalAfterTaxing - TaxingValue;

            return TotalBeforeTaxing;
        }

    }
}
