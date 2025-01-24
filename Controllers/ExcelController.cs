using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelDataFixer.Controllers;

public class ExcelController : Controller
{
    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public IActionResult ProcessExcel(IFormFile excelFile)
    {
        if (excelFile == null || excelFile.Length == 0)
        {
            ViewBag.Message = "Please upload a valid Excel file.";
            return View("Index");
        }

        // Set the license context globally
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var stream = new MemoryStream())
        {
            excelFile.CopyTo(stream);

            // Ensure the stream's position is reset to the start
            stream.Position = 0;

            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    ViewBag.Message = "No worksheet found in the file.";
                    return View("Index");
                }

                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;

                // Add "Dear" prefix to column A (assuming column A contains names)
                for (int row = 2; row <= rowCount; row++) // Assuming headers are in row 1
                {
                    var cellValue = worksheet.Cells[row, 1].Text;
                    if (!string.IsNullOrWhiteSpace(cellValue) && !cellValue.StartsWith("Dear"))
                    {
                        worksheet.Cells[row, 1].Value = $"Dear {cellValue}";
                    }
                }

                // Check for missing data in required columns (A and B as examples)
                var missingDataRows = new List<int>();
                for (int row = 2; row <= rowCount; row++)
                {
                    if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Text) ||
                        string.IsNullOrWhiteSpace(worksheet.Cells[row, 2].Text))
                    {
                        missingDataRows.Add(row);
                    }
                }

                if (missingDataRows.Any())
                {
                    ViewBag.MissingData = $"Rows with missing data: {string.Join(", ", missingDataRows)}";
                }

                // Export the processed file
                var outputStream = new MemoryStream();
                package.SaveAs(outputStream);
                outputStream.Position = 0;

                return File(outputStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ProcessedFile.xlsx");
            }
        }
    }
}
