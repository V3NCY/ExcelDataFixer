using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Data;

namespace ExcelDataFixer.Controllers;

public class ExcelController : Controller
{
    private static DataTable _backupDataTable;
    private static DataTable _currentDataTable;

    public IActionResult Index()
    {
        ViewBag.DataTable = null; // Clear old data on page load
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

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var stream = new MemoryStream())
        {
            excelFile.CopyTo(stream);
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

                if (rowCount > 20000)
                {
                    ViewBag.Message = "The uploaded file exceeds the maximum allowed rows (20,000). Please reduce the data size.";
                    return View("Index");
                }

                DataTable table = new DataTable();

                for (int col = 1; col <= colCount; col++)
                {
                    string columnName = worksheet.Cells[1, col].Text.Trim();

                    if (string.IsNullOrWhiteSpace(columnName))
                    {
                        columnName = $"Column {col}";
                    }
                    else
                    {
                        int duplicateCount = 1;
                        string originalName = columnName;

                        while (table.Columns.Contains(columnName))
                        {
                            columnName = $"{originalName}_{duplicateCount++}";
                        }
                    }

                    table.Columns.Add(columnName);
                }


                for (int row = 2; row <= rowCount; row++)
                {
                    var dataRow = table.NewRow();
                    bool isReplacementRow = worksheet.Cells[row, 1].Text.Trim().Equals("Васил Динолов", StringComparison.OrdinalIgnoreCase) &&
                                            worksheet.Cells[row, 3].Text.Trim().Equals("vasil.dinolov@orakgroup.com", StringComparison.OrdinalIgnoreCase);

                    for (int col = 1; col <= colCount; col++)
                    {
                        if (isReplacementRow)
                        {
                            if (col == 1) dataRow[col - 1] = "Цветан Карабов";
                            else if (col == 2) dataRow[col - 1] = "0888227303";
                            else if (col == 3) dataRow[col - 1] = "tsvetan.karabov@orakgroup.com";
                            else if (col == 6) dataRow[col - 1] = "Директор Дигитално образование и Иновации";
                            else dataRow[col - 1] = worksheet.Cells[row, col].Text; // Keep original for other columns
                        }

                        else
                        {
                            var cellValue = worksheet.Cells[row, col].Text;

                            // Replace the name Васил Динолов with Цветан Карабов in Column 3
                            if (col == 3 && cellValue.Trim().Equals("Васил Динолов", StringComparison.OrdinalIgnoreCase))
                            {
                                dataRow[col - 1] = "Цветан Карабов";
                            }
                            // Replace the number 0886988005 with 0888227303 in Column 4
                            else if (col == 4 && cellValue.Trim().Equals("0886988005", StringComparison.OrdinalIgnoreCase))
                            {
                                dataRow[col - 1] = "0888227303";
                            }
                            // Replace the email vasil.dinolov@orakgroup.com with tsvetan.karabov@orakgroup.com in Column 5
                            else if (col == 5 && cellValue.Trim().Equals("vasil.dinolov@orakgroup.com", StringComparison.OrdinalIgnoreCase))
                            {
                                dataRow[col - 1] = "tsvetan.karabov@orakgroup.com";
                            }
                            // Automatically add titles based on names from Column 25 in Column 2
                            else if (col == 2)
                            {
                                var nameFromColumn25 = worksheet.Cells[row, 25]?.Text?.Trim(); // Взима стойността от колоната 25
                                if (!string.IsNullOrWhiteSpace(nameFromColumn25))
                                {
                                    // Извличаме фамилията (последната дума от името)
                                    var lastName = nameFromColumn25.Split(' ').Last();

                                    // Проверка за окончанията на фамилията
                                    if (lastName.EndsWith("ова") || lastName.EndsWith("ева")) // Женски фамилии
                                    {
                                        dataRow[col - 1] = $"Уважаема г-жо {lastName}";
                                    }
                                    else if (lastName.EndsWith("ина") || lastName.EndsWith("рян")) // Женски фамилии
                                    {
                                        dataRow[col - 1] = $"Уважаема г-жо {lastName}";
                                    }
                                    else if (lastName.EndsWith("ска")) // Женски фамилии
                                    {
                                        dataRow[col - 1] = $"Уважаема г-жо {lastName}";
                                    }
                                    else if (lastName.EndsWith("ов") || lastName.EndsWith("ев")) // Мъжки фамилии
                                    {
                                        dataRow[col - 1] = $"Уважаеми г-н {lastName}";
                                    }
                                    else
                                    {
                                        dataRow[col - 1] = $"Уважаеми {lastName}"; // Дефолтно обръщение, ако окончанието е неутрално
                                    }
                                }
                                else
                                {
                                    dataRow[col - 1] = string.Empty; // Ако колоната е празна
                                }
                            }

                            else if (col == 6)
                            {
                                // Add titles to Column 6 based on names in Column 3
                                var nameFromColumn3 = worksheet.Cells[row, 3].Text.Trim();
                                if (!string.IsNullOrWhiteSpace(nameFromColumn3))
                                {
                                    switch (nameFromColumn3)
                                    {
                                        case "Виктория Добрева":
                                            dataRow[col - 1] = "Старши търговски сътрудник";
                                            break;
                                        case "Борислава Димова":
                                            dataRow[col - 1] = "Старши търговски сътрудник";
                                            break;
                                        case "Христина Илчева":
                                            dataRow[col - 1] = "Старши търговски сътрудник";
                                            break;
                                        case "Йордан Тотев":
                                            dataRow[col - 1] = "Търговски представител - област Бургас";
                                            break;
                                        case "Милена Цанова":
                                            dataRow[col - 1] = "Търговски представител - област Варна";
                                            break;
                                        case "Мариета Йорданова":
                                            dataRow[col - 1] = "Търговски представител - София 2";
                                            break;
                                        case "Цветан Карабов":
                                            dataRow[col - 1] = "Директор Дигитално образование и Иновации";
                                            break;
                                        default:
                                            dataRow[col - 1] = string.Empty;
                                            break;
                                    }
                                }
                            }
                            else if (col == 4)
                            {
                                // Add titles to Column 4 based on names in Column 3
                                var nameFromColumn3 = worksheet.Cells[row, 3].Text.Trim();
                                if (!string.IsNullOrWhiteSpace(nameFromColumn3))
                                {


                                    switch (nameFromColumn3)
                                    {
                                        case "Виктория Добрева":
                                            dataRow[col - 1] = "0882927244";
                                            break;
                                        case "Борислава Димова":
                                            dataRow[col - 1] = "0882538928";
                                            break;
                                        case "Христина Илчева":
                                            dataRow[col - 1] = "0883244264";
                                            break;
                                        case "Йордан Тотев":
                                            dataRow[col - 1] = "0886866222";
                                            break;
                                        case "Милена Цанова":
                                            dataRow[col - 1] = "0884754064";
                                            break;
                                        case "Мариета Йорданова":
                                            dataRow[col - 1] = "0887585137";
                                            break;
                                        case "Цветан Карабов":
                                            dataRow[col - 1] = "0888227303";
                                            break;
                                        default:
                                            dataRow[col - 1] = string.Empty;
                                            break;
                                    }
                                }
                            }
                            else if (col == 5)
                            {
                                // Add titles to Column 5 based on names in Column 3
                                var nameFromColumn3 = worksheet.Cells[row, 3].Text.Trim();
                                if (!string.IsNullOrWhiteSpace(nameFromColumn3))
                                {
                                    switch (nameFromColumn3)
                                    {
                                        case "Виктория Добрева":
                                            dataRow[col - 1] = "victoria.dobreva@orakgroup.com";
                                            break;
                                        case "Борислава Димова":
                                            dataRow[col - 1] = "borislava.dimova@orakgroup.com";
                                            break;
                                        case "Христина Илчева":
                                            dataRow[col - 1] = "hristina.ilcheva@orakgroup.com";
                                            break;
                                        case "Йордан Тотев":
                                            dataRow[col - 1] = "yordan.totev@orakgroup.com";
                                            break;
                                        case "Милена Цанова":
                                            dataRow[col - 1] = "milena.tsanova@orakgroup.com";
                                            break;
                                        case "Мариета Йорданова":
                                            dataRow[col - 1] = "marieta.yordanova@orakgroup.com";
                                            break;
                                        case "Цветан Карабов":
                                            dataRow[col - 1] = "cvetan.karabov@orakgroup.com";
                                            break;
                                        default:
                                            dataRow[col - 1] = string.Empty;
                                            break;
                                    }
                                }
                            }
                            else if (col == 32)
                            {
                                // Add email based on Column 22
                                var codeFromColumn22 = worksheet.Cells[row, 22].Text.Trim();
                                if (!string.IsNullOrWhiteSpace(codeFromColumn22))
                                {
                                    dataRow[col - 1] = $"info-{codeFromColumn22}@edu.mon.bg";
                                }
                                else
                                {
                                    dataRow[col - 1] = string.Empty;
                                }
                            }
                            else
                            {
                                dataRow[col - 1] = cellValue;
                            }
                        }
                    }
                    table.Rows.Add(dataRow);
                }

                _currentDataTable = table.Copy();
                _backupDataTable = table.Copy();
                ViewBag.DataTable = table;

                return View("Index");
            }
        }
    }

    [HttpPost]
    public IActionResult ClearData()
    {
        ViewBag.DataTable = null;
        ViewBag.Message = "Data cleared successfully.";
        return View("Index");
    }

    [HttpPost]
    public IActionResult RevertData()
    {
        if (_backupDataTable != null)
        {
            ViewBag.DataTable = _backupDataTable.Copy();
            ViewBag.Message = "Data reverted successfully.";
        }
        else
        {
            ViewBag.Message = "No data to revert.";
        }
        return View("Index");
    }

    [HttpPost]
    public IActionResult ExportExcel()
    {
        if (_currentDataTable == null || _currentDataTable.Rows.Count == 0)
        {
            ViewBag.Message = "No data to export.";
            return View("Index");
        }

        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");

            // Add column headers
            for (int col = 0; col < _currentDataTable.Columns.Count; col++)
            {
                worksheet.Cells[1, col + 1].Value = _currentDataTable.Columns[col].ColumnName;
            }

            // Add rows
            for (int row = 0; row < _currentDataTable.Rows.Count; row++)
            {
                for (int col = 0; col < _currentDataTable.Columns.Count; col++)
                {
                    worksheet.Cells[row + 2, col + 1].Value = _currentDataTable.Rows[row][col];
                }
            }

            var stream = new MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            var fileName = $"ProcessedFile-edited.xlsx";
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

    }
    
}