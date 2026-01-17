using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace MonitoringSystem.Pages
{
    public class IndexModel : PageModel
    {
        //public string connectionString = "Server=XDZALL\\SQLEXPRESS;Database=PROMOSYS;Trusted_Connection=True;Encrypt=False";
        private readonly string connectionString = "Data Source=DESKTOP-NBPATD6\\MSSQLSERVERR;trusted_connection=true;trustservercertificate=True;Database=LatestPROMOSYS;Integrated Security=True;Encrypt=False";
        private readonly ILogger<IndexModel> _logger;
        private bool IsValidInput(int month, int year) => month >= 1 && month <= 12 && year > 0;

        public IndexModel(ILogger<IndexModel> logger)
        {
            _logger = logger;
        }

        public void OnGet() { }

        public IActionResult OnPostExportToExcel(int month, int year)
        {
            if (!IsValidInput(month, year))
                return Content("Invalid month or year input.", "text/plain");

            try
            {
                var (startDate, endDate) = GetDateRange(month, year);
                var productionData = LoadProductionData(startDate, endDate);
                var lossTimeData = LoadLossTimeData(startDate, endDate);
                var excelStream = GenerateExcelFile(startDate, endDate, productionData, lossTimeData);

                string fileName = GenerateFileName(month, year);
                return File(excelStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred during Excel export.");
                return Content("An error occurred while generating the Excel file.", "text/plain");
            }
        }

        private (DateTime startDate, DateTime endDate) GetDateRange(int month, int year)
        {
            var startDate = new DateTime(year, month, 1);
            var endDate = startDate.AddMonths(1).AddDays(-1);
            return (startDate, endDate);
        }

        private Dictionary<string, Dictionary<DateTime, int>> LoadProductionData(DateTime startDate, DateTime endDate)
        {
            var productionData = new Dictionary<string, Dictionary<DateTime, int>>();
            ExecuteReader(@"
                SELECT MD.ProductName, COUNT(OE.Product_Id) AS ProductionCount, CAST(OE.Date AS DATE) AS ProductionDate
                FROM MasterData MD
                LEFT JOIN OEESN OE ON MD.Product_Id = OE.Product_Id
                WHERE OE.Date >= @startDate AND OE.Date < @endDatePlusOne
                GROUP BY MD.ProductName, CAST(OE.Date AS DATE)",
                new[]
                {
                    new SqlParameter("@startDate", startDate),
                    new SqlParameter("@endDatePlusOne", endDate.AddDays(1))
                },
                reader =>
                {
                    var productName = reader["ProductName"].ToString();
                    var productionDate = Convert.ToDateTime(reader["ProductionDate"]).Date;
                    var productionCount = Convert.ToInt32(reader["ProductionCount"]);

                    if (!productionData.ContainsKey(productName))
                        productionData[productName] = new Dictionary<DateTime, int>();

                    productionData[productName][productionDate] = productionCount;
                });

            return productionData;
        }
        
        private void ExecuteReader(string query, SqlParameter[] parameters, Action<SqlDataReader> readAction)
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();

            using var command = new SqlCommand(query, connection);
            command.Parameters.AddRange(parameters);

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                readAction(reader);
            }
        }

        private MemoryStream GenerateExcelFile(DateTime startDate, DateTime endDate, Dictionary<string, Dictionary<DateTime, int>> productionData, Dictionary<string, Dictionary<DateTime, int>> lossTimeData)
        {
            var stream = new MemoryStream();
            using var package = new ExcelPackage(stream);
            var worksheet = package.Workbook.Worksheets.Add("Monthly Report");

            WriteHeader(worksheet, startDate, endDate);
            WriteData(worksheet, startDate, endDate, productionData);

            int lossTimeStartRow = productionData.Count + 5;

            WriteLossTimeTable(worksheet, startDate, endDate, lossTimeData, lossTimeStartRow);


            worksheet.Cells.AutoFitColumns();
            package.Save();
            stream.Position = 0;
            return stream;
        }

        private void WriteLossTimeTable(ExcelWorksheet worksheet, DateTime startDate, DateTime endDate, Dictionary<string, Dictionary<DateTime, int>> lossTimeData, int lossTimeStartRow)
        {
            worksheet.Cells[lossTimeStartRow, 1].Value = "LOSS TIME";
            worksheet.Cells[lossTimeStartRow, 1, lossTimeStartRow, 32].Merge = true;
            worksheet.Cells[lossTimeStartRow, 1].Style.Font.Bold = true;

            WriteDateHeaders(worksheet, startDate, endDate, lossTimeStartRow + 1);
            worksheet.Cells[lossTimeStartRow + 1, 33].Value = "Total";
            worksheet.Cells[lossTimeStartRow + 1, 33].Style.Font.Bold = true;
            ApplyCellStyle(worksheet.Cells[lossTimeStartRow + 1, 33], System.Drawing.Color.LightGreen);

            var categories = lossTimeData.Keys.OrderBy(k => k).ToList();
            var rowIndex = lossTimeStartRow + 2;
            var grandTotal = 0;

            foreach (var category in categories)
            {
                WriteLossTimeCategory(worksheet, startDate, endDate, lossTimeData, category, ref rowIndex, ref grandTotal);
            }

            WriteLossTimeTotals(worksheet, startDate, endDate, categories, lossTimeData, rowIndex, grandTotal);
        }

        private void WriteLossTimeCategory(ExcelWorksheet worksheet, DateTime startDate, DateTime endDate, Dictionary<string, Dictionary<DateTime, int>> lossTimeData, string category, ref int rowIndex, ref int grandTotal)
        {
            worksheet.Cells[rowIndex, 1].Value = category;
            var rowTotal = 0;
            var colIndex = 2;

            for (var date = startDate; date <= endDate; date = date.AddDays(1))
            {
                var dailyValue = lossTimeData[category].GetValueOrDefault(date, 0);
                worksheet.Cells[rowIndex, colIndex].Value = dailyValue;
                rowTotal += dailyValue;
                colIndex++;
            }

            worksheet.Cells[rowIndex, colIndex].Value = rowTotal;
            ApplyCellStyle(worksheet.Cells[rowIndex, colIndex], System.Drawing.Color.LightGreen, isBold: true);

            grandTotal += rowTotal;
            rowIndex++;
        }

        private void WriteLossTimeTotals(ExcelWorksheet worksheet, DateTime startDate, DateTime endDate, List<string> categories, Dictionary<string, Dictionary<DateTime, int>> lossTimeData, int rowIndex, int grandTotal)
        {
            worksheet.Cells[rowIndex, 1].Value = "Total";
            worksheet.Cells[rowIndex, 1].Style.Font.Bold = true;
            var colIndex = 2;

            for (var date = startDate; date <= endDate; date = date.AddDays(1))
            {
                var dailyTotal = categories.Sum(category => lossTimeData[category].GetValueOrDefault(date, 0));
                worksheet.Cells[rowIndex, colIndex].Value = dailyTotal;
                ApplyCellStyle(worksheet.Cells[rowIndex, colIndex], System.Drawing.Color.LightGreen, isBold: true);
                colIndex++;
            }

            worksheet.Cells[rowIndex, colIndex].Value = grandTotal;
            ApplyCellStyle(worksheet.Cells[rowIndex, colIndex], System.Drawing.Color.LightGreen, isBold: true);
        }

        private void WriteHeader(ExcelWorksheet worksheet, DateTime startDate, DateTime endDate)
        {
            worksheet.Cells[1, 1].Value = "PRODUCTION RESULT";
            worksheet.Cells[1, 1, 1, 32].Merge = true;
            worksheet.Cells[1, 1].Style.Font.Bold = true;

            WriteDateHeaders(worksheet, startDate, endDate, 2);
            worksheet.Cells[2, 33].Value = "Total";
            worksheet.Cells[2, 33].Style.Font.Bold = true;
            ApplyCellStyle(worksheet.Cells[2, 33], System.Drawing.Color.LightGreen);
        }

        private void WriteDateHeaders(ExcelWorksheet worksheet, DateTime startDate, DateTime endDate, int dateRow)
        {
            var colIndex = 2;
            for (var date = startDate; date <= endDate; date = date.AddDays(1))
            {
                worksheet.Cells[dateRow, colIndex].Value = date.ToString("ddd, dd MMM yyyy", CultureInfo.InvariantCulture);
                if (date.DayOfWeek == DayOfWeek.Sunday)
                {
                    ApplyCellStyle(worksheet.Cells[dateRow, colIndex], System.Drawing.Color.Red);
                }
                colIndex++;
            }
        }

        private void WriteData(ExcelWorksheet worksheet, DateTime startDate, DateTime endDate, Dictionary<string, Dictionary<DateTime, int>> productionData)
        {
            var sortedKeys = productionData.Keys.OrderBy(k => k).ToList();
            var rowIndex = 3;
            var grandTotal = 0;

            foreach (var product in sortedKeys)
            {
                WriteProductData(worksheet, startDate, endDate, productionData, product, ref rowIndex, ref grandTotal);
            }

            WriteTotals(worksheet, startDate, endDate, sortedKeys, productionData, rowIndex, grandTotal);
        }

        private void WriteProductData(ExcelWorksheet worksheet, DateTime startDate, DateTime endDate, Dictionary<string, Dictionary<DateTime, int>> productionData, string product, ref int rowIndex, ref int grandTotal)
        {
            worksheet.Cells[rowIndex, 1].Value = product;
            var rowTotal = 0;
            var colIndex = 2;

            for (var date = startDate; date <= endDate; date = date.AddDays(1))
            {
                var dailyValue = productionData[product].GetValueOrDefault(date, 0);
                worksheet.Cells[rowIndex, colIndex].Value = dailyValue;
                rowTotal += dailyValue;
                colIndex++;
            }

            worksheet.Cells[rowIndex, colIndex].Value = rowTotal;
            ApplyCellStyle(worksheet.Cells[rowIndex, colIndex], System.Drawing.Color.LightGreen, isBold: true);

            grandTotal += rowTotal;
            rowIndex++;
        }

        private void WriteTotals(ExcelWorksheet worksheet, DateTime startDate, DateTime endDate, List<string> sortedKeys, Dictionary<string, Dictionary<DateTime, int>> productionData, int rowIndex, int grandTotal)
        {
            worksheet.Cells[rowIndex, 1].Value = "Total";
            worksheet.Cells[rowIndex, 1].Style.Font.Bold = true;
            var colIndex = 2;

            for (var date = startDate; date <= endDate; date = date.AddDays(1))
            {
                var dailyTotal = sortedKeys.Sum(key => productionData[key].GetValueOrDefault(date, 0));
                worksheet.Cells[rowIndex, colIndex].Value = dailyTotal;
                ApplyCellStyle(worksheet.Cells[rowIndex, colIndex], System.Drawing.Color.LightGreen, isBold: true);
                colIndex++;
            }

            worksheet.Cells[rowIndex, colIndex].Value = grandTotal;
            ApplyCellStyle(worksheet.Cells[rowIndex, colIndex], System.Drawing.Color.LightGreen, isBold: true);
        }

        private Dictionary<string, Dictionary<DateTime, int>> LoadLossTimeData(DateTime startDate, DateTime endDate)
        {
            var lossTimeData = new Dictionary<string, Dictionary<DateTime, int>>();

            // Initialize categories
            var categories = new[] { "MAN", "MACHINE", "SHORTAGE", "QUALITY" };
            foreach (var category in categories)
            {
                lossTimeData[category] = new Dictionary<DateTime, int>();
            }

            ExecuteReader(@"
                    SELECT CAST(Date AS DATE) AS LossDate, Reason, SUM(LossTime) AS TotalLossTime
                    FROM AssemblyLossTime
                    WHERE Date >= @startDate AND Date < @endDatePlusOne
                    GROUP BY CAST(Date AS DATE), Reason",
                new[]
                {
                    new SqlParameter("@startDate", startDate),
                    new SqlParameter("@endDatePlusOne", endDate.AddDays(1))
                },
                reader =>
                {
                    var lossDate = Convert.ToDateTime(reader["LossDate"]).Date;
                    var reason = reader["Reason"].ToString();
                    var lossTimeMinutes = Convert.ToInt32(reader["TotalLossTime"]);

                    // Only include defined categories
                    if (lossTimeData.ContainsKey(reason))
                    {
                        lossTimeData[reason][lossDate] = lossTimeMinutes;
                    }
                });

            return lossTimeData;
        }


        private void ApplyCellStyle(ExcelRange cell, System.Drawing.Color color, bool isBold = false)
        {
            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(color);
            cell.Style.Font.Bold = isBold;
        }

        private string GenerateFileName(int month, int year)
        {
            return $"Monthly Report {CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month)} {year}.xlsx";
        }
    }
}