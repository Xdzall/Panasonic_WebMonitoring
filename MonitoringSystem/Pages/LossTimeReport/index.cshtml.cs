using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;
using MonitoringSystem.Models;
using MonitoringSystem.Data;
using System;

namespace MonitoringSystem.Pages.LossTimeReport
{
    public class indexModel : PageModel
    {
        private readonly ApplicationDbContext _context;

        public indexModel(ApplicationDbContext context)
        {
            _context = context;
        }

        [BindProperty(SupportsGet = true)]
        public int SelectedYear { get; set; } = DateTime.Today.Year;

        // Filter Utama untuk Tampilan Chart
        [BindProperty(SupportsGet = true)]
        public string MachineLine { get; set; } = "All";

        // Property Khusus untuk Form Upload di dalam Pop-up
        [BindProperty]
        public string UploadMachineLine { get; set; }

        [BindProperty]
        public IFormFile UploadedExcel { get; set; }

        public string ChartDataJson { get; set; } = "{}";

        public void OnGet()
        {
            // --- STRUKTUR DATA 12 BULAN (FISCAL YEAR: APRIL - MARET) ---
            string[] months = { "April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March" };
            double[] actualData = new double[12];
            double[] planData = new double[12];

            // 1. AMBIL DATA ACTUAL (Raw SQL)
            // Jika MachineLine == "All", query akan menjumlahkan (SUM) data dari semua mesin
            var actualsFromDb = GetMonthlyActualData(SelectedYear, MachineLine);

            // Mapping hasil SQL ke Array Chart
            foreach (var item in actualsFromDb)
            {
                // Konversi Bulan Kalender (1-12) ke Index Array Fiskal (0-11)
                // April(4) -> 0 ... Maret(3) -> 11
                int arrayIndex = (item.Key - 4 + 12) % 12;
                actualData[arrayIndex] = Math.Round(item.Value, 1);
            }

            // 2. AMBIL DATA PLAN (EF Core)
            var planQuery = _context.LossTimePlans.AsQueryable();

            // Filter Tahun Fiskal
            planQuery = planQuery.Where(x =>
                (x.Year == SelectedYear && x.Month >= 4) ||
                (x.Year == SelectedYear + 1 && x.Month <= 3)
            );

            // Filter Mesin untuk Plan
            if (MachineLine != "All")
            {
                planQuery = planQuery.Where(x => x.MachineLine == MachineLine);
            }
            // Jika "All", kita tidak filter by MachineLine, otomatis EF akan men-SUM semua mesin yang ada

            var plansFromDb = planQuery.ToList()
                .GroupBy(x => x.Month)
                .Select(g => new { Month = g.Key, Total = g.Sum(x => x.TargetMinutes) })
                .ToList();

            foreach (var item in plansFromDb)
            {
                int arrayIndex = (item.Month - 4 + 12) % 12;
                planData[arrayIndex] = Math.Round(item.Total, 1);
            }

            // 3. Output JSON
            var chartPayload = new
            {
                Labels = months,
                ActualData = actualData,
                PlanData = planData
            };

            ChartDataJson = System.Text.Json.JsonSerializer.Serialize(chartPayload);
        }

        // --- QUERY SQL MANUAL (Perbaikan Logic 'All') ---
        private Dictionary<int, double> GetMonthlyActualData(int fiscalYear, string line)
        {
            var result = new Dictionary<int, double>();
            var connectionString = _context.Database.GetDbConnection().ConnectionString;

            DateTime startDate = new DateTime(fiscalYear, 4, 1);
            DateTime endDate = new DateTime(fiscalYear + 1, 3, 31);

            // Query dasar: Ambil Bulan dan Sum Durasi
            string query = @"
                SELECT 
                    MONTH(Date) as MonthVal,
                    SUM(LossTime) / 60.0 as TotalMinutes
                FROM AssemblyLossTime
                WHERE Date >= @Start AND Date <= @End
            ";

            // Tambahkan filter mesin JIKA user tidak memilih "All"
            // Pastikan nilai 'line' (misal 'MCH1-01') SAMA PERSIS dengan isi kolom 'Line' di database Anda
            if (line != "All")
            {
                query += " AND MachineCode = @MachineCode";
            }

            query += " GROUP BY MONTH(Date)";

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@Start", startDate);
                        cmd.Parameters.AddWithValue("@End", endDate);

                        if (line != "All")
                        {
                            cmd.Parameters.AddWithValue("@MachineCode", line);
                        }

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                int m = Convert.ToInt32(reader["MonthVal"]);
                                double val = reader["TotalMinutes"] != DBNull.Value ? Convert.ToDouble(reader["TotalMinutes"]) : 0;
                                result[m] = val;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("SQL Error: " + ex.Message);
            }

            return result;
        }

        // --- IMPORT EXCEL (Melalui Pop-up) ---
        public async Task<IActionResult> OnPostImportExcelAsync()
        {
            if (UploadedExcel == null || UploadedExcel.Length == 0)
            {
                TempData["Error"] = "File Excel belum dipilih.";
                return RedirectToPage(new { SelectedYear, MachineLine });
            }

            // Validasi: Machine Line harus dipilih di dalam Pop-up, tidak boleh "All" saat upload
            if (UploadMachineLine == "All" || string.IsNullOrEmpty(UploadMachineLine))
            {
                TempData["Error"] = "Saat Upload, anda harus memilih Mesin Spesifik (CU atau CS) di dalam Pop-up.";
                return RedirectToPage(new { SelectedYear, MachineLine });
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                using (var stream = new MemoryStream())
                {
                    await UploadedExcel.CopyToAsync(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var sheet = package.Workbook.Worksheets[0];
                        int rowCount = sheet.Dimension.Rows;

                        // Mapping Kolom Excel ke Bulan
                        var monthColMap = new Dictionary<int, int> {
                            { 4, 3 }, { 5, 5 }, { 6, 7 }, { 7, 9 }, { 8, 11 }, { 9, 13 },
                            { 10, 15 }, { 11, 17 }, { 12, 19 }, { 1, 21 }, { 2, 23 }, { 3, 25 }
                        };

                        var newPlans = new List<LossTimePlan>();

                        for (int row = 4; row <= rowCount; row++)
                        {
                            var catName = sheet.Cells[row, 2].Text?.Trim();
                            if (string.IsNullOrEmpty(catName) || catName.ToLower().Contains("loss category")) continue;

                            foreach (var map in monthColMap)
                            {
                                var valText = sheet.Cells[row, map.Value].Text;
                                if (double.TryParse(valText, out double val) && val > 0)
                                {
                                    int dataYear = (map.Key >= 4) ? SelectedYear : SelectedYear + 1;

                                    newPlans.Add(new LossTimePlan
                                    {
                                        Category = catName,
                                        MachineLine = this.UploadMachineLine, // Gunakan pilihan dari Pop-up
                                        Month = map.Key,
                                        Year = dataYear,
                                        TargetMinutes = val
                                    });
                                }
                            }
                        }

                        // Hapus data lama (hanya untuk mesin yang dipilih di pop-up)
                        var dataToDelete = _context.LossTimePlans
                            .Where(x => x.MachineLine == this.UploadMachineLine &&
                                        ((x.Year == SelectedYear && x.Month >= 4) ||
                                         (x.Year == SelectedYear + 1 && x.Month <= 3)));

                        _context.LossTimePlans.RemoveRange(dataToDelete);

                        if (newPlans.Any())
                        {
                            _context.LossTimePlans.AddRange(newPlans);
                            await _context.SaveChangesAsync();
                            TempData["Success"] = $"Berhasil import Plan untuk {UploadMachineLine}.";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Gagal Import: " + ex.Message;
            }

            return RedirectToPage(new { SelectedYear, MachineLine });
        }

        public IActionResult OnGetDownloadTemplate()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "planlosstime", "Template BP Loss Time.xlsx");
            if (!System.IO.File.Exists(filePath)) return NotFound("File template tidak ditemukan di server.");
            var bytes = System.IO.File.ReadAllBytes(filePath);
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "LossTimePlan_Template.xlsx");
        }
    }
}