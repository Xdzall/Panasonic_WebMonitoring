using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using System.Globalization;
using System;
using System.Linq;
using System.Text.Json;
using MonitoringSystem.Data;
using System.Drawing;

namespace MonitoringSystem.Pages.LossTimeReport
{
    public class indexModel : PageModel
    {
        // BARU: Properti untuk menampung tahun yang dipilih dari UI
        [BindProperty(SupportsGet = true)]
        public int SelectedYear { get; set; } = DateTime.Today.Year;

        private readonly ApplicationDbContext _context;

        public indexModel(ApplicationDbContext context)
        {
            _context = context;
        }
        [BindProperty]
        public string MachineLine { get; set; } = "All";
        public class LossReportData
        {
            public Dictionary<string, Dictionary<int, int>> MonthlyLosses { get; set; } = new Dictionary<string, Dictionary<int, int>>();
            public Dictionary<int, int> TotalLossByMonth { get; set; } = new Dictionary<int, int>();
            public Dictionary<int, int> FixedLossByMonth { get; set; } = new Dictionary<int, int>();
        }

        public LossReportData ReportData { get; set; } = new LossReportData();
        public string ChartDataJson { get; set; } = "{}";
        public List<string> MonthLabels { get; set; } = new List<string>();
        public List<double> ActualLossByMonth { get; set; } = new List<double>();
        public List<double> ActualWtByMonth { get; set; } = new List<double>();
        public List<string> TotalLossVsTotalWtPresentage { get; set; } = new List<string>();
        public List<string> ActualLossVsActualWtPresentage { get; set; } = new List<string>();
        public List<double> TotalWorkingTimeByMonth { get; set; } = new List<double>();
        public List<string> AllCategories { get; set; } = new List<string>
        {
            "Change Model", "Material Shortage External", "MP Adjustment",
            "Material Shortage Internal", "Material Shortage Inhouse", "Quality Trouble",
            "Machine Trouble", "Rework", "Loss Awal Hari", "Other"
        };

        public string connectionString = "Server=10.83.33.103;trusted_connection=false;Database=PROMOSYS;User Id=sa;Password=sa;Persist Security Info=False;Encrypt=False";

        // DIUBAH: Logika utama untuk menentukan rentang berdasarkan TAHUN KALENDER
        public void OnGet()
        {
            try
            {
                //rentang berdasarkan TAHUN FISKAL(April -Maret)
                DateTime fisicalStartDate = new DateTime(SelectedYear, 4, 1);
                DateTime fisicalEndDate = fisicalStartDate.AddYears(1).AddDays(-1);

                LoadData(fisicalStartDate, fisicalEndDate);
                CalculateTotalWorkingTime(fisicalStartDate);
                PrepareChartData(fisicalStartDate);
                CalculateDeriveMetrics();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in OnGet: {ex.Message}");
                InitializeEmptyData();
            }
        }

        public IActionResult OnPost()
        {
            // Handler untuk form submit, cukup panggil OnGet untuk reload data
            OnGet();
            return Page();
        }

        private void CalculateTotalWorkingTime(DateTime fisicalStartDate)
        {
            TotalWorkingTimeByMonth = new List<double>();
            var holidays = GetHolidays(fisicalStartDate.Year, fisicalStartDate.Year + 1);
            DateTime today = DateTime.Today;

            for (int i = 0; i < 12; i++)
            {
                DateTime currentMonthStart = fisicalStartDate.AddMonths(i);

                if (currentMonthStart.Year > today.Year || (currentMonthStart.Year == today.Year && currentMonthStart.Month > today.Month))
                {
                TotalWorkingTimeByMonth.Add(0);
                continue;
                }
                int workDays = 0;
                int daysToCount;
                
                if (currentMonthStart.Year == today.Year && currentMonthStart.Month == today.Month)
                {
                    daysToCount = today.Day;
                }
                else
                {
                    daysToCount = DateTime.DaysInMonth(currentMonthStart.Year, currentMonthStart.Month);
                }

                for (int day =1; day <= daysToCount; day++)
                {
                    DateTime currentDate = new DateTime(currentMonthStart.Year, currentMonthStart.Month, day);
                    if (currentDate.DayOfWeek != DayOfWeek.Saturday &&
                        currentDate.DayOfWeek != DayOfWeek.Sunday &&
                        !holidays.Contains(currentDate.Date))
                    {
                        workDays++;
                    }
                }
                TotalWorkingTimeByMonth.Add(workDays * 473);
            }
        }

        private void CalculateDeriveMetrics()
        {
            for (int i = 0; i < 12; i++)
            {
                double totalLoss = SecondsToMinutes(ReportData.TotalLossByMonth[i]);
                double fixedLoss = SecondsToMinutes(ReportData.FixedLossByMonth[i]);
                double totalWt = TotalWorkingTimeByMonth[i];

                double actualLoss = totalLoss - fixedLoss;
                double actualWt = totalWt - fixedLoss;

                ActualLossByMonth.Add(actualLoss > 0 ? actualLoss : 0);
                ActualWtByMonth.Add(actualWt > 0 ? actualWt : 0);

                TotalLossVsTotalWtPresentage.Add(totalWt > 0 ? (totalLoss / totalWt).ToString("P1") : "0%");
                ActualLossVsActualWtPresentage.Add(actualWt > 0 ? (actualLoss / actualWt).ToString("P1") : "0%");
            }
        }

        private HashSet<DateTime>GetHolidays(int startYear, int endYear)
        {
            var holidays = new HashSet<DateTime>();
            for (int year = startYear; year <= endYear; year++)
            {
                holidays.Add(new DateTime(year, 1, 1));
                holidays.Add(new DateTime(year, 5, 1));
                holidays.Add(new DateTime(year, 8, 17));
                holidays.Add(new DateTime(year, 12, 25));
                //tambahkan hari libur lainnya jika ada
            }
            return holidays;
        }

        // DIUBAH: Logika pengelompokan data
        private void LoadData(DateTime startDate, DateTime endDate)
        {
            InitializeEmptyData();

            string query = @"
                SELECT [Date], [Reason], [LossTime] FROM AssemblyLossTime 
                WHERE [Date] >= @StartDate AND [Date] < @EndDate";

            if (!string.Equals(MachineLine, "All", StringComparison.OrdinalIgnoreCase))
            {
                query += " AND [MachineCode] = @MachineCode";
            }

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@StartDate", startDate);
                        command.Parameters.AddWithValue("@EndDate", endDate.AddDays(1)); // Include the last day

                        if (!string.Equals(MachineLine, "All", StringComparison.OrdinalIgnoreCase))
                        {
                            command.Parameters.AddWithValue("@MachineCode", MachineLine);
                        }

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                DateTime date = reader.GetDateTime(0);
                                string reason = reader.GetString(1);
                                int lossDuration = reader.GetInt32(2);

                                int monthIndex = ((date.Year - startDate.Year) * 12) + date.Month - startDate.Month;

                                if (monthIndex >= 0 && monthIndex < 12)
                                {
                                    string category = CategorizeReason(reason);
                                    ReportData.MonthlyLosses[category][monthIndex] += lossDuration;
                                    ReportData.TotalLossByMonth[monthIndex] += lossDuration;

                                    if (category == "Loss Awal Hari")
                                    {
                                        ReportData.FixedLossByMonth[monthIndex] += lossDuration;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading data: {ex.Message}");
            }
        }

        // DIUBAH: Label bulan menjadi nama bulan standar
        private void PrepareChartData(DateTime fiscalStartDate)
        {
            MonthLabels = Enumerable.Range(0, 12)
                .Select(i => fiscalStartDate.AddMonths(i).ToString("MMMM", CultureInfo.CurrentCulture))
                .ToList();

            var datasets = new List<object>();
            var backgroundColors = new[] { "#FF6384", "#36A2EB", "#FFCE56", "#4BC0C0", "#9966FF", "#FF9F40", "#C9CBCF" };
            int colorIndex = 0;

            foreach (var category in AllCategories)
            {
                var data = ReportData.MonthlyLosses[category]
                .Select(kvp => Math.Ceiling(kvp.Value / 60.0))
                .ToArray();

                datasets.Add(new
                {
                    label = category,
                    data = data,
                    backgroundColor = backgroundColors[colorIndex % backgroundColors.Length],
                    stack = "loss"
                });
                colorIndex++;
            }

            ChartDataJson = JsonSerializer.Serialize(new { labels = MonthLabels, datasets });
        }

        private void InitializeEmptyData()
        {
            ReportData = new LossReportData();
            foreach (var category in AllCategories)
            {
                ReportData.MonthlyLosses[category] = new Dictionary<int, int>();
                for (int i = 0; i < 12; i++) ReportData.MonthlyLosses[category][i] = 0;
            }

            // PERBAIKAN: Menambahkan inisialisasi untuk semua dictionary sebelum digunakan
            ReportData.TotalLossByMonth = new Dictionary<int, int>();
            ReportData.FixedLossByMonth = new Dictionary<int, int>();

            for (int i = 0; i < 12; i++)
            {
                ReportData.TotalLossByMonth[i] = 0;
                ReportData.FixedLossByMonth[i] = 0;
            }
        }

        private string CategorizeReason(string reason)
        {
            reason = reason?.ToLower() ?? "";
            if (reason.Contains("change model")) return "Change Model";
            if (reason.Contains("material shortage external")) return "Material Shortage External";
            if (reason.Contains("mp adjustment")) return "MP Adjustment";
            if (reason.Contains("material shortage internal")) return "Material Shortage Internal";
            if (reason.Contains("material shortage inhouse")) return "Material Shortage Inhouse";
            if (reason.Contains("quality trouble")) return "Quality Trouble";
            if (reason.Contains("machine trouble")) return "Machine Trouble";
            if (reason.Contains("rework")) return "Rework";
            if (reason.Contains("loss awal hari")) return "Loss Awal Hari";
            return "Other";
        }

        public double SecondsToMinutes(int seconds)
        {
            return Math.Ceiling(seconds / 60.0);
        }
    }
}