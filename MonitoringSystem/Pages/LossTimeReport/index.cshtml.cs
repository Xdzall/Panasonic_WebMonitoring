using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using System.Globalization;
using System;
using System.Linq;
using System.Text.Json;
using MonitoringSystem.Data;

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

        public class LossReportData
        {
            public Dictionary<string, Dictionary<int, int>> MonthlyLosses { get; set; } = new Dictionary<string, Dictionary<int, int>>();
            public Dictionary<int, int> TotalLossByMonth { get; set; } = new Dictionary<int, int>();
        }

        public LossReportData ReportData { get; set; } = new LossReportData();
        public string ChartDataJson { get; set; } = "{}";
        public List<string> MonthLabels { get; set; } = new List<string>();

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
                // Tentukan periode dari 1 Januari hingga 31 Desember dari tahun yang dipilih
                DateTime startDate = new DateTime(SelectedYear, 1, 1);
                DateTime endDate = new DateTime(SelectedYear, 12, 31);

                LoadData(startDate, endDate);
                PrepareChartData();
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

        // DIUBAH: Logika pengelompokan data
        private void LoadData(DateTime startDate, DateTime endDate)
        {
            InitializeEmptyData();

            string query = @"
                SELECT [Date], [Reason], [LossTime]
                FROM AssemblyLossTime 
                WHERE [Date] >= @StartDate AND [Date] <= @EndDate";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@StartDate", startDate);
                        command.Parameters.AddWithValue("@EndDate", endDate);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                DateTime date = reader.GetDateTime(0);
                                string reason = reader.GetString(1);
                                int lossDuration = reader.GetInt32(2);

                                // LOGIKA BARU: Indeks bulan didapat langsung dari nomor bulan (Jan=0, Feb=1, dst.)
                                int monthIndex = date.Month - 1;

                                string category = CategorizeReason(reason);

                                ReportData.MonthlyLosses[category][monthIndex] += lossDuration;
                                ReportData.TotalLossByMonth[monthIndex] += lossDuration;
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
        private void PrepareChartData()
        {
            MonthLabels = Enumerable.Range(1, 12)
                .Select(i => CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i))
                .ToList();

            var datasets = new List<object>();
            var backgroundColors = new[] { "#FF6384", "#36A2EB", "#FFCE56", "#4BC0C0", "#9966FF", "#FF9F40", "#C9CBCF" };
            int colorIndex = 0;

            foreach (var category in AllCategories)
            {
                if (ReportData.MonthlyLosses[category].Values.Sum() > 0)
                {
                    datasets.Add(new
                    {
                        label = category,
                        data = ReportData.MonthlyLosses[category].Values.Select(v => Math.Round(v / 60.0, 1)).ToArray(),
                        backgroundColor = backgroundColors[colorIndex % backgroundColors.Length],
                        stack = "loss"
                    });
                    colorIndex++;
                }
            }
            ChartDataJson = JsonSerializer.Serialize(new { labels = MonthLabels, datasets });
        }

        private void InitializeEmptyData()
        {
            ReportData = new LossReportData();
            foreach (var category in AllCategories)
            {
                ReportData.MonthlyLosses[category] = new Dictionary<int, int>();
                for (int i = 0; i < 12; i++)
                {
                    ReportData.MonthlyLosses[category][i] = 0;
                }
            }
            for (int i = 0; i < 12; i++)
            {
                ReportData.TotalLossByMonth[i] = 0;
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
            return Math.Round(seconds / 60.0, 1);
        }
    }
}