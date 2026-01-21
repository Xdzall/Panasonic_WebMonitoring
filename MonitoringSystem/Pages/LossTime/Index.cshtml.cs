using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using ClosedXML.Excel;
using MonitoringSystem.Data;
using System.Text.Json;
using System.Text;

namespace MonitoringSystem.Pages.LossTime
{
    public class IndexModel : PageModel
    {
        private readonly ApplicationDbContext _context;

        public IndexModel(ApplicationDbContext context)
        {
            _context = context;
        }

        public List<LossTimeRecord> LossTimeData { get; set; } = new List<LossTimeRecord>();
        public int TotalDuration { get; set; }
        public int CurrentPage { get; set; } = 1;
        public int PageSize { get; set; } = 10;
        public int TotalPages => (int)Math.Ceiling((double)TotalRecords / PageSize);
        public int TotalRecords { get; set; }
        public bool HasDataToDisplay => TotalRecords > 0;

        [BindProperty]
        public DateTime StartSelectedDate { get; set; } = DateTime.Today;

        [BindProperty]
        public DateTime EndSelectedDate { get; set; } = DateTime.Today;

        [BindProperty]
        public int SelectedMonth { get; set; } = DateTime.Today.Month;
        [BindProperty]
        public int SelectedYear { get; set; } = DateTime.Today.Year;

        [BindProperty]
        public string MachineLine { get; set; } = "All";

        [BindProperty]
        public List<string> SelectedShifts { get; set; } = new List<string> { "1", "2", "3" };

        [BindProperty]
        public int SelectedPageSize { get; set; } = 10;

        [BindProperty]
        public string AdditionalBreakTime1Start { get; set; } = "";
        [BindProperty]
        public string AdditionalBreakTime1End { get; set; } = "";
        [BindProperty]
        public string AdditionalBreakTime2Start { get; set; } = "";
        [BindProperty]
        public string AdditionalBreakTime2End { get; set; } = "";

        public bool IsFiltering { get; set; } = false;
        public Dictionary<string, int> CategorySummary { get; set; } = new Dictionary<string, int>();
        public string ChartDataJson { get; set; }
        public string DailyChartDataJson { get; set; }

        public List<string> AllCategories { get; set; } = new List<string>
        {
            "Model Changing Loss",
            "Material Shortage External",
            "Man Power Adjustment",
            "Material Shortage Internal",
            "Material Shortage Inhouse",
            "Quality Trouble",
            "Machine & Tools Trouble",
            "Rework",
            "Morning Assembly",
            "Reason Not Fill"
        };

        public Dictionary<string, string> CategoryAbbreviations = new()
        {
            { "Model Changing Loss", "Mdl Change" },
            { "Material Shortage External", "Mtrl Shortage Ex" },
            { "Man Power Adjustment", "MP Adjust" },
            { "Material Shortage Internal", "Mtrl Shortage Int" },
            { "Material Shortage Inhouse", "Mtrl Shortage Inhs" },
            { "Quality Trouble", "Quality" },
            { "Machine & Tools Trouble", "MC Trouble" },
            { "Rework", "Rework" },
            { "Morning Assembly", "Morning Assy" },
            { "Reason Not Fill", "Reason NF" }
        };


        private readonly Dictionary<string, string> CategoryColors = new Dictionary<string, string>
        {
            { "Model Changing Loss", "#FF6384" },
            { "Material Shortage External", "#36A2EB" },
            { "Man Power Adjustment", "#FFCE56" },
            { "Material Shortage Internal", "#4BC0C0" },
            { "Material Shortage Inhouse", "#9966FF" },
            { "Quality Trouble", "#FF9F40" },
            { "Machine & Tools Trouble", "#C9CBCF" },
            { "Rework", "#FF9F80" },
            { "Morning Assembly", "#198754" },
            { "Reason Not Fill", "#77DD77" }
        };

        private readonly List<(TimeSpan Start, TimeSpan End)> FixedBreakTimes = new List<(TimeSpan, TimeSpan)>
        {
            (new TimeSpan(7, 0, 0), new TimeSpan(7, 5, 0)),
            (new TimeSpan(9, 30, 0), new TimeSpan(9, 35, 0)),
            (new TimeSpan(15, 30, 0), new TimeSpan(15, 35, 0)),
            (new TimeSpan(18, 15, 0), new TimeSpan(18, 45, 0))
        };

        public string connectionString = "Server=XDZALL\\SQLEXPRESS;Database=PROMOSYS;Trusted_Connection=True;Encrypt=False";

        public void OnGet(int pageNumber = 1, int pageSize = 10)
        {
            CurrentPage = pageNumber;
            PageSize = pageSize;
            SelectedPageSize = pageSize;
            SetDatesFromMonthYear();
            LoadBreakTimeForToday();
            LoadData();
        }

        public void SetDatesFromMonthYear()
        {
            StartSelectedDate = new DateTime(SelectedYear, SelectedMonth, 1);
            EndSelectedDate = StartSelectedDate.AddMonths(1).AddDays(-1);
        }

        public IActionResult OnPostFilter()
        {
            CurrentPage = 1;
            PageSize = SelectedPageSize;
            SetDatesFromMonthYear();
            if (SelectedShifts == null || !SelectedShifts.Any()) SelectedShifts = new List<string> { "1", "2", "3" };
            LoadBreakTimeForToday();
            IsFiltering = true;
            LoadData();
            return Page();
        }

        public IActionResult OnPostChangePage(int pageNumber, int pageSize, int selectedMonth, int selectedYear,
            string machineLine, List<string> selectedShifts,
            string additionalBreakTime1Start, string additionalBreakTime1End,
            string additionalBreakTime2Start, string additionalBreakTime2End)
        {
            CurrentPage = pageNumber;
            PageSize = pageSize;
            SelectedMonth = selectedMonth;
            SelectedYear = selectedYear;
            SetDatesFromMonthYear();
            MachineLine = machineLine;
            SelectedShifts = selectedShifts ?? new List<string> { "1", "2", "3" };
            AdditionalBreakTime1Start = additionalBreakTime1Start;
            AdditionalBreakTime1End = additionalBreakTime1End;
            AdditionalBreakTime2Start = additionalBreakTime2Start;
            AdditionalBreakTime2End = additionalBreakTime2End;
            LoadData();
            return Page();
        }

        public IActionResult OnPostReset()
        {
            ModelState.Clear();
            SelectedMonth = DateTime.Today.Month;
            SelectedYear = DateTime.Today.Year;
            SetDatesFromMonthYear();
            MachineLine = "All";
            SelectedShifts = new List<string> { "1", "2", "3" };
            SelectedPageSize = 10;
            PageSize = 10;
            IsFiltering = false;
            CurrentPage = 1;
            LoadBreakTimeForToday();
            LoadData();
            return Page();
        }

        private void LoadBreakTimeForToday()
        {
            var today = DateTime.Today;
            var latestBreakTime = _context.AdditionalBreakTimes.Where(bt => bt.Date == today).OrderByDescending(bt => bt.CreatedAt).FirstOrDefault();
            if (latestBreakTime != null)
            {
                AdditionalBreakTime1Start = latestBreakTime.BreakTime1Start?.ToString(@"hh\:mm");
                AdditionalBreakTime1End = latestBreakTime.BreakTime1End?.ToString(@"hh\:mm");
                AdditionalBreakTime2Start = latestBreakTime.BreakTime2Start?.ToString(@"hh\:mm");
                AdditionalBreakTime2End = latestBreakTime.BreakTime2End?.ToString(@"hh\:mm");
            }
        }

        private List<(TimeSpan Start, TimeSpan End)> GetAllBreakTimes()
        {
            var breakTimes = new List<(TimeSpan Start, TimeSpan End)>();
            breakTimes.AddRange(FixedBreakTimes);
            if (!string.IsNullOrEmpty(AdditionalBreakTime1Start) && !string.IsNullOrEmpty(AdditionalBreakTime1End))
                if (TryParseTimeSpan(AdditionalBreakTime1Start, out TimeSpan start1) && TryParseTimeSpan(AdditionalBreakTime1End, out TimeSpan end1)) breakTimes.Add((start1, end1));
            if (!string.IsNullOrEmpty(AdditionalBreakTime2Start) && !string.IsNullOrEmpty(AdditionalBreakTime2End))
                if (TryParseTimeSpan(AdditionalBreakTime2Start, out TimeSpan start2) && TryParseTimeSpan(AdditionalBreakTime2End, out TimeSpan end2)) breakTimes.Add((start2, end2));
            return breakTimes;
        }

        private bool TryParseTimeSpan(string timeString, out TimeSpan result)
        {
            string[] formats = { "HH:mm", "H:mm", "HH:mm:ss", "H:mm:ss" };
            if (TimeSpan.TryParseExact(timeString, formats, null, out result)) return true;
            if (DateTime.TryParse(timeString, out DateTime dateTime)) { result = dateTime.TimeOfDay; return true; }
            result = TimeSpan.Zero; return false;
        }

        private bool IsInBreakTime(TimeSpan startTime, TimeSpan endTime, List<(TimeSpan Start, TimeSpan End)> breakTimes)
        {
            foreach (var (breakStart, breakEnd) in breakTimes)
                if ((startTime >= breakStart && startTime <= breakEnd) || (endTime >= breakStart && endTime <= breakEnd) || (startTime <= breakStart && endTime >= breakEnd)) return true;
            return false;
        }

        private void LoadData()
        {
            var breakTimes = GetAllBreakTimes();

            // 1. Ambil Data untuk Chart Summary & Harian (Data Bulan Ini)
            var currentRecords = GetLossTimeRecords(StartSelectedDate, EndSelectedDate, breakTimes);

            // 2. Ambil Data Bulan Lalu untuk Chart Summary
            DateTime prevMonthDate = StartSelectedDate.AddMonths(-1);
            DateTime lastMonthStart = new DateTime(prevMonthDate.Year, prevMonthDate.Month, 1);
            DateTime lastMonthEnd = lastMonthStart.AddMonths(1).AddDays(-1);
            var lastMonthRecords = GetLossTimeRecords(lastMonthStart, lastMonthEnd, breakTimes);

            // 3. Siapkan Chart Summary (Shift & Comparison)
            PrepareSummaryChartData(currentRecords, lastMonthRecords);

            // 4. Siapkan Chart Harian (Daily Stacked)
            PrepareDailyChartData(currentRecords);

            // 5. Load Data Tabel (Pagination)
            LoadPaginatedData(breakTimes);
        }

        private List<LossTimeRecord> GetLossTimeRecords(DateTime start, DateTime end, List<(TimeSpan Start, TimeSpan End)> breakTimes)
        {
            var records = new List<LossTimeRecord>();
            string query = BuildQueryBase();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    AddQueryParameters(command, start, end);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            TimeSpan startTime = reader.GetTimeSpan(reader.GetOrdinal("StartTime"));
                            TimeSpan endTime = reader.GetTimeSpan(reader.GetOrdinal("EndTime"));
                            if (IsInBreakTime(startTime, endTime, breakTimes)) continue;

                            string reason = reader.IsDBNull(reader.GetOrdinal("Reason")) ? string.Empty : reader.GetString(reader.GetOrdinal("Reason"));
                            records.Add(new LossTimeRecord
                            {
                                Date = reader.GetDateTime(reader.GetOrdinal("Date")), // Penting untuk grouping harian
                                LossTime = reason,
                                Duration = reader.IsDBNull(reader.GetOrdinal("LossTime")) ? 0 : reader.GetInt32(reader.GetOrdinal("LossTime")),
                                Shift = reader.IsDBNull(reader.GetOrdinal("Shift")) ? string.Empty : reader.GetString(reader.GetOrdinal("Shift")),
                                Category = CategorizeReason(reason)
                            });
                        }
                    }
                }
            }
            return records;
        }
       
        private void PrepareSummaryChartData(List<LossTimeRecord> currentRecords, List<LossTimeRecord> lastMonthRecords)
        {
            try
            {
                var categoryStats = AllCategories.Select(cat => new
                {
                    Name = cat,
                    S1 = currentRecords.Where(r => r.Category == cat && r.Shift == "1").Sum(r => r.Duration),
                    S2 = currentRecords.Where(r => r.Category == cat && r.Shift == "2").Sum(r => r.Duration),
                    S3 = currentRecords.Where(r => r.Category == cat && r.Shift == "3").Sum(r => r.Duration),
                    TotalCurrent = currentRecords.Where(r => r.Category == cat).Sum(r => r.Duration),
                    TotalLast = lastMonthRecords.Where(r => r.Category == cat).Sum(r => r.Duration)
                }).ToList();

                var sortedStats = categoryStats.OrderByDescending(x => x.TotalCurrent).ToList();

                var chartData = new
                {
                    labels = sortedStats
                    .Select(x => CategoryAbbreviations.ContainsKey(x.Name)
                        ? CategoryAbbreviations[x.Name]
                        : x.Name)
                    .ToArray(),

                                fullLabels = sortedStats
                    .Select(x => x.Name)
                    .ToArray(),
                    shift1Data = sortedStats.Select(x => Math.Round(x.S1 / 60.0, 2)).ToArray(),
                    shift2Data = sortedStats.Select(x => Math.Round(x.S2 / 60.0, 2)).ToArray(),
                    shift3Data = sortedStats.Select(x => Math.Round(x.S3 / 60.0, 2)).ToArray(),
                    lastMonthData = sortedStats.Select(x => Math.Round(x.TotalLast / 60.0, 2)).ToArray()
                };
                ChartDataJson = JsonSerializer.Serialize(chartData);
            }
            catch (Exception) { ChartDataJson = "{}"; }
        }

        private void PrepareDailyChartData(List<LossTimeRecord> currentRecords)
        {
            try
            {
                // 1. Dapatkan daftar hari dalam bulan tersebut (1 s/d 30/31)
                int daysInMonth = DateTime.DaysInMonth(SelectedYear, SelectedMonth);
                var days = Enumerable.Range(1, daysInMonth).ToArray();

                // 2. Group data berdasarkan Hari dan Kategori
                var dailyGroups = currentRecords
                    .GroupBy(r => new { Day = r.Date.Day, r.Category })
                    .ToDictionary(g => g.Key, g => g.Sum(x => x.Duration));

                // 3. Buat Datasets (Satu dataset per Kategori)
                var datasets = AllCategories.Select(category => new
                {
                    label = category,
                    data = days.Select(day =>
                    {
                        var key = new { Day = day, Category = category };
                        return dailyGroups.ContainsKey(key) ? Math.Round(dailyGroups[key] / 60.0, 2) : 0;
                    }).ToArray(),
                    backgroundColor = CategoryColors.ContainsKey(category) ? CategoryColors[category] : "#cccccc",
                    stack = "DayStack" // Stack ID yang sama agar bertumpuk
                }).ToList();

                // 4. Buat Object Chart
                var dailyChartData = new
                {
                    labels = days.Select(d => d.ToString()).ToArray(), // Label hari 1, 2, 3...
                    datasets = datasets
                };

                DailyChartDataJson = JsonSerializer.Serialize(dailyChartData);
            }
            catch (Exception) { DailyChartDataJson = "{}"; }
        }

        private void LoadPaginatedData(List<(TimeSpan Start, TimeSpan End)> breakTimes)
        {
            TotalRecords = GetTotalRecords(breakTimes);
            EnsureValidCurrentPage();
            string query = BuildQueryBase();
            query += " ORDER BY [Date] DESC, Time OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    AddQueryParameters(command, StartSelectedDate, EndSelectedDate);
                    command.Parameters.AddWithValue("@Offset", (CurrentPage - 1) * PageSize);
                    command.Parameters.AddWithValue("@PageSize", PageSize);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        LossTimeData.Clear();
                        while (reader.Read())
                        {
                            TimeSpan startTime = reader.GetTimeSpan(reader.GetOrdinal("StartTime"));
                            TimeSpan endTime = reader.GetTimeSpan(reader.GetOrdinal("EndTime"));
                            if (IsInBreakTime(startTime, endTime, breakTimes)) continue;
                            string reason = reader.IsDBNull(reader.GetOrdinal("Reason")) ? string.Empty : reader.GetString(reader.GetOrdinal("Reason"));
                            LossTimeData.Add(new LossTimeRecord
                            {
                                Nomor = reader.IsDBNull(reader.GetOrdinal("Id")) ? 0 : reader.GetInt32(reader.GetOrdinal("Id")),
                                Date = reader.IsDBNull(reader.GetOrdinal("Date")) ? DateTime.MinValue : reader.GetDateTime(reader.GetOrdinal("Date")),
                                LossTime = reason,
                                Start = startTime,
                                End = endTime,
                                Duration = reader.IsDBNull(reader.GetOrdinal("LossTime")) ? 0 : reader.GetInt32(reader.GetOrdinal("LossTime")),
                                Location = reader.IsDBNull(reader.GetOrdinal("MachineCode")) ? string.Empty : reader.GetString(reader.GetOrdinal("MachineCode")),
                                Shift = reader.IsDBNull(reader.GetOrdinal("Shift")) ? string.Empty : reader.GetString(reader.GetOrdinal("Shift")),
                                Category = CategorizeReason(reason),
                                DetailedReason = reader.IsDBNull(reader.GetOrdinal("DetailedReason")) ? null : reader.GetString(reader.GetOrdinal("DetailedReason"))
                            });
                        }
                    }
                }
            }
        }

        private void CalculateAllDataSummary(List<(TimeSpan Start, TimeSpan End)> breakTimes) => PrepareDailyChartData(GetLossTimeRecords(StartSelectedDate, EndSelectedDate, breakTimes));

        private string BuildQueryBase()
        {
            string query = @"
                SELECT Id, Date, Reason, DetailedReason, CAST(Time AS TIME) AS StartTime, CAST(EndDateTime AS TIME) AS EndTime, LossTime, MachineCode, 
                       CASE WHEN (DATEPART(HOUR, Time) = 7 AND DATEPART(MINUTE, Time) >= 0) OR (DATEPART(HOUR, Time) > 7 AND DATEPART(HOUR, Time) < 15) OR (DATEPART(HOUR, Time) = 15 AND DATEPART(MINUTE, Time) <= 45) THEN '1'
                           WHEN (DATEPART(HOUR, Time) = 15 AND DATEPART(MINUTE, Time) > 45) OR (DATEPART(HOUR, Time) > 15 AND DATEPART(HOUR, Time) < 23) OR (DATEPART(HOUR, Time) = 23 AND DATEPART(MINUTE, Time) <= 15) THEN '2'
                           ELSE '3' END AS Shift
                FROM AssemblyLossTime WHERE [Date] >= @StartDate AND [Date] <= @EndDate";
            if (!string.IsNullOrEmpty(MachineLine) && MachineLine != "All") query += " AND MachineCode = @MachineLine";
            if (SelectedShifts != null && SelectedShifts.Any() && SelectedShifts.Count < 3)
            {
                query += " AND (";
                List<string> shiftConditions = new List<string>();
                foreach (var shift in SelectedShifts)
                {
                    if (shift == "1") shiftConditions.Add("((DATEPART(HOUR, Time) = 7 AND DATEPART(MINUTE, Time) >= 0) OR (DATEPART(HOUR, Time) > 7 AND DATEPART(HOUR, Time) < 15) OR (DATEPART(HOUR, Time) = 15 AND DATEPART(MINUTE, Time) <= 45))");
                    else if (shift == "2") shiftConditions.Add("((DATEPART(HOUR, Time) = 15 AND DATEPART(MINUTE, Time) > 45) OR (DATEPART(HOUR, Time) > 15 AND DATEPART(HOUR, Time) < 23) OR (DATEPART(HOUR, Time) = 23 AND DATEPART(MINUTE, Time) <= 15))");
                    else if (shift == "3") shiftConditions.Add("((DATEPART(HOUR, Time) = 23 AND DATEPART(MINUTE, Time) > 15) OR (DATEPART(HOUR, Time) >= 0 AND DATEPART(HOUR, Time) < 7))");
                }
                query += string.Join(" OR ", shiftConditions);
                query += ")";
            }
            return query;
        }

        private void AddQueryParameters(SqlCommand command, DateTime start, DateTime end)
        {
            command.Parameters.AddWithValue("@StartDate", start);
            command.Parameters.AddWithValue("@EndDate", end);
            if (!string.IsNullOrEmpty(MachineLine) && MachineLine != "All") command.Parameters.AddWithValue("@MachineLine", MachineLine);
        }

        private int GetTotalRecords(List<(TimeSpan Start, TimeSpan End)> breakTimes)
        {
            int count = 0;
            string query = BuildQueryBase();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    AddQueryParameters(command, StartSelectedDate, EndSelectedDate);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            TimeSpan startTime = reader.GetTimeSpan(reader.GetOrdinal("StartTime"));
                            TimeSpan endTime = reader.GetTimeSpan(reader.GetOrdinal("EndTime"));
                            if (!IsInBreakTime(startTime, endTime, breakTimes)) count++;
                        }
                    }
                }
            }
            return count;
        }

        private void EnsureValidCurrentPage() { if (TotalRecords == 0) { CurrentPage = 1; return; } int maxPages = (int)Math.Ceiling((double)TotalRecords / PageSize); if (CurrentPage > maxPages) CurrentPage = maxPages; else if (CurrentPage < 1) CurrentPage = 1; }

        private string CategorizeReason(string reason)
        {
            reason = reason?.ToLower() ?? "";
            if (reason.Contains("model changing loss")) return "Model Changing Loss";
            else if (reason.Contains("material shortage external")) return "Material Shortage External";
            else if (reason.Contains("man power adjustment")) return "Man Power Adjustment";
            else if (reason.Contains("material shortage internal")) return "Material Shortage Internal";
            else if (reason.Contains("material shortage inhouse")) return "Material Shortage Inhouse";
            else if (reason.Contains("quality trouble")) return "Quality Trouble";
            else if (reason.Contains("machine & tools trouble")) return "Machine & Tools Trouble";
            else if (reason.Contains("rework")) return "Rework";
            else if (reason.Contains("morning assembly")) return "Morning Assembly";
            else return "Reason Not Fill";
        }

        public int GetTotalDurationAllCategories() => CategorySummary.Values.Sum();
        public double SecondsToMinutes(int seconds) => Math.Round(seconds / 60.0, 2);
        public List<int> GetPageSizeOptions() => new List<int> { 10 };

        public IActionResult OnPostExportExcel()
        {
            // Logic Export Excel sama seperti sebelumnya, disingkat disini untuk fokus pada Chart
            LoadBreakTimeForToday();
            SetDatesFromMonthYear();
            var breakTimes = GetAllBreakTimes();
            string query = BuildQueryBase();
            query += " ORDER BY [Date] DESC, Time";
            var exportData = GetLossTimeRecords(StartSelectedDate, EndSelectedDate, breakTimes);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Loss Time Data");
                worksheet.Cell(1, 1).Value = "No"; worksheet.Cell(1, 2).Value = "Date"; worksheet.Cell(1, 3).Value = "Category";
                worksheet.Cell(1, 4).Value = "Start Time"; worksheet.Cell(1, 5).Value = "End Time"; worksheet.Cell(1, 6).Value = "Duration (Sec)";
                worksheet.Cell(1, 7).Value = "Location"; worksheet.Cell(1, 8).Value = "Shift"; worksheet.Cell(1, 9).Value = "Detailed Reason";

                int row = 2; int index = 1;
                foreach (var item in exportData)
                {
                    worksheet.Cell(row, 1).Value = index++; worksheet.Cell(row, 2).Value = item.Date; worksheet.Cell(row, 3).Value = item.Category;
                    worksheet.Cell(row, 4).Value = item.Start.ToString(@"hh\:mm\:ss"); worksheet.Cell(row, 5).Value = item.End.ToString(@"hh\:mm\:ss");
                    worksheet.Cell(row, 6).Value = item.Duration; worksheet.Cell(row, 7).Value = item.Location;
                    worksheet.Cell(row, 8).Value = item.Shift; worksheet.Cell(row, 9).Value = item.DetailedReason;
                    row++;
                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"LossTime_{StartSelectedDate:yyyyMMdd}-{EndSelectedDate:yyyyMMdd}.xlsx");
                }
            }
        }
    }

    public class LossTimeRecord
    {
        public int Nomor { get; set; }
        public DateTime Date { get; set; }
        public string LossTime { get; set; }
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public int Duration { get; set; }
        public string Location { get; set; }
        public string Shift { get; set; }
        public string Category { get; set; }    
        public string DetailedReason { get; set; }
    }
}