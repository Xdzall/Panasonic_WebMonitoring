using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System.Globalization;
using System.Text;

namespace MonitoringSystem.Pages.ProductionReport
{
    public class IndexModel : PageModel
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly IConfiguration _configuration;
        private string connectionString;

        public bool IsCurrentMonthView { get; private set; }
        public int DaysInMonth { get; private set; }
        public List<string> ChartLabels { get; private set; } = new List<string>();
        public List<decimal> NormalData { get; private set; } = new List<decimal>();
        public List<decimal> OvertimeData { get; private set; } = new List<decimal>();
        public List<int> PlanData { get; private set; } = new List<int>();

        //public List<int> EstimateData { get; private set; } = new List<int>();
        public List<int> NoOfDirectWorkers { get; private set; } = new List<int>();
        public List<int> DailyWorkTime { get; private set; } = new List<int>();
        public List<int> OvertimeOperators { get; private set; } = new List<int>();
        public List<int> OvertimeMinutes { get; private set; } = new List<int>();
        public List<int> DailyLossTime { get; private set; } = new List<int>();

        private class DailyData { public int Day { get; set; } public decimal LastNormalReading { get; set; } = 0; public decimal LastOvertimeReading { get; set; } = 0; public int Plan { get; set; } = 0; 
        //public int Estimate { get; set; } = 0; 
        public int NoOfOperator { get; set; } = 0; public int OtOperatorCount { get; set; } = 0; public TimeSpan LastOtTime { get; set; } = TimeSpan.Zero; }
        public class RestTime { public int Duration { get; set; } public TimeSpan StartTime { get; set; } public TimeSpan EndTime { get; set; } }

        public IndexModel(IWebHostEnvironment webHostEnvironment, IConfiguration configuration)
        {
            _webHostEnvironment = webHostEnvironment;
            _configuration = configuration;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [BindProperty(SupportsGet = true)] public int SelectedMonth { get; set; } = DateTime.Now.Month;
        [BindProperty(SupportsGet = true)] public int SelectedYear { get; set; } = DateTime.Now.Year;
        [BindProperty(SupportsGet = true)] public string MachineLine { get; set; } = "MCH1-01";
        [BindProperty(SupportsGet = true)] public List<string> SelectedShifts { get; set; } = new List<string>();

        public void OnGet()
        {
            if (!SelectedShifts.Any())
            {
                SelectedShifts = new List<string> { "1" };
            }
            LoadChartData();
        }

        public IActionResult OnPost(string submitButton)
        {
            if (submitButton == "reset")
            {
                return RedirectToPage(new { SelectedYear = DateTime.Now.Year, SelectedMonth = DateTime.Now.Month, MachineLine = "MCH1-01" });
            }
            return RedirectToPage(new
            {
                SelectedYear = this.SelectedYear,
                SelectedMonth = this.SelectedMonth,
                MachineLine = this.MachineLine,
                SelectedShifts = this.SelectedShifts
            });
        }

        [BindProperty]
        public IFormFile UploadedFile { get; set; }
        [BindProperty]
        public string TargetMachine { get; set; }
        [BindProperty]
        public int TargetMonth { get; set; }
        [BindProperty]
        public int TargetYear { get; set; }

        public async Task<IActionResult> OnPostUploadAsync()
        {
            if (UploadedFile == null || UploadedFile.Length == 0)
            {
                TempData["UploadError"] = "Silakan pilih file untuk diunggah.";
                return RedirectToPage();
            }

            List<List<string>> dataInMemory;
            try
            {
                dataInMemory = await ReadDataFromUpload(UploadedFile);
            }
            catch (Exception ex)
            {
                TempData["UploadError"] = $"Gagal membaca file: {ex.Message}";
                return RedirectToPage();
            }

            if (!dataInMemory.Any() || !dataInMemory.First().Any())
            {
                TempData["UploadError"] = "File yang diunggah kosong atau formatnya tidak dikenali.";
                return RedirectToPage();
            }

            string expectedPrefix = TargetMachine == "MCH1-01" ? "CU" : "CS";
            string selectedMachineName = TargetMachine == "MCH1-01" ? "CU" : "CS";
            string firstModelName = dataInMemory.First().First().Trim();

            if (!firstModelName.StartsWith(expectedPrefix, StringComparison.OrdinalIgnoreCase))
            {
                string fileContentPrefix = firstModelName.Split('-').FirstOrDefault() ?? "Tidak dikenal";
                TempData["UploadError"] = $"File berisi data '{fileContentPrefix}', tetapi Anda memilih machine '{selectedMachineName}'. Unggahan dibatalkan.";
                return RedirectToPage();
            }

            string machineType = TargetMachine == "MCH1-01" ? "cu" : "cs";
            string monthName = CultureInfo.GetCultureInfo("en-US").DateTimeFormat.GetMonthName(TargetMonth);

            var planCsvPath = Path.Combine(_webHostEnvironment.WebRootPath, "data", machineType, "plan", "csv", $"{monthName}_{TargetYear}_plan.csv");
            var planXlsxPath = Path.Combine(_webHostEnvironment.WebRootPath, "data", machineType, "plan", "xlsx", $"{monthName}_{TargetYear}_plan.xlsx");
            //var estimateCsvPath = Path.Combine(_webHostEnvironment.WebRootPath, "data", machineType, "estimate", "csv", $"{monthName}_{TargetYear}_estimate.csv");
            //var estimateXlsxPath = Path.Combine(_webHostEnvironment.WebRootPath, "data", machineType, "estimate", "xlsx", $"{monthName}_{TargetYear}_estimate.xlsx");

            Directory.CreateDirectory(Path.GetDirectoryName(planCsvPath));
            Directory.CreateDirectory(Path.GetDirectoryName(planXlsxPath));
            //Directory.CreateDirectory(Path.GetDirectoryName(estimateCsvPath));
            //Directory.CreateDirectory(Path.GetDirectoryName(estimateXlsxPath));

            WriteCsvFile(planCsvPath, dataInMemory);
            WriteXlsxFile(planXlsxPath, dataInMemory);

            string successMessage = $"Plan untuk {monthName} {TargetYear} telah berhasil diunggah dan diperbarui.";
            TempData["UploadSuccess"] = successMessage;

            return RedirectToPage(new { SelectedYear = TargetYear, SelectedMonth = TargetMonth, MachineLine = TargetMachine });
        }

        private async Task<List<List<string>>> ReadDataFromUpload(IFormFile file)
        {
            var data = new List<List<string>>();
            var extension = Path.GetExtension(file.FileName).ToLower();
            using var stream = new MemoryStream();
            await file.CopyToAsync(stream);
            stream.Position = 0;

            if (extension == ".csv")
            {
                using var reader = new StreamReader(stream);
                while (!reader.EndOfStream)
                {
                    var line = await reader.ReadLineAsync();
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        data.Add(line.Split(',').ToList());
                    }
                }
            }
            else if (extension == ".xlsx")
            {
                using var package = new ExcelPackage(stream);
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet != null)
                {
                    for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var rowData = new List<string>();
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            rowData.Add(worksheet.Cells[row, col].Text);
                        }
                        data.Add(rowData);
                    }
                }
            }
            else
            {
                throw new InvalidOperationException("Format file tidak didukung. Harap unggah file .csv atau .xlsx");
            }
            return data;
        }

        private void WriteCsvFile(string filePath, List<List<string>> data)
        {
            var lines = data.Select(row => string.Join(",", row));
            System.IO.File.WriteAllLines(filePath, lines, Encoding.UTF8);
        }

        private void WriteXlsxFile(string filePath, List<List<string>> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Data");
            for (int row = 0; row < data.Count; row++)
            {
                for (int col = 0; col < data[row].Count; col++)
                {
                    worksheet.Cells[row + 1, col + 1].Value = data[row][col];
                }
            }
            var fileInfo = new FileInfo(filePath);
            package.SaveAs(fileInfo);
        }

        public IActionResult OnGetDownloadTemplate(string type)
        {
            if (string.IsNullOrEmpty(type) || (type.ToLower() != "cu" && type.ToLower() != "cs"))
            {
                return NotFound("Invalid template type.");
            }
            string wwwrootPath = _webHostEnvironment.WebRootPath;
            var templateFileName = $"template_{type.ToLower()}.xlsx";
            var templateFilePath = Path.Combine(wwwrootPath, "data", type.ToLower(), "plan", $"{type.ToLower()}_plan_template.xlsx");
            if (System.IO.File.Exists(templateFilePath))
            {
                return File(System.IO.File.ReadAllBytes(templateFilePath), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", templateFileName);
            }
            return NotFound($"Template file not found.");
        }

        private void LoadChartData()
        {
            this.connectionString = _configuration.GetConnectionString("DefaultConnection");

            var dailyLosses = GetDailyLossTimeTotals();
            bool isCurrentMonthView = (SelectedYear == DateTime.Now.Year && SelectedMonth == DateTime.Now.Month);
            this.IsCurrentMonthView = isCurrentMonthView;
            string dateFilter = isCurrentMonthView ? "AND CAST(SDate AS DATE) <= @TodayDate" : "";
            string machineType = MachineLine == "MCH1-01" ? "cu" : "cs";
            string monthName = CultureInfo.GetCultureInfo("en-US").DateTimeFormat.GetMonthName(SelectedMonth);
            this.DaysInMonth = DateTime.DaysInMonth(SelectedYear, SelectedMonth);

            TimeSpan workDayStart = new TimeSpan(7, 5, 0);
            TimeSpan workDayEnd = new TimeSpan(15, 55, 0);

            for (int day = 1; day <= this.DaysInMonth; day++)
            {
                var currentDate = new DateTime(SelectedYear, SelectedMonth, day);
                var dayType = DetermineTypeOfDay(currentDate.DayOfWeek);
                int workTimeForThisDay = 0;

                if (isCurrentMonthView && currentDate.Date > DateTime.Now.Date)
                {
                    workTimeForThisDay = 0;
                }
                else if (dayType != "WEEKEND")
                {
                    // Pilih jadwal istirahat berdasarkan hari
                    var breaksForThisDay = (dayType == "FRIDAY") ? this.FridayBreakTimes : this.RegularDayBreakTimes;
                    var definedBreaks = breaksForThisDay.Select(b => new RestTime { StartTime = b.Start, EndTime = b.End }).ToList();

                    if (currentDate.Date < DateTime.Now.Date)
                    {
                        // Kalkulasi untuk satu hari penuh
                        int totalDuration = (int)(workDayEnd - workDayStart).TotalMinutes;
                        int totalRest = GetTotalRestTime(definedBreaks, workDayStart, workDayEnd, workDayEnd);
                        workTimeForThisDay = totalDuration - totalRest;
                    }
                    else // Ini untuk hari ini (dinamis)
                    {
                        var currentTime = DateTime.Now.TimeOfDay;
                        if (currentTime > workDayStart)
                        {
                            var effectiveEndTime = (currentTime > workDayEnd) ? workDayEnd : currentTime;
                            int totalDuration = (int)(effectiveEndTime - workDayStart).TotalMinutes;
                            int totalRest = GetTotalRestTime(definedBreaks, workDayStart, workDayEnd, currentTime);
                            workTimeForThisDay = totalDuration - totalRest;
                        }
                    }
                }

                DailyWorkTime.Add(workTimeForThisDay < 0 ? 0 : workTimeForThisDay);
            }

            var combinedData = Enumerable.Range(1, this.DaysInMonth).Select(day => new DailyData { Day = day }).ToList();
            var planFilePath = Path.Combine(_webHostEnvironment.WebRootPath, "data", machineType, "plan", "csv", $"{monthName}_{SelectedYear}_plan.csv");
            if (System.IO.File.Exists(planFilePath)) { var planValues = ReadDataFromCsv(planFilePath, this.DaysInMonth); for (int i = 0; i < planValues.Count; i++) { combinedData[i].Plan = planValues[i]; } }
            //var estimateFilePath = Path.Combine(_webHostEnvironment.WebRootPath, "data", machineType, "estimate", "csv", $"{monthName}_{SelectedYear}_estimate.csv");
            //if (System.IO.File.Exists(estimateFilePath)) { var estimateValues = ReadDataFromCsv(estimateFilePath, this.DaysInMonth); for (int i = 0; i < estimateValues.Count; i++) { combinedData[i].Estimate = estimateValues[i]; } }

            string shiftsForSql = SelectedShifts.Any() ? string.Join(",", SelectedShifts.Select(s => $"'{s}'")) : "'0'";
            var sql = $@"
                        WITH ShiftData AS (
                            SELECT
                                CAST(SDate AS DATE) AS ReportDate,
                                SDate, TotalUnit, NoOfOperator,
                                CASE
                                    WHEN CAST(SDate AS TIME) >= '07:00:00' AND CAST(SDate AS TIME) < '16:00:00' THEN 1
                                    WHEN CAST(SDate AS TIME) >= '16:00:00' AND CAST(SDate AS TIME) < '23:15:00' THEN 2
                                    ELSE 3
                                END as Shift
                            FROM oeesn
                            WHERE YEAR(SDate) = @SelectedYear 
                              AND MONTH(SDate) = @SelectedMonth 
                              AND MachineCode = @MachineLine
                              {dateFilter}
                        ),
                        DailyAggregates AS (
                            SELECT
                                ReportDate,
                                MAX(CASE WHEN Shift = 1 THEN TotalUnit END) as Shift1_EndReading,
                                MAX(CASE WHEN Shift IN (2, 3) THEN TotalUnit END) as OT_EndReading,
                                -- AWAL PERUBAHAN: Ambil NoOfOperator per periode shift
                                MAX(CASE WHEN Shift = 1 THEN NoOfOperator END) as Shift1_Operators,
                                MAX(CASE WHEN Shift IN (2, 3) THEN NoOfOperator END) as OT_Operators,
                                -- AKHIR PERUBAHAN
                                MAX(CASE WHEN Shift IN (2, 3) THEN NoOfOperator END) as OT_OperatorCount,
                                MAX(CASE WHEN Shift IN (2, 3) THEN SDate END) as OT_LastSDate
                            FROM ShiftData
                            GROUP BY ReportDate
                        )
                        SELECT
                            DAY(ReportDate) as Day,
                            CASE WHEN '1' IN ({shiftsForSql}) THEN ISNULL(Shift1_EndReading, 0) ELSE 0 END as LastNormalReading,
                            CASE 
                                WHEN ('2' IN ({shiftsForSql}) OR '3' IN ({shiftsForSql})) THEN ISNULL(OT_EndReading, 0) 
                                ELSE CASE WHEN '1' IN ({shiftsForSql}) THEN ISNULL(Shift1_EndReading, 0) ELSE 0 END
                            END as LastOvertimeReading,
        
                            -- AWAL PERUBAHAN: Logika baru untuk NoOfOperator yang mengikuti filter shift
                            CASE
                                WHEN CASE WHEN '1' IN ({shiftsForSql}) THEN ISNULL(Shift1_Operators, 0) ELSE 0 END >
                                     CASE WHEN '2' IN ({shiftsForSql}) OR '3' IN ({shiftsForSql}) THEN ISNULL(OT_Operators, 0) ELSE 0 END
                                THEN CASE WHEN '1' IN ({shiftsForSql}) THEN ISNULL(Shift1_Operators, 0) ELSE 0 END
                                ELSE CASE WHEN '2' IN ({shiftsForSql}) OR '3' IN ({shiftsForSql}) THEN ISNULL(OT_Operators, 0) ELSE 0 END
                            END as NoOfOperator,
                            -- AKHIR PERUBAHAN

                            ISNULL(OT_OperatorCount, 0) as OtOperatorCount,
                            ISNULL(CAST(OT_LastSDate AS TIME), '00:00:00') as LastOtTime
                        FROM DailyAggregates;";

            try
            {
                using (var connection = new SqlConnection(this.connectionString))
                {
                    using (var command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@SelectedYear", SelectedYear); command.Parameters.AddWithValue("@SelectedMonth", SelectedMonth); command.Parameters.AddWithValue("@MachineLine", MachineLine);
                        if (isCurrentMonthView) { command.Parameters.AddWithValue("@TodayDate", DateTime.Now.Date); }
                        connection.Open();
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                int day = (int)reader["Day"]; var dataForDay = combinedData.FirstOrDefault(d => d.Day == day);
                                if (dataForDay != null) { dataForDay.LastNormalReading = (decimal)reader["LastNormalReading"]; dataForDay.LastOvertimeReading = (decimal)reader["LastOvertimeReading"]; dataForDay.NoOfOperator = Convert.ToInt32(reader["NoOfOperator"]); dataForDay.OtOperatorCount = Convert.ToInt32(reader["OtOperatorCount"]); dataForDay.LastOtTime = (TimeSpan)reader["LastOtTime"]; }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }

            foreach (var data in combinedData)
            {
                ChartLabels.Add(data.Day.ToString()); PlanData.Add(data.Plan); 
                //EstimateData.Add(data.Estimate); 
                NormalData.Add(data.LastNormalReading);
                var overtimeValue = data.LastOvertimeReading - data.LastNormalReading;
                OvertimeData.Add(overtimeValue > 0 ? overtimeValue : 0);
                NoOfDirectWorkers.Add(data.NoOfOperator); OvertimeOperators.Add(data.OtOperatorCount);
                int netOvertimeMinutes = 0; TimeSpan otStartTime = new TimeSpan(16, 0, 0);
                if (data.LastOtTime > otStartTime)
                {
                    var currentDate = new DateTime(SelectedYear, SelectedMonth, data.Day); var dayType = DetermineTypeOfDay(currentDate.DayOfWeek); var allRestTimes = GetRestTime(dayType);
                    int grossMinutes = (int)(data.LastOtTime - otStartTime).TotalMinutes; int restMinutesInOt = GetTotalRestTime(allRestTimes, otStartTime, data.LastOtTime, data.LastOtTime);
                    netOvertimeMinutes = grossMinutes - restMinutesInOt;
                }
                OvertimeMinutes.Add(netOvertimeMinutes > 0 ? netOvertimeMinutes : 0);
                dailyLosses.TryGetValue(data.Day, out int totalSeconds);
                DailyLossTime.Add(totalSeconds / 60);

                // Aturan Bisnis: Jika tidak ada produksi Normal, maka Working Time dianggap 0.
                for (int i = 0; i < this.DaysInMonth; i++)
                {
                    // Cek apakah ada data di list dan apakah nilai NormalData pada hari ke-i adalah 0
                    if (NormalData.Count > i && NormalData[i] == 0)
                    {
                        // Jika ya, atur juga DailyWorkTime pada hari ke-i menjadi 0
                        if (DailyWorkTime.Count > i)
                        {
                            DailyWorkTime[i] = 0;
                        }
                    }
                }
            }
        }

        private readonly List<(TimeSpan Start, TimeSpan End)> RegularDayBreakTimes = new List<(TimeSpan, TimeSpan)>
        {
            (new TimeSpan(9, 30, 0), new TimeSpan(9, 35, 0)),    // Istirahat 5 menit
            (new TimeSpan(12, 0, 0), new TimeSpan(12, 45, 0)),   // Istirahat 45 menit
            (new TimeSpan(14, 30, 0), new TimeSpan(14, 35, 0))   // Istirahat 5 menit
        };

        private readonly List<(TimeSpan Start, TimeSpan End)> FridayBreakTimes = new List<(TimeSpan, TimeSpan)>
        {
            (new TimeSpan(9, 30, 0), new TimeSpan(9, 35, 0)),    // Istirahat 5 menit
            (new TimeSpan(11, 50, 0), new TimeSpan(13, 15, 0)),  // Istirahat Sholat Jumat 85 menit
            (new TimeSpan(14, 30, 0), new TimeSpan(14, 35, 0))   // Istirahat 5 menit
        };
        private bool IsInBreakTime(TimeSpan startTime, TimeSpan endTime, List<(TimeSpan Start, TimeSpan End)> breakTimes) { foreach (var (breakStart, breakEnd) in breakTimes) { if (startTime < breakEnd && endTime > breakStart) { return true; } } return false; }
        private Dictionary<int, int> GetDailyLossTimeTotals()
        {
            var dailyTotals = new Dictionary<int, int>();

            string shiftFilterSql = "";
            // Hanya tambahkan filter jika tidak semua shift dipilih
            if (SelectedShifts.Any() && SelectedShifts.Count < 3)
            {
                var shiftConditions = new List<string>();
                // Logika CASE WHEN ini sekarang SAMA PERSIS dengan halaman Loss Time
                string shiftCaseWhen = @"
            CASE 
                WHEN (DATEPART(HOUR, Time) = 7 AND DATEPART(MINUTE, Time) >= 0) OR 
                     (DATEPART(HOUR, Time) > 7 AND DATEPART(HOUR, Time) < 15) OR 
                     (DATEPART(HOUR, Time) = 15 AND DATEPART(MINUTE, Time) <= 45) THEN '1'
                WHEN (DATEPART(HOUR, Time) = 15 AND DATEPART(MINUTE, Time) > 45) OR 
                     (DATEPART(HOUR, Time) > 15 AND DATEPART(HOUR, Time) < 23) OR 
                     (DATEPART(HOUR, Time) = 23 AND DATEPART(MINUTE, Time) <= 15) THEN '2'
                ELSE '3'
            END";

                foreach (var shift in SelectedShifts)
                {
                    shiftConditions.Add($"{shiftCaseWhen} = '{shift}'");
                }
                shiftFilterSql = $"AND ({string.Join(" OR ", shiftConditions)})";
            }

            string query = $@"
        SELECT 
            CAST(Date AS DATE) as FullDate,
            CAST(Time AS TIME) as StartTime, 
            CAST(EndDateTime AS TIME) as EndTime, 
            LossTime as Duration
        FROM AssemblyLossTime
        WHERE YEAR(Date) = @Year AND MONTH(Date) = @Month AND MachineCode = @Machine
        {shiftFilterSql}"; // Filter yang sudah sinkron disisipkan di sini

            try
            {
                using (var connection = new SqlConnection(this.connectionString))
                {
                    using (var command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Year", SelectedYear);
                        command.Parameters.AddWithValue("@Month", SelectedMonth);
                        command.Parameters.AddWithValue("@Machine", MachineLine);

                        connection.Open();
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var fullDate = (DateTime)reader["FullDate"];
                                var day = fullDate.Day;
                                var startTime = (TimeSpan)reader["StartTime"];
                                var endTime = (TimeSpan)reader["EndTime"];
                                var duration = Convert.ToInt32(reader["Duration"]);

                                var dayType = DetermineTypeOfDay(fullDate.DayOfWeek);
                                var breaksForThisDay = (dayType == "FRIDAY") ? this.FridayBreakTimes : this.RegularDayBreakTimes;

                                if (!IsInBreakTime(startTime, endTime, breaksForThisDay))
                                {
                                    if (!dailyTotals.ContainsKey(day))
                                    {
                                        dailyTotals[day] = 0;
                                    }
                                    dailyTotals[day] += duration;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching loss time: {ex.Message}");
            }
            return dailyTotals;
        }

        //ini kayaknya ga kepakek hehe
        private List<(TimeSpan Start, TimeSpan End)> GetAdditionalBreakTimesForDate(DateTime date)
        {
            var additionalBreaks = new List<(TimeSpan, TimeSpan)>();
            try
            {
                using (var connection = new SqlConnection(this.connectionString))
                {
                    connection.Open();
                    string sql = @"
                SELECT TOP 1 BreakTime1Start, BreakTime1End, BreakTime2Start, BreakTime2End 
                FROM AdditionalBreakTimes 
                WHERE CAST(Date AS DATE) = @Date
                ORDER BY CreatedAt DESC";
                    using (var command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@Date", date.Date);
                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                if (!reader.IsDBNull(0) && !reader.IsDBNull(1))
                                    additionalBreaks.Add((reader.GetTimeSpan(0), reader.GetTimeSpan(1)));
                                if (!reader.IsDBNull(2) && !reader.IsDBNull(3))
                                    additionalBreaks.Add((reader.GetTimeSpan(2), reader.GetTimeSpan(3)));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting additional breaks: {ex.Message}");
            }
            return additionalBreaks;
        }
        public int GetTotalRestTime(List<RestTime> listRestTime, TimeSpan StartTime, TimeSpan EndTime, TimeSpan CurrentTime) { int TotalRestTime = 0; bool isToday = (SelectedYear == DateTime.Now.Year && SelectedMonth == DateTime.Now.Month); TotalRestTime = listRestTime.Sum(rest => { if (isToday && rest.StartTime > CurrentTime) { return 0; } TimeSpan effectiveRestStart = rest.StartTime < StartTime ? StartTime : rest.StartTime; TimeSpan effectiveRestEnd = rest.EndTime > EndTime ? EndTime : rest.EndTime; if (isToday && effectiveRestEnd > CurrentTime) { effectiveRestEnd = CurrentTime; } return effectiveRestStart < effectiveRestEnd ? (int)(effectiveRestEnd - effectiveRestStart).TotalMinutes : 0; }); return TotalRestTime; }
        public string DetermineTypeOfDay(DayOfWeek day) { return day switch { DayOfWeek.Monday or DayOfWeek.Tuesday or DayOfWeek.Wednesday or DayOfWeek.Thursday => "REGULAR", DayOfWeek.Friday => "FRIDAY", DayOfWeek.Saturday or DayOfWeek.Sunday => "WEEKEND", _ => "REGULAR" }; }
        public List<RestTime> GetRestTime(string dayTipe) { List<RestTime> listRestTime = new List<RestTime>(); try { using (SqlConnection connection = new SqlConnection(this.connectionString)) { connection.Open(); string GetRestTime = @"SELECT Duration, StartTime, EndTime FROM RestTime WHERE DayType = @DayTipe"; using (SqlCommand command = new SqlCommand(GetRestTime, connection)) { command.Parameters.AddWithValue("@DayTipe", dayTipe); using (SqlDataReader dataReader = command.ExecuteReader()) { while (dataReader.Read()) { if (!dataReader.IsDBNull(0)) { listRestTime.Add(new RestTime { Duration = dataReader.GetInt32(0), StartTime = dataReader.GetTimeSpan(1), EndTime = dataReader.GetTimeSpan(2) }); } } } } } } catch (Exception ex) { Console.WriteLine("Exception GetRestTime: " + ex.ToString()); } return listRestTime; }
        private List<int> ReadDataFromCsv(string filePath, int DaysInMonth) { var dailyTotals = new List<int>(new int[DaysInMonth]); try { var lines = System.IO.File.ReadAllLines(filePath); foreach (var line in lines) { if (string.IsNullOrWhiteSpace(line)) continue; var valuesForThisLine = line.Split(',').Skip(1).Select(field => int.TryParse(field.Trim(), out int num) ? num : 0).ToList(); for (int i = 0; i < valuesForThisLine.Count && i < DaysInMonth; i++) { dailyTotals[i] += valuesForThisLine[i]; } } } catch { return new List<int>(new int[DaysInMonth]); } return dailyTotals; }
    }
}