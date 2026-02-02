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
        public List<int> OriginalPlanData { get; private set; } = new List<int>();
        public List<int> PlanData { get; private set; } = new List<int>();

        //public List<int> EstimateData { get; private set; } = new List<int>();
        public List<int> NoOfDirectWorkers { get; private set; } = new List<int>();
        public List<int> DailyWorkTime { get; private set; } = new List<int>();
        public List<int> OvertimeOperators { get; private set; } = new List<int>();
        public List<int> OvertimeMinutes { get; private set; } = new List<int>();
        public List<int> DailyLossTime { get; private set; } = new List<int>();
        public List<int> PlanOvertimeData { get; private set; } = new List<int>();

        private class DailyData {
        public int Day { get; set; } 
        public decimal LastNormalReading { get; set; } = 0;
        public decimal LastOvertimeReading { get; set; } = 0;
        public int Plan { get; set; } = 0;
        public int PlanOvertime { get; set; } = 0;
        public int OriginalPlan { get; set; } = 0;
        public int NoOfOperator { get; set; } = 0;
        public int OtOperatorCount { get; set; } = 0;
        public TimeSpan LastOtTime { get; set; } = TimeSpan.Zero;
        }

        public class RestTime { public int Duration { get; set; } public TimeSpan StartTime { get; set; } public TimeSpan EndTime { get; set; } }

        public IndexModel(IWebHostEnvironment webHostEnvironment, IConfiguration configuration)
        {
            _webHostEnvironment = webHostEnvironment;
            _configuration = configuration;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [BindProperty(SupportsGet = true)] public int SelectedMonth { get; set; } = DateTime.Now.Month;
        [BindProperty(SupportsGet = true)] public int SelectedYear { get; set; } = DateTime.Now.Year;
        [BindProperty(SupportsGet = true)] public string MachineLine { get; set; } = "All";
        [BindProperty(SupportsGet = true)] public List<string> SelectedShifts { get; set; } = new List<string>();
        public void OnGet()
        {
            // Normalisasi: jika kosong atau ada "All", set ke "All" saja
            if (!SelectedShifts.Any() || SelectedShifts.Contains("All"))
            {
                SelectedShifts = new List<string> { "All" };
            }
            // Jika user somehow pilih "All" + shift lain, prioritaskan "All"
            else if (SelectedShifts.Count > 1 && SelectedShifts.Contains("All"))
            {
                SelectedShifts = new List<string> { "All" };
            }

            LoadChartData();
        }

        public IActionResult OnPost(string submitButton)
        {
            if (submitButton == "reset")
            {
                return RedirectToPage(new { SelectedYear = DateTime.Now.Year, SelectedMonth = DateTime.Now.Month, MachineLine = "All" });
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
    this.DaysInMonth = DateTime.DaysInMonth(SelectedYear, SelectedMonth);

    var combinedData = Enumerable.Range(1, this.DaysInMonth).Select(day => new DailyData { Day = day }).ToList();
            // Di bagian planShiftFilter
            string planShiftFilter = "";
            if (!SelectedShifts.Contains("All") && SelectedShifts.Any())
            {
                var conditions = SelectedShifts.Select(s => $"pr.shift LIKE '%{s}%'");
                planShiftFilter = $"AND ({string.Join(" OR ", conditions)})";
            }

            List<string> actualShifts = new List<string>();
            if (SelectedShifts.Contains("All"))
            {
                actualShifts.AddRange(new[] { "1", "2", "3" });
            }
            else
            {
                actualShifts.AddRange(SelectedShifts.Where(s => s != "NS"));
            }
            string shiftsForSql = actualShifts.Any() ? string.Join(",", actualShifts.Select(s => $"'{s}'")) : "'0'";
            string planSql = $@"
    SELECT 
        DAY(pp.CurrentDate) as Day,
        SUM(ISNULL(pr.Quantity, 0)) as TotalPlanQuantity,
        SUM(ISNULL(pr.Overtime, 0)) as TotalPlanOvertime
    FROM [PROMOSYS].[dbo].[ProductionPlan] pp
    LEFT JOIN [PROMOSYS].[dbo].[ProductionRecords] pr ON pp.Id = pr.PlanId
    WHERE YEAR(pp.CurrentDate) = @SelectedYear 
      AND MONTH(pp.CurrentDate) = @SelectedMonth
      {planShiftFilter}
      {(MachineLine != "All" ? "AND pr.MachineCode = @MachineLine" : "AND pr.MachineCode IN ('MCH1-01', 'MCH1-02')")}
    GROUP BY DAY(pp.CurrentDate)";

            // ✅ ACTUAL SQL - Conditional append
        string actualSql = $@"
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
          {dateFilter}
          {(MachineLine != "All" ? "AND MachineCode = @MachineLine" : "AND MachineCode IN ('MCH1-01', 'MCH1-02')")}
    ),
    DailyAggregates AS (
        SELECT
            ReportDate,
            MAX(CASE WHEN Shift = 1 THEN TotalUnit END) as Shift1_EndReading,
            MAX(CASE WHEN Shift IN (2, 3) THEN TotalUnit END) as OT_EndReading,
            MAX(CASE WHEN Shift = 1 THEN NoOfOperator END) as Shift1_Operators,
            MAX(CASE WHEN Shift IN (2, 3) THEN NoOfOperator END) as OT_Operators,
            MAX(CASE WHEN Shift IN (2, 3) THEN NoOfOperator END) as OT_OperatorCount,
            MAX(CASE WHEN Shift IN (2, 3) THEN SDate END) as OT_LastSDate
        FROM ShiftData
        GROUP BY ReportDate
    )
    SELECT
        DAY(ReportDate) as Day,
        CASE WHEN '1' IN (" + shiftsForSql + @") THEN ISNULL(Shift1_EndReading, 0) ELSE 0 END as LastNormalReading,
        CASE 
            WHEN ('2' IN (" + shiftsForSql + @") OR '3' IN (" + shiftsForSql + @")) THEN ISNULL(OT_EndReading, 0) 
            ELSE CASE WHEN '1' IN (" + shiftsForSql + @") THEN ISNULL(Shift1_EndReading, 0) ELSE 0 END
        END as LastOvertimeReading,
        CASE
            WHEN CASE WHEN '1' IN (" + shiftsForSql + @") THEN ISNULL(Shift1_Operators, 0) ELSE 0 END >
                 CASE WHEN '2' IN (" + shiftsForSql + @") OR '3' IN (" + shiftsForSql + @") THEN ISNULL(OT_Operators, 0) ELSE 0 END
            THEN CASE WHEN '1' IN (" + shiftsForSql + @") THEN ISNULL(Shift1_Operators, 0) ELSE 0 END
            ELSE CASE WHEN '2' IN (" + shiftsForSql + @") OR '3' IN (" + shiftsForSql + @") THEN ISNULL(OT_Operators, 0) ELSE 0 END
        END as NoOfOperator,
        ISNULL(OT_OperatorCount, 0) as OtOperatorCount,
        ISNULL(CAST(OT_LastSDate AS TIME), '00:00:00') as LastOtTime
    FROM DailyAggregates;";

    try
    {
        using (var connection = new SqlConnection(this.connectionString))
        {
            connection.Open();
            using (var planCmd = new SqlCommand(planSql, connection))
            {
                planCmd.Parameters.AddWithValue("@SelectedYear", SelectedYear);
                planCmd.Parameters.AddWithValue("@SelectedMonth", SelectedMonth);
                
                if (MachineLine != "All")
                {
                    planCmd.Parameters.AddWithValue("@MachineLine", MachineLine);
                }
                
                using (var reader = planCmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        int day = reader.GetInt32(0);
                        var data = combinedData.FirstOrDefault(d => d.Day == day);
                        if (data != null)
                        {
                            data.Plan = Convert.ToInt32(reader["TotalPlanQuantity"]);
                            data.PlanOvertime = Convert.ToInt32(reader["TotalPlanOvertime"]);
                        }
                    }
                }
            }
        }

        using (var connection = new SqlConnection(this.connectionString))
        {
            using (var command = new SqlCommand(actualSql, connection))
            {
                command.Parameters.AddWithValue("@SelectedYear", SelectedYear);
                command.Parameters.AddWithValue("@SelectedMonth", SelectedMonth);
                
                if (MachineLine != "All")
                {
                    command.Parameters.AddWithValue("@MachineLine", MachineLine);
                }
                
                if (isCurrentMonthView)
                {
                    command.Parameters.AddWithValue("@TodayDate", DateTime.Now.Date);
                }
                
                connection.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        int day = (int)reader["Day"];
                        var dataForDay = combinedData.FirstOrDefault(d => d.Day == day);
                        if (dataForDay != null)
                        {
                            dataForDay.LastNormalReading = (decimal)reader["LastNormalReading"];
                            dataForDay.LastOvertimeReading = (decimal)reader["LastOvertimeReading"];
                            dataForDay.NoOfOperator = Convert.ToInt32(reader["NoOfOperator"]);
                            dataForDay.OtOperatorCount = Convert.ToInt32(reader["OtOperatorCount"]);
                            dataForDay.LastOtTime = (TimeSpan)reader["LastOtTime"];
                        }
                    }
                }
            }
        }
    }
            catch (Exception ex) { Console.WriteLine(ex.Message); }

            foreach (var data in combinedData)
            {
                ChartLabels.Add(data.Day.ToString());
                PlanData.Add(data.Plan);
                PlanOvertimeData.Add(data.PlanOvertime);
                OriginalPlanData.Add(data.OriginalPlan);
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

                var currentDay = new DateTime(SelectedYear, SelectedMonth, data.Day);
                string typeOfDay = DetermineTypeOfDay(currentDay.DayOfWeek);

                // Working Time need improvement
                int standardMinutes = (typeOfDay == "FRIDAY") ? 435 : 473;
                if (typeOfDay == "WEEKEND") standardMinutes = 0;

             
                if (!SelectedShifts.Contains("All") && !SelectedShifts.Contains("1"))
                {
                    standardMinutes = 0;
                }

                DailyWorkTime.Add(standardMinutes + (netOvertimeMinutes > 0 ? netOvertimeMinutes : 0));

                for (int i = 0; i < this.DaysInMonth; i++)
                {
                    if (NormalData.Count > i && NormalData[i] == 0)
                    {
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
            (new TimeSpan(9, 30, 0), new TimeSpan(9, 35, 0)),
            (new TimeSpan(12, 0, 0), new TimeSpan(12, 45, 0)),
            (new TimeSpan(14, 30, 0), new TimeSpan(14, 35, 0))
        };

        private readonly List<(TimeSpan Start, TimeSpan End)> FridayBreakTimes = new List<(TimeSpan, TimeSpan)>
        {
            (new TimeSpan(9, 30, 0), new TimeSpan(9, 35, 0)),
            (new TimeSpan(11, 50, 0), new TimeSpan(13, 15, 0)),
            (new TimeSpan(14, 30, 0), new TimeSpan(14, 35, 0))
        };
        private bool IsInBreakTime(TimeSpan startTime, TimeSpan endTime, List<(TimeSpan Start, TimeSpan End)> breakTimes) { foreach (var (breakStart, breakEnd) in breakTimes) { if (startTime < breakEnd && endTime > breakStart) { return true; } } return false; }
        private Dictionary<int, int> GetDailyLossTimeTotals()
        {
            var dailyTotals = new Dictionary<int, int>();

            string shiftFilterSql = "";
            if (SelectedShifts.Any() && SelectedShifts.Count < 3)
            {
                var shiftConditions = new List<string>();
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

            // ✅ DEFINISIKAN FILTER MACHINE
            string lossTimeMachineFilter = (MachineLine == "All")
                ? "AND MachineCode IN ('MCH1-01', 'MCH1-02')"
                : "AND MachineCode = @Machine";

            // ✅ PAKAI VARIABLE {lossTimeMachineFilter} DI QUERY (BUKAN HARDCODE)
            string query = $@"
        SELECT 
            CAST(Date AS DATE) as FullDate,
            CAST(Time AS TIME) as StartTime, 
            CAST(EndDateTime AS TIME) as EndTime, 
            LossTime as Duration
        FROM AssemblyLossTime
        WHERE YEAR(Date) = @Year 
          AND MONTH(Date) = @Month 
          {lossTimeMachineFilter}
          {shiftFilterSql}";

            try
            {
                using (var connection = new SqlConnection(this.connectionString))
                {
                    using (var command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Year", SelectedYear);
                        command.Parameters.AddWithValue("@Month", SelectedMonth);

                        // ✅ HANYA TAMBAH PARAMETER @Machine JIKA BUKAN "All"
                        if (MachineLine != "All")
                        {
                            command.Parameters.AddWithValue("@Machine", MachineLine);
                        }

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
    }
}