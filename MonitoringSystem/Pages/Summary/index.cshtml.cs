using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using System.Data;

namespace MonitoringSystem.Pages.Summary
{
    // Kelas Data Mentah
    public class RawOeeData { public DateTime SDate { get; set; } public string MachineCode { get; set; } public string ProductName { get; set; } public int SUT { get; set; } public int NoOfOperator { get; set; } public decimal TargetUnit { get; set; } }

    public class DailyPlanQty
    {
        public decimal Quantity { get; set; }
        public string MachineCode { get; set; }
    }

    public class RawNgData { public DateTime SDate { get; set; } public string MachineCode { get; set; } }
    public class RawRestTimeData { public string DayType { get; set; } public TimeSpan StartTime { get; set; } public TimeSpan EndTime { get; set; } }
    public class RawLossTimeData { public DateTime Date { get; set; } public TimeSpan StartTime { get; set; } public TimeSpan EndTime { get; set; } public string MachineCode { get; set; } }

    public class SummaryModel : PageModel
    {
        public string connectionString = "Server=10.83.33.103;User Id=sa;Password=sa;Database=PROMOSYS;Trusted_Connection=False;TrustServerCertificate=True;Encrypt=False";
        public string errorMessage = "";

        // Properti publik untuk menampung semua data mentah
        public List<RawOeeData> RawOeeData { get; private set; }
        public List<RawNgData> RawNgData { get; private set; }
        public List<DailyPlanQty> DailyPlanData { get; private set; }
        public List<PlanQty> PlanData { get; private set; }
        public List<RawRestTimeData> RawRestTimeData { get; private set; }
        public List<RawLossTimeData> RawLossTimeData { get; private set; } // PENAMBAHAN BARU

        public SummaryModel()
        {
            // Inisialisasi list
            RawOeeData = new List<RawOeeData>();
            RawNgData = new List<RawNgData>();
            PlanData = new List<PlanQty>();
            RawRestTimeData = new List<RawRestTimeData>();
            RawLossTimeData = new List<RawLossTimeData>(); // PENAMBAHAN BARU
            DailyPlanData = new List<DailyPlanQty>();
        }

        [BindProperty] public DateTime StartSelectedDate { get; set; } = DateTime.Today;
        [BindProperty] public DateTime EndSelectedDate { get; set; } = DateTime.Today;

        private void InitializePage() { LoadRawDataForDateRange(); }
        public void OnGet() { StartSelectedDate = DateTime.Today; EndSelectedDate = DateTime.Today; InitializePage(); }
        public void OnPost() { if (EndSelectedDate < StartSelectedDate) { EndSelectedDate = StartSelectedDate; } InitializePage(); }

        // GANTI SELURUH ISI METODE DENGAN KODE DI BAWAH INI
        private void LoadRawDataForDateRange()
        {
            var globalStartTime = StartSelectedDate.Date;
            var globalEndTime = EndSelectedDate.Date.AddDays(1).AddHours(7);

            // Kosongkan list agar tidak double
            RawOeeData.Clear();
            RawNgData.Clear();
            PlanData.Clear();
            RawRestTimeData.Clear();
            RawLossTimeData.Clear();

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // 🔹 Query 1: Ambil data OEE + Target langsung dari OEESN
                    string oeeSql = @"
                SELECT 
                    OEESN.SDate,
                    OEESN.MachineCode,
                    md.ProductName,
                    md.SUT,
                    OEESN.NoOfOperator,
                    ISNULL(OEESN.TargetUnit, 0) AS TargetUnit
                FROM OEESN
                JOIN MasterData md ON OEESN.Product_Id = md.Product_Id
                WHERE OEESN.SDate >= @StartTime AND OEESN.SDate < @EndTime";

                    using (SqlCommand cmd = new SqlCommand(oeeSql, connection))
                    {
                        cmd.Parameters.AddWithValue("@StartTime", globalStartTime);
                        cmd.Parameters.AddWithValue("@EndTime", globalEndTime);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                RawOeeData.Add(new RawOeeData
                                {
                                    SDate = reader.GetDateTime(0),
                                    MachineCode = reader.GetString(1),
                                    ProductName = reader.GetString(2),
                                    SUT = reader.GetInt32(3),
                                    NoOfOperator = reader.IsDBNull(4) ? 0 : reader.GetInt32(4),
                                    TargetUnit = reader.IsDBNull(5) ? 0 : Convert.ToDecimal(reader.GetValue(5))
                                });
                            }
                        }
                    }

                    // 🔹 Query 2: NG Data
                    string ngSql = @"SELECT SDate, MachineCode FROM NG_RPTS WHERE SDate >= @StartTime AND SDate < @EndTime";
                    using (SqlCommand cmd = new SqlCommand(ngSql, connection))
                    {
                        cmd.Parameters.AddWithValue("@StartTime", globalStartTime);
                        cmd.Parameters.AddWithValue("@EndTime", globalEndTime);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                                RawNgData.Add(new RawNgData
                                {
                                    SDate = reader.GetDateTime(0),
                                    MachineCode = reader.GetString(1)
                                });
                        }
                    }

                    // 🔹 Query 3: Waktu Istirahat
                    string restTimeSql = @"SELECT DayType, StartTime, EndTime FROM RestTime";
                    using (SqlCommand cmd = new SqlCommand(restTimeSql, connection))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                RawRestTimeData.Add(new RawRestTimeData
                                {
                                    DayType = reader.GetString(0),
                                    StartTime = reader.GetTimeSpan(1),
                                    EndTime = reader.GetTimeSpan(2)
                                });
                            }
                        }
                    }

                    // 🔹 Query 4: Loss Time
                    string lossTimeSql = @"SELECT Date, Time, EndDateTime, MachineCode 
                                   FROM AssemblyLossTime 
                                   WHERE Date >= @StartTime AND Date < @EndTime";
                    using (SqlCommand cmd = new SqlCommand(lossTimeSql, connection))
                    {
                        cmd.Parameters.AddWithValue("@StartTime", globalStartTime);
                        cmd.Parameters.AddWithValue("@EndTime", globalEndTime);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                RawLossTimeData.Add(new RawLossTimeData
                                {
                                    Date = reader.GetDateTime(0),
                                    StartTime = reader.GetTimeSpan(1),
                                    EndTime = reader.GetTimeSpan(2),
                                    MachineCode = reader.GetString(3)
                                });
                            }
                        }
                    }

                    // 🔹 Query 5: Daily Plan dari tabel ProductionRecords + ProductionPlan
                    string dailyPlanSql = @"
                        SELECT 
                            CAST(SUM(pr.Quantity) AS DECIMAL(10,0)) AS Quantity,
                            pr.MachineCode
                        FROM ProductionRecords pr
                        JOIN ProductionPlan pp ON pr.PlanId = pp.Id
                        WHERE CAST(pp.CurrentDate AS DATE) >= @StartTime
                          AND CAST(pp.CurrentDate AS DATE) < @EndTime
                        GROUP BY pr.MachineCode;
                    ";

                    using (SqlCommand cmd = new SqlCommand(dailyPlanSql, connection))
                    {
                        cmd.Parameters.AddWithValue("@StartTime", globalStartTime);
                        cmd.Parameters.AddWithValue("@EndTime", globalEndTime);

                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                DailyPlanData.Add(new DailyPlanQty
                                {
                                    MachineCode = reader.GetString(1),
                                    Quantity = reader.IsDBNull(0) ? 0 : Convert.ToDecimal(reader.GetValue(0))
                                });
                            }
                        }
                    }


                    // ✅ Debug Log
                    Console.WriteLine($"[DEBUG] Loaded {RawOeeData.Count} OEE rows");
                    Console.WriteLine($"[DEBUG] Total TargetUnit: {RawOeeData.Sum(x => x.TargetUnit)}");
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Gagal memuat data mentah dari database: " + ex.Message;
                Console.WriteLine("❌ ERROR di LoadRawDataForDateRange(): " + ex.ToString());
            }
        }


        // --- Metode Perhitungan ---

        public List<RestTime> GetRestTime(string dayTipe) => RawRestTimeData.Where(r => r.DayType == dayTipe).Select(r => new RestTime { StartTime = r.StartTime, EndTime = r.EndTime }).ToList();
        public List<ActualQty> GetActualData(TimeSpan startTime, TimeSpan endTime) => RawOeeData.Where(d => d.SDate.TimeOfDay >= startTime && d.SDate.TimeOfDay < endTime).GroupBy(d => d.MachineCode).Select(g => new ActualQty { MachineCode = g.Key, Quantity = g.Count() }).ToList();
        public List<PlanQty> GetPlanFromTargetUnit(DateTime startTime, DateTime endTime)
        {
            return RawOeeData
                .Where(d => d.SDate >= startTime && d.SDate < endTime)
                .GroupBy(d => d.MachineCode)
                .Select(g =>
                {
                    var latest = g.OrderByDescending(d => d.SDate).FirstOrDefault();
                    return new PlanQty
                    {
                        MachineCode = g.Key,
                        Quantity = latest?.TargetUnit ?? 0 // Ambil nilai TargetUnit terbaru
                    };
                })
                .ToList();
        }


        public List<SUTModel> GetSUTModel(TimeSpan startTime, TimeSpan endTime) => RawOeeData.Where(d => d.SDate.TimeOfDay >= startTime && d.SDate.TimeOfDay < endTime).OrderByDescending(d => d.SDate).GroupBy(d => d.MachineCode).Select(g => new SUTModel { MachineCode = g.Key, SUT = g.First().SUT }).ToList();
        public List<TotalDefect> GetTotalDefect(TimeSpan startTime, TimeSpan endTime) => RawNgData.Where(d => d.SDate.TimeOfDay >= startTime && d.SDate.TimeOfDay < endTime).GroupBy(d => d.MachineCode).Select(g => new TotalDefect { MachineCode = g.Key, DefectQty = g.Count() }).ToList();
        public int GetTotalChangeModel(string machineCode, TimeSpan startTime, TimeSpan endTime) => RawOeeData.Where(d => d.MachineCode == machineCode && d.SDate.TimeOfDay >= startTime && d.SDate.TimeOfDay < endTime).Select(d => d.ProductName).Distinct().Count();

        // PENAMBAHAN BARU: Metode untuk menghitung Loss Time dari data mentah
        public int CalculateTotalLossTime(string machineCode, TimeSpan shiftStart, TimeSpan shiftEnd, List<RestTime> breaks)
        {
            var totalLossSeconds = RawLossTimeData
                .Where(lt =>
                    lt.MachineCode == machineCode &&
                    lt.StartTime >= shiftStart &&
                    lt.EndTime < shiftEnd &&
                    !breaks.Any(b => lt.StartTime < b.EndTime && lt.EndTime > b.StartTime) // Cek tidak tumpang tindih dengan istirahat
                )
                .Sum(lt => (lt.EndTime - lt.StartTime).TotalSeconds);

            return (int)totalLossSeconds / 60; // Kembalikan dalam menit
        }

        public List<DailyPlanQty> GetDailyPlanData(DateTime startTime, DateTime endTime)
        {
            return DailyPlanData
                .Where(d => d != null)
                .GroupBy(d => d.MachineCode)
                .Select(g => new DailyPlanQty
                {
                    MachineCode = g.Key,
                    Quantity = g.Sum(x => x.Quantity)
                })
                .ToList();
        }

        // --- Metode Helper ---
        public string DetermineTypeOfDay(DayOfWeek day) => day switch { DayOfWeek.Friday => "FRIDAY", DayOfWeek.Saturday or DayOfWeek.Sunday => "WEEKEND", _ => "REGULAR" };
        public int GetTotalRestTime(List<RestTime> listRestTime, TimeSpan StartTime, TimeSpan EndTime, TimeSpan CurrentTime)
        {
            int totalRestMinutes = 0;
            TimeSpan effectiveCurrentTime = (StartSelectedDate.Date == DateTime.Today) ? CurrentTime : EndTime;
            foreach (var rest in listRestTime)
            {
                var overlapStart = new TimeSpan(Math.Max(StartTime.Ticks, rest.StartTime.Ticks));
                var overlapEnd = new TimeSpan(Math.Min(effectiveCurrentTime.Ticks, rest.EndTime.Ticks));
                if (overlapEnd > overlapStart) totalRestMinutes += (int)(overlapEnd - overlapStart).TotalMinutes;
            }
            return totalRestMinutes;
        }

        // Metode ini akan menangani rentang waktu yang melewati tengah malam

        public List<ActualQty> GetActualData(DateTime startTime, DateTime endTime) =>
            RawOeeData
                .Where(d => d.SDate >= startTime && d.SDate < endTime)
                .GroupBy(d => d.MachineCode)
                .Select(g => new ActualQty { MachineCode = g.Key, Quantity = g.Count() })
                .ToList();

        public List<TotalDefect> GetTotalDefect(DateTime startTime, DateTime endTime) =>
            RawNgData
                .Where(d => d.SDate >= startTime && d.SDate < endTime)
                .GroupBy(d => d.MachineCode)
                .Select(g => new TotalDefect { MachineCode = g.Key, DefectQty = g.Count() })
                .ToList();

        public int CalculateTotalLossTime(string machineCode, DateTime shiftStart, DateTime shiftEnd, List<RestTime> breaks)
        {
            var totalLossSeconds = RawLossTimeData
                .Where(lt =>
                {
                    // Buat DateTime lengkap untuk setiap loss event
                    var lossStart = lt.Date.Date + lt.StartTime;
                    var lossEnd = lt.Date.Date + lt.EndTime;
                    // Jika loss melewati tengah malam, tambahkan satu hari ke waktu akhir
                    if (lossEnd < lossStart) lossEnd = lossEnd.AddDays(1);

                    return lt.MachineCode == machineCode &&
                           lossStart >= shiftStart &&
                           lossEnd < shiftEnd &&
                           !breaks.Any(b => {
                               var breakStart = shiftStart.Date + b.StartTime;
                               var breakEnd = shiftStart.Date + b.EndTime;
                               if (b.StartTime > b.EndTime) breakEnd = breakEnd.AddDays(1); // Handle istirahat lintas hari jika ada
                               return lossStart < breakEnd && lossEnd > breakStart;
                           });
                })
                .Sum(lt => {
                    var lossEnd = lt.Date.Date + lt.EndTime;
                    if (lossEnd < (lt.Date.Date + lt.StartTime)) lossEnd = lossEnd.AddDays(1);
                    return (lossEnd - (lt.Date.Date + lt.StartTime)).TotalSeconds;
                });

            return (int)totalLossSeconds / 60; // Kembalikan dalam menit
        }


        public int GetTotalChangeModel(string machineCode, DateTime startTime, DateTime endTime) =>
            RawOeeData
                .Where(d => d.MachineCode == machineCode && d.SDate >= startTime && d.SDate < endTime)
                .Select(d => d.ProductName)
                .Distinct()
                .Count();

        public List<StartEndModel> GetStartEndModel(string machineCode, DateTime startTime, DateTime endTime)
        {
            var results = new List<StartEndModel>();
            var filteredData = RawOeeData
                .Where(d => d.MachineCode == machineCode && d.SDate >= startTime && d.SDate < endTime)
                .OrderBy(d => d.SDate)
                .ToList();

            if (!filteredData.Any()) return results;

            var currentBlock = new StartEndModel { Model = filteredData.First().ProductName, StartTime = filteredData.First().SDate, EndTime = filteredData.First().SDate };
            for (int i = 1; i < filteredData.Count; i++)
            {
                if (filteredData[i].ProductName == currentBlock.Model)
                {
                    currentBlock.EndTime = filteredData[i].SDate;
                }
                else
                {
                    results.Add(currentBlock);
                    currentBlock = new StartEndModel { Model = filteredData[i].ProductName, StartTime = filteredData[i].SDate, EndTime = filteredData[i].SDate };
                }
            }
            results.Add(currentBlock);
            return results;
        }

        public int CalculateTotalWorkingTimeForRange(TimeSpan shiftStart, TimeSpan shiftEnd)
        {
            double totalMinutes = 0;

            // Loop untuk setiap hari dalam rentang tanggal yang dipilih
            for (var day = StartSelectedDate.Date; day <= EndSelectedDate.Date; day = day.AddDays(1))
            {
                var dayType = DetermineTypeOfDay(day.DayOfWeek);

                // Lewati hari libur (misal: Weekend)
                if (dayType == "WEEKEND")
                {
                    continue;
                }

                // Ambil jadwal istirahat untuk tipe hari ini
                var breaksForDay = RawRestTimeData
                    .Where(r => r.DayType == dayType)
                    .ToList();

                double shiftDurationMinutes = (shiftEnd - shiftStart).TotalMinutes;
                double restMinutesForDay = 0;

                foreach (var rest in breaksForDay)
                {
                    // Hitung durasi istirahat yang berada dalam jam kerja shift
                    var overlapStart = new TimeSpan(Math.Max(shiftStart.Ticks, rest.StartTime.Ticks));
                    var overlapEnd = new TimeSpan(Math.Min(shiftEnd.Ticks, rest.EndTime.Ticks));
                    if (overlapEnd > overlapStart)
                    {
                        restMinutesForDay += (overlapEnd - overlapStart).TotalMinutes;
                    }
                }

                // Tambahkan waktu kerja bersih hari ini ke total
                totalMinutes += (shiftDurationMinutes - restMinutesForDay);
            }

            // Logika untuk hari ini (jika rentang mencakup hari ini, maka waktu kerja dihitung sampai jam saat ini)
            if (EndSelectedDate.Date == DateTime.Today && StartSelectedDate.Date == DateTime.Today)
            {
                var workingDurationToday = (DateTime.Now.TimeOfDay > shiftStart ? (DateTime.Now.TimeOfDay < shiftEnd ? DateTime.Now.TimeOfDay - shiftStart : shiftEnd - shiftStart) : TimeSpan.Zero);
                var totalRestTimeToday = GetTotalRestTime(GetRestTime(DetermineTypeOfDay(DateTime.Today.DayOfWeek)), shiftStart, shiftEnd, DateTime.Now.TimeOfDay);
                return Math.Max(0, (int)workingDurationToday.TotalMinutes - totalRestTimeToday);
            }


            return (int)totalMinutes;
        }

        public List<StartEndModel> GetStartEndModel(string machineCode, TimeSpan startTime, TimeSpan endTime)
        {
            var results = new List<StartEndModel>();
            var filteredData = RawOeeData.Where(d => d.MachineCode == machineCode && d.SDate.TimeOfDay >= startTime && d.SDate.TimeOfDay < endTime).OrderBy(d => d.SDate).ToList();
            if (!filteredData.Any()) return results;
            var currentBlock = new StartEndModel { Model = filteredData.First().ProductName, StartTime = filteredData.First().SDate, EndTime = filteredData.First().SDate };
            for (int i = 1; i < filteredData.Count; i++)
            {
                if (filteredData[i].ProductName == currentBlock.Model) { currentBlock.EndTime = filteredData[i].SDate; }
                else { results.Add(currentBlock); currentBlock = new StartEndModel { Model = filteredData[i].ProductName, StartTime = filteredData[i].SDate, EndTime = filteredData[i].SDate }; }
            }
            results.Add(currentBlock);
            return results;
        }
        public int GetAverageWorkersForPeriod(string machineCode, DateTime startTime, DateTime endTime)
        {
            var workersPerDay = RawOeeData
                .Where(d => d.MachineCode == machineCode && d.SDate >= startTime && d.SDate < endTime)
                .GroupBy(d => d.SDate.Date)
                .Select(g => g.Max(d => d.NoOfOperator)) 
                .ToList();

            if (workersPerDay.Any())
            {
                return (int)Math.Round(workersPerDay.Average());
            }
            return 0;
        }

        // Tambahkan metode baru ini di dalam kelas SummaryModel
        public double CalculateSingleDayElapsedWorkingTime(DateTime day, TimeSpan shiftStart, TimeSpan shiftEnd, List<RestTime> shiftBreaks)
        {
            var now = DateTime.Now;

            // Jika hari yang diproses adalah di masa depan, waktu kerja adalah 0
            if (day.Date > now.Date)
            {
                return 0;
            }

            // Jika hari yang diproses adalah di masa lalu, hitung waktu kerja bersih satu shift penuh
            // BLOK BARU YANG SUDAH DIPERBAIKI
            if (day.Date < now.Date)
            {
                double totalShiftMinutes = (shiftEnd - shiftStart).TotalMinutes;
                double actualBreakMinutesInShift = 0;

                // Loop melalui setiap jadwal istirahat yang ada untuk hari itu
                foreach (var rest in shiftBreaks)
                {
                    // Tentukan waktu mulai dan akhir dari tumpang tindih (overlap) antara jam kerja dan jam istirahat
                    var overlapStart = new TimeSpan(Math.Max(shiftStart.Ticks, rest.StartTime.Ticks));
                    var overlapEnd = new TimeSpan(Math.Min(shiftEnd.Ticks, rest.EndTime.Ticks));

                    // Jika ada tumpang tindih yang valid (waktu akhir > waktu mulai)
                    if (overlapEnd > overlapStart)
                    {
                        // Tambahkan durasi tumpang tindih tersebut ke total istirahat yang valid
                        actualBreakMinutesInShift += (overlapEnd - overlapStart).TotalMinutes;
                    }
                }
                return totalShiftMinutes - actualBreakMinutesInShift;
            }

            // Jika hari yang diproses adalah HARI INI, hitung secara dinamis
            // Waktu efektif berakhir adalah waktu saat ini, kecuali jika shift sudah selesai
            TimeSpan effectiveEndTime = (now.TimeOfDay > shiftEnd) ? shiftEnd : now.TimeOfDay;

            // Jika shift belum dimulai, waktu kerja adalah 0
            if (effectiveEndTime <= shiftStart)
            {
                return 0;
            }

            // Hitung durasi kotor dari awal shift hingga waktu efektif berakhir
            double grossDurationMinutes = (effectiveEndTime - shiftStart).TotalMinutes;

            // Hitung total istirahat yang sudah terjadi dalam durasi tersebut
            int restMinutesElapsed = GetTotalRestTime(shiftBreaks, shiftStart, shiftEnd, now.TimeOfDay);

            return Math.Max(0, grossDurationMinutes - restMinutesElapsed);
        }

        // Tambahkan metode baru ini di dalam kelas SummaryModel

        public double CalculateOvernightShiftElapsedWorkingTime(DateTime startDay, TimeSpan shiftStart, TimeSpan shiftEnd, List<RestTime> shiftBreaks)
        {
            var now = DateTime.Now;

            // Tentukan waktu mulai dan selesai shift secara penuh (start di hari ini, end di hari besok)
            var shiftStartDateTime = startDay.Date + shiftStart;
            var shiftEndDateTime = startDay.Date.AddDays(1) + shiftEnd;

            // 1. Jika shift di masa depan (waktu sekarang masih sebelum shift dimulai)
            if (now < shiftStartDateTime)
            {
                return 0;
            }

            // 2. Jika shift sudah sepenuhnya berlalu (waktu sekarang sudah melewati akhir shift)
            if (now >= shiftEndDateTime)
            {
                double totalShiftMinutes = (shiftEndDateTime - shiftStartDateTime).TotalMinutes;
                // Asumsi jam istirahat untuk shift malam tidak ada atau sudah termasuk dalam kalkulasi total.
                // Jika ada, perlu penanganan khusus. Untuk saat ini, kita anggap total durasi.
                double totalBreakMinutes = 0; // Sesuaikan jika ada jadwal istirahat tetap untuk shift 3
                return totalShiftMinutes - totalBreakMinutes;
            }

            // 3. Jika shift sedang berjalan (waktu sekarang ada di antara mulai dan selesai shift)
            // Durasi kotor dari awal shift hingga saat ini
            double grossDurationMinutes = (now - shiftStartDateTime).TotalMinutes;

            // Hitung istirahat yang sudah terjadi (jika ada)
            // NOTE: Metode GetTotalRestTime yang ada mungkin perlu penyesuaian untuk shift malam.
            // Untuk saat ini kita asumsikan 0 untuk penyederhanaan.
            int restMinutesElapsed = 0;

            return Math.Max(0, grossDurationMinutes - restMinutesElapsed);
        }

        public class PlanQty { public decimal Quantity { get; set; } public string MachineCode { get; set; } }
        public class ActualQty { public int Quantity { get; set; } public string MachineCode { get; set; } }
        public class SUTModel { public int SUT { get; set; } public string MachineCode { get; set; } }
        public class RestTime { public TimeSpan StartTime { get; set; } public TimeSpan EndTime { get; set; } }
        public class TotalDefect { public string MachineCode { get; set; } public int DefectQty { get; set; } }
        public class StartEndModel { public string Model { get; set; } public DateTime StartTime { get; set; } public DateTime EndTime { get; set; } }
    }
}