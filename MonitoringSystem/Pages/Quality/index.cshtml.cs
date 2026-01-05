using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using System.Reflection.PortableExecutable;
using System.Collections.Generic; // Tambahkan ini
using System; // Tambahkan ini

namespace MonitoringSystem.Pages.Quality
{
    public class QualityModel : PageModel
    {
        public string connectionString = "Server=10.83.33.103;trusted_connection=false;Database=PROMOSYS;User Id=sa;Password=sa;Persist Security Info=False;Encrypt=False;MultipleActiveResultSets=True";

        public int TotalPlan { get; set; }
        public int DefectQuantity { get; set; }
        public double DefectRatio { get; set; }
        public string errorMessage = "";

        public List<DailyDefect> TopDailyDefects { get; set; }

        public List<DailyDefect> DefectProblems { get; set; }
        public List<DefectByModel> DefectsByModel { get; set; }

        //public List<MonthlyDefectData> MonthlyDefects { get; set; }
        public List<YearlyDefectData> YearlyDefects { get; set; }

        [BindProperty(SupportsGet = true)]
        public string MachineCode { get; set; }

        [BindProperty(SupportsGet = true)]
        public string StartDate { get; set; }

        [BindProperty(SupportsGet = true)]
        public string EndDate { get; set; }
        public double CurrentTargetRatio { get; set; }
        [BindProperty]
        public double NewTargetRatio { get; set; }

        [BindProperty(SupportsGet = true)]
        public string Station { get; set; }
        public QualityModel()
        {
            TopDailyDefects = new List<DailyDefect>();
            DefectProblems = new List<DailyDefect>();
            DefectsByModel = new List<DefectByModel>();
            //MonthlyDefects = new List<MonthlyDefectData>();
            YearlyDefects = new List<YearlyDefectData>();
        }

        public async Task<IActionResult> OnPostUpdateTargetRatioAsync()
        {
            ModelState.Remove("MachineCode");
            ModelState.Remove("StartDate");
            ModelState.Remove("EndDate");
            ModelState.Remove("Station");

            if (ModelState.IsValid)
            {
                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        await connection.OpenAsync();
                        string insertQuery = "INSERT INTO TargetRatioDefect.dbo.TargetRatio (Ratio) VALUES (@NewRatio)";
                        using (SqlCommand command = new SqlCommand(insertQuery, connection))
                        {
                            command.Parameters.AddWithValue("@NewRatio", NewTargetRatio);
                            await command.ExecuteNonQueryAsync();
                        }
                    }
                }
                catch (Exception ex)
                {
                    errorMessage = "Database error while updating target ratio: " + ex.Message;
                    Console.WriteLine(errorMessage);
                    return Page();
                }
            }
            return RedirectToPage(new { MachineCode, StartDate, EndDate, Station });
        }

        public void OnGet()
        {
            if (string.IsNullOrEmpty(MachineCode))
            {
                MachineCode = "MCH1-01";
            }
            if (string.IsNullOrEmpty(StartDate))
            {
                StartDate = DateTime.Now.ToString("yyyy-MM-dd");
                EndDate = DateTime.Now.ToString("yyyy-MM-dd");
            }
            LoadData();
        }

        public void OnPost()
        {
            LoadData();
        }

        private void LoadData()
        {
            TotalPlan = 0;
            DefectQuantity = 0;
            DefectRatio = 100;
            TopDailyDefects.Clear();
            DefectProblems.Clear();
            DefectsByModel.Clear();
            //MonthlyDefects.Clear();
            YearlyDefects.Clear();

            DateTime startDateParsed, endDateParsed;
            if (!DateTime.TryParse(StartDate, out startDateParsed))
            {
                startDateParsed = DateTime.Now.Date;
            }
            if (!DateTime.TryParse(EndDate, out endDateParsed))
            {
                endDateParsed = DateTime.Now.Date;
            }

            string stationFilterClause = "";
            if (!string.IsNullOrEmpty(Station))
            {
                stationFilterClause = " AND Station = @Station";
            }

            Action<SqlCommand> addStationParameter = (cmd) => {
                if (!string.IsNullOrEmpty(Station))
                {
                    cmd.Parameters.AddWithValue("@Station", Station);
                }
            };

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // === GET TARGET RATIO (Dipindahkan ke atas agar selalu dieksekusi pertama) ===
                    string getTargetRatio = "SELECT TOP 1 Ratio FROM TargetRatioDefect.dbo.TargetRatio ORDER BY ID DESC";
                    using (SqlCommand command = new SqlCommand(getTargetRatio, connection))
                    {
                        var result = command.ExecuteScalar();
                        if (result != DBNull.Value && result != null)
                        {
                            CurrentTargetRatio = Convert.ToDouble(result);
                        }
                    }

                    string getTotalProduction = @"
                    SELECT
                         COUNT(TotalUnit)
                    FROM
                        OEESN
                    WHERE
                        MachineCode = @MachineCode
                       AND CAST(Date AS DATE) BETWEEN @StartDate AND @EndDate;";

                    using (SqlCommand command = new SqlCommand(getTotalProduction, connection))
                    {
                        command.Parameters.AddWithValue("@MachineCode", MachineCode);
                        command.Parameters.AddWithValue("@StartDate", startDateParsed);
                        command.Parameters.AddWithValue("@EndDate", endDateParsed);
                        var result = command.ExecuteScalar();
                        if (result != DBNull.Value && result != null)
                        {
                            TotalPlan = Convert.ToInt32(result);
                        }
                    }

                    string getTotalDefect = $@"
                    SELECT
                        COUNT(*)
                    FROM
                        NG_RPTS
                    WHERE
                        MachineCode = @MachineCode
                        AND CAST(SDate AS DATE) BETWEEN @StartDate AND @EndDate
                        {stationFilterClause};";

                    using (SqlCommand command = new SqlCommand(getTotalDefect, connection))
                    {
                        command.Parameters.AddWithValue("@MachineCode", MachineCode);
                        command.Parameters.AddWithValue("@StartDate", startDateParsed);
                        command.Parameters.AddWithValue("@EndDate", endDateParsed);
                        addStationParameter(command);
                        var result = command.ExecuteScalar();
                        if (result != DBNull.Value && result != null)
                        {
                            DefectQuantity = Convert.ToInt32(result);
                        }
                    }

                    // =================== KODE YANG DIPERBAIKI ===================
                    string getTopDailyDefect = $@"
                    SELECT TOP 5
                        Cause,
                        COUNT(*) AS DefectCount
                    FROM
                        NG_RPTS
                    WHERE
                        CAST(SDate AS DATE) BETWEEN @StartDate AND @EndDate 
                        AND MachineCode = @MachineCode
                        {stationFilterClause}
                    GROUP BY
                        Cause
                    ORDER BY
                        DefectCount DESC;";
                    // ==========================================================

                    using (SqlCommand command = new SqlCommand(getTopDailyDefect, connection))
                    {
                        command.Parameters.AddWithValue("@MachineCode", MachineCode);
                        command.Parameters.AddWithValue("@StartDate", startDateParsed);
                        command.Parameters.AddWithValue("@EndDate", endDateParsed);
                        addStationParameter(command);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                TopDailyDefects.Add(new DailyDefect
                                {
                                    Cause = reader.GetString(0),
                                    Quantity = reader.GetInt32(1)
                                });
                            }
                        }
                    }

                    string getDefectProblem = $@"
                    SELECT
                        Cause,
                        COUNT(*) AS DefectCount
                    FROM
                        NG_RPTS
                    WHERE
                        CAST(SDate AS DATE) BETWEEN @StartDate AND @EndDate
                        AND MachineCode = @MachineCode
                        {stationFilterClause}
                    GROUP BY
                        Cause
                    ORDER BY
                        DefectCount DESC;";
                    using (SqlCommand command = new SqlCommand(getDefectProblem, connection))
                    {
                        command.Parameters.AddWithValue("@MachineCode", MachineCode);
                        command.Parameters.AddWithValue("@StartDate", startDateParsed);
                        command.Parameters.AddWithValue("@EndDate", endDateParsed);
                        addStationParameter(command);
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                DefectProblems.Add(new DailyDefect
                                {
                                    Cause = reader.GetString(0),
                                    Quantity = reader.GetInt32(1)
                                });
                            }
                        }
                    }

                    //string getDefectsByModel = $@"
                    //SELECT
                    //    md.ProductName,
                    //    COUNT(*) AS DefectCount
                    //FROM
                    //    NG_RPTS ng
                    //JOIN
                    //    MasterData md ON 
                    //    LTRIM(RTRIM(CAST(ng.Product_Id AS VARCHAR(255)))) = LTRIM(RTRIM(CAST(md.Product_Id AS VARCHAR(255))))
                    //WHERE
                    //    ng.MachineCode = @MachineCode
                    //    AND CAST(ng.SDate AS DATE) BETWEEN @StartDate AND @EndDate
                    //    {stationFilterClause.Replace("Station", "ng.Station")}
                    //GROUP BY
                    //    md.ProductName
                    //ORDER BY
                    //    DefectCount DESC;";

                    //using (SqlCommand command = new SqlCommand(getDefectsByModel, connection))
                    //{
                    //    command.Parameters.AddWithValue("@MachineCode", MachineCode);
                    //    command.Parameters.AddWithValue("@StartDate", startDateParsed);
                    //    command.Parameters.AddWithValue("@EndDate", endDateParsed);
                    //    addStationParameter(command);
                    //    using (SqlDataReader reader = command.ExecuteReader())
                    //    {
                    //        while (reader.Read())
                    //        {
                    //            DefectsByModel.Add(new DefectByModel
                    //            {
                    //                ProductName = reader.IsDBNull(0) ? "Nama Produk Kosong" : reader.GetString(0),
                    //                Quantity = reader.IsDBNull(1) ? 0 : reader.GetInt32(1)
                    //            });
                    //        }
                    //    }
                    //}

                    string getDefectsByModel = $@"
                    SELECT
                        Station,
                        COUNT(*) AS DefectCount
                    FROM
                        NG_RPTS
                    WHERE
                        MachineCode = @MachineCode
                        AND CAST(SDate AS DATE) BETWEEN @StartDate AND @EndDate
                        {stationFilterClause}
                    GROUP BY
                        Station
                    ORDER BY
                        DefectCount DESC;";

                    // 1. Siapkan Dictionary untuk menampung hasil database sementara
                    var tempResults = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                    using (SqlCommand command = new SqlCommand(getDefectsByModel, connection))
                    {
                        command.Parameters.AddWithValue("@MachineCode", MachineCode);
                        command.Parameters.AddWithValue("@StartDate", startDateParsed);
                        command.Parameters.AddWithValue("@EndDate", endDateParsed);
                        addStationParameter(command); // Menggunakan action yang sudah Anda buat

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string stationName = reader.IsDBNull(0) ? "Unknown" : reader.GetString(0);
                                int qty = reader.IsDBNull(1) ? 0 : reader.GetInt32(1);

                                if (!tempResults.ContainsKey(stationName))
                                {
                                    tempResults.Add(stationName, qty);
                                }
                            }
                        }
                    }

                    // 2. Tentukan Master List Station berdasarkan MachineCode
                    // Ini memastikan nama station SELALU muncul di legend meskipun Qty 0
                    List<string> masterStations = new List<string>();

                    if (MachineCode == "MCH1-01")
                    {
                        // Sesuaikan nama ini PERSIS dengan yang ada di database/tombol HTML Anda
                        masterStations = new List<string> {
                            "PREPARING",
                            "GAS LEAK",
                            "STARTING - RUNNING",
                            "INNER",
                            "FINAL",
                            "DETAIL"
                        };
                    }
                    else if (MachineCode == "MCH1-02")
                    {
                        masterStations = new List<string> {
                            "Chassis",
                            "Preparing",
                            "Running",
                            "Final"
                        };
                    }

                    // 3. Gabungkan Master List dengan Hasil Database
                    // Jika filter Station aktif (user memilih 1 station), kita hanya tampilkan itu.
                    // Jika filter Station kosong (All Station), kita tampilkan semua list master.

                    if (string.IsNullOrEmpty(Station))
                    {
                        // Logika untuk "ALL STATION": Loop semua master station
                        foreach (var st in masterStations)
                        {
                            DefectsByModel.Add(new DefectByModel
                            {
                                ProductName = st, // Label Station
                                // Ambil qty dari DB jika ada, jika tidak ada set 0
                                Quantity = tempResults.ContainsKey(st) ? tempResults[st] : 0
                            });
                        }

                        // Opsional: Masukkan station lain yang mungkin ada di DB tapi tidak ada di master list (misal typo di DB)
                        foreach (var kvp in tempResults)
                        {
                            if (!masterStations.Contains(kvp.Key, StringComparer.OrdinalIgnoreCase))
                            {
                                DefectsByModel.Add(new DefectByModel { ProductName = kvp.Key, Quantity = kvp.Value });
                            }
                        }
                    }
                    else
                    {
                        // Logika jika user memfilter station tertentu (tetap gunakan hasil murni DB)
                        foreach (var kvp in tempResults)
                        {
                            DefectsByModel.Add(new DefectByModel { ProductName = kvp.Key, Quantity = kvp.Value });
                        }
                    }

                    //string getMonthlyDefects = $@"
                    //SELECT
                    //    DAY(SDate) AS DayOfMonth,
                    //    Cause,
                    //    COUNT(*) AS DefectCount
                    //FROM
                    //    NG_RPTS
                    //WHERE
                    //    MachineCode = @MachineCode
                    //    AND MONTH(SDate) = MONTH(GETDATE())
                    //    AND YEAR(SDate) = YEAR(GETDATE())
                    //    {stationFilterClause}
                    //GROUP BY
                    //    DAY(SDate), Cause
                    //ORDER BY
                    //    DayOfMonth, DefectCount DESC;";

                    //using (SqlCommand command = new SqlCommand(getMonthlyDefects, connection))
                    //{
                    //    command.Parameters.AddWithValue("@MachineCode", MachineCode);
                    //    addStationParameter(command);
                    //    using (SqlDataReader reader = command.ExecuteReader())
                    //    {
                    //        while (reader.Read())
                    //        {
                    //            MonthlyDefects.Add(new MonthlyDefectData
                    //            {
                    //                Day = reader.GetInt32(0),
                    //                Cause = reader.GetString(1),
                    //                Quantity = reader.GetInt32(2)
                    //            });
                    //        }
                    //    }
                    //}

                    string getYearlyDefects = $@"
                    SELECT
                        MONTH(SDate) AS MonthNumber, -- Ambil Angka Bulan (1-12)
                        Cause,
                        COUNT(*) AS DefectCount
                    FROM
                        NG_RPTS
                    WHERE
                        MachineCode = @MachineCode
                        AND YEAR(SDate) = YEAR(@StartDate) -- Filter berdasarkan TAHUN dari StartDate
                        {stationFilterClause}
                    GROUP BY
                        MONTH(SDate), Cause
                    ORDER BY
                        MonthNumber, DefectCount DESC;";

                    using (SqlCommand command = new SqlCommand(getYearlyDefects, connection))
                    {
                        command.Parameters.AddWithValue("@MachineCode", MachineCode);
                        command.Parameters.AddWithValue("@StartDate", startDateParsed); // Menggunakan tahun dari input user
                        addStationParameter(command);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                YearlyDefects.Add(new YearlyDefectData
                                {
                                    Month = reader.GetInt32(0), // Kolom 0 sekarang adalah Bulan
                                    Cause = reader.GetString(1),
                                    Quantity = reader.GetInt32(2)
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Database error: " + ex.Message;
                Console.WriteLine(errorMessage);
            }

            if (TotalPlan > 0)
            {
                DefectRatio = (1 - (double)DefectQuantity / TotalPlan) * 100;
            }
        }

        public class DailyDefect
        {
            public string Cause { get; set; }
            public int Quantity { get; set; }
        }

        public class DefectByModel
        {
            public string? ProductName { get; set; }
            public int Quantity { get; set; }
        }
        //public class MonthlyDefectData
        //{
        //    public int Day { get; set; }
        //    public string Cause { get; set; }
        //    public int Quantity { get; set; }
        //}
        public class YearlyDefectData
        {
            public int Month { get; set; } // Ganti Day jadi Month
            public string Cause { get; set; }
            public int Quantity { get; set; }
        }
    }
}