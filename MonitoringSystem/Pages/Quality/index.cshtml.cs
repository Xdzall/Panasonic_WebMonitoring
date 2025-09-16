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

        [BindProperty(SupportsGet = true)]
        public string MachineCode { get; set; }

        [BindProperty(SupportsGet = true)]
        public string StartDate { get; set; }

        [BindProperty(SupportsGet = true)]
        public string EndDate { get; set; }
        public QualityModel()
        {
            TopDailyDefects = new List<DailyDefect>();
            DefectProblems = new List<DailyDefect>();
            DefectsByModel = new List<DefectByModel>();
        }

        public void OnGet()
        {
            // Set a default machine code for the initial page load
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

            DateTime startDateParsed, endDateParsed;
            if (!DateTime.TryParse(StartDate, out startDateParsed))
            {
                startDateParsed = DateTime.Now.Date;
            }
            if (!DateTime.TryParse(EndDate, out endDateParsed))
            {
                endDateParsed = DateTime.Now.Date;
            }

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string getTotalProduction = @"
                    SELECT
                        SUM(TotalUnit)
                    FROM
                     OEESN
                    WHERE
                    MachineCode = @MachineCode
                    AND Date = (SELECT MAX(Date) FROM OEESN WHERE MachineCode = @MachineCode);";

                    using (SqlCommand command = new SqlCommand(getTotalProduction, connection))
                    {
                        command.Parameters.AddWithValue("@MachineCode", MachineCode);
                        var result = command.ExecuteScalar();
                        if (result != DBNull.Value && result != null)
                        {
                            TotalPlan = Convert.ToInt32(result);
                        }
                    }

                    string getTotalDefect = @"
                    SELECT
                        COUNT(*)
                    FROM
                        NG_RPTS
                    WHERE
                        CAST(SDate AS DATE) = (SELECT MAX(CAST(SDate AS DATE)) FROM NG_RPTS WHERE MachineCode = @MachineCode)
                        AND MachineCode = @MachineCode;";

                    using (SqlCommand command = new SqlCommand(getTotalDefect, connection))
                    {
                        command.Parameters.AddWithValue("@MachineCode", MachineCode);
                        var result = command.ExecuteScalar();
                        if (result != DBNull.Value && result != null)
                        {
                            DefectQuantity = Convert.ToInt32(result);
                        }
                    }

                    string getTopDailyDefect = @"
            SELECT TOP 5
                Cause,
                COUNT(*) AS DefectCount
            FROM
                NG_RPTS
            WHERE
                CAST(SDate AS DATE) BETWEEN @StartDate AND @EndDate 
                AND MachineCode = @MachineCode
            GROUP BY
                Cause
            ORDER BY
                DefectCount DESC;";

                    using (SqlCommand command = new SqlCommand(getTopDailyDefect, connection))
                    {
                        command.Parameters.AddWithValue("@MachineCode", MachineCode);
                        command.Parameters.AddWithValue("@StartDate", startDateParsed);
                        command.Parameters.AddWithValue("@EndDate", endDateParsed);

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

                    string getDefectProblem = @"
                    SELECT
                        Cause,
                        COUNT(*) AS DefectCount
                    FROM
                        NG_RPTS
                    WHERE
                        CAST(SDate AS DATE) BETWEEN @StartDate AND @EndDate
                        AND MachineCode = @MachineCode
                    GROUP BY
                        Cause
                    ORDER BY
                        DefectCount DESC;";
                    using (SqlCommand command = new SqlCommand(getDefectProblem, connection))
                    {
                        command.Parameters.AddWithValue("@MachineCode", MachineCode);
                        command.Parameters.AddWithValue("@StartDate", startDateParsed);
                        command.Parameters.AddWithValue("@EndDate", endDateParsed);

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


                    string getDefectsByModel = @"
                    SELECT
                        md.ProductName,
                        COUNT(*) AS DefectCount
                    FROM
                        NG_RPTS ng
                    JOIN
                        MasterData md ON 
                        -- Menyamakan tipe data INT dan VARCHAR menjadi teks yang bersih
                        LTRIM(RTRIM(CAST(ng.Product_Id AS VARCHAR(255)))) = LTRIM(RTRIM(CAST(md.Product_Id AS VARCHAR(255))))
                    WHERE
                        ng.MachineCode = @MachineCode
                        AND CAST(ng.SDate AS DATE) BETWEEN @StartDate AND @EndDate
                    GROUP BY
                        md.ProductName
                    ORDER BY
                        DefectCount DESC;";

                    using (SqlCommand command = new SqlCommand(getDefectsByModel, connection))
                    {
                        command.Parameters.AddWithValue("@MachineCode", MachineCode);
                        command.Parameters.AddWithValue("@StartDate", startDateParsed);
                        command.Parameters.AddWithValue("@EndDate", endDateParsed);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                DefectsByModel.Add(new DefectByModel
                                {
                                    ProductName = reader.IsDBNull(0) ? "Nama Produk Kosong" : reader.GetString(0),
                                    Quantity = reader.IsDBNull(1) ? 0 : reader.GetInt32(1)
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
    }
}