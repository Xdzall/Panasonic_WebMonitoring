using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using System.Reflection.PortableExecutable;

namespace MonitoringSystem.Pages.Quality
{
    public class QualityModel : PageModel
    {
        public string connectionString = "Server=10.83.33.103;trusted_connection=false;Database=PROMOSYS;User Id=sa;Password=sa;Persist Security Info=False;Encrypt=False";

        // Properties to hold the results for the Razor Page
        public int TotalPlan { get; set; }
        public int DefectQuantity { get; set; }
        public double DefectRatio { get; set; }
        public string errorMessage = "";

        // Bind this property to the form input
        [BindProperty(SupportsGet = true)]
        public string MachineCode { get; set; }

        public void OnGet()
        {
            // Set a default machine code for the initial page load
            if (string.IsNullOrEmpty(MachineCode))
            {
                MachineCode = "MCH1-01";
            }
            LoadData();
        }

        public void OnPost()
        {
            // MachineCode property is automatically bound from the submitted form
            Console.WriteLine($"MachineCode yang diterima: {MachineCode}");
            LoadData();
        }

        private void LoadData()
        {
            // Reset values to avoid stale data
            TotalPlan = 0;
            DefectQuantity = 0;
            DefectRatio = 100;

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // Query to get Total Production
                    //string getTotalProduction = @"
                    //    SELECT SUM(Quantity) 
                    //    FROM ProductionRecords
                    //    JOIN ProductionPlan ON ProductionRecords.PlanId = ProductionPlan.Id
                    //    WHERE ProductionRecords.MachineCode = @MachineCode AND CAST(ProductionPlan.CurrentDate AS DATE) = CAST(GETDATE() AS DATE);";
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

                    // Query to get Total Defect
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
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Database error: " + ex.Message;
                Console.WriteLine(errorMessage);
            }

            // Calculate Defect Ratio after fetching data
            if (TotalPlan > 0)
            {
                DefectRatio = (1 - (double)DefectQuantity / TotalPlan) * 100;
            }
        }

        // You can keep this method and modify it to use the MachineCode property
        // if you plan to call it from a different part of the code
        public void GetDailyDefect()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string GetDailyDefect = @"SELECT Cause, COUNT(*) FROM NG_RPTS WHERE CAST(SDate AS DATE) = @Date AND MachineCode = @MachineCode GROUP BY Cause";
                    using (SqlCommand command = new SqlCommand(GetDailyDefect, connection))
                    {
                        command.Parameters.AddWithValue("@Date", DateTime.Now.Date);
                        command.Parameters.AddWithValue("@MachineCode", MachineCode);
                        // You'll need to read the results here
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.ToString());
            }
        }

        public class DailyDefect
        {
            public string Cause { get; set; }
            public int Quantity { get; set; }
        }
    }
}