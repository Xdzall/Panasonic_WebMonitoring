using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using System.Text.Json;

namespace MonitoringSystem.Pages.Shared
{
    public class ProductionPlanModel : PageModel
    {
        public List<ProductName> listProducts = new List<ProductName>();
        public List<ProductionRecord> listRecords = new List<ProductionRecord>();
        //public string dbcon = "Data Source=DESKTOP-2VG5S76\\VE_SERVER;Initial Catalog=PROMOSYS;User ID=sa;Password=gerrys0803;";
        public string dbcon = "Server=10.83.33.103;User Id=sa;Password=sa;Database=PROMOSYS;Trusted_Connection=False;TrustServerCertificate=True;Encrypt=False";
        //public string dbcon = "Data Source=DESKTOP-NBPATD6\\MSSQLSERVERR;trusted_connection=true;trustservercertificate=True;Database=PROMOSYS;Integrated Security=True;Encrypt=False";

        public string? ProductNames { get; set; }
        public string? MachineCode { get; set; }
        public string? TotalQuantityCS { get; set; }
        public string? Comment { get; set; }
        public DateTime CurrentDate { get; set; }

        bool allFieldsEmpty = true;

        public void OnGet()
        {
            getListModelName();
            InsertProductionPlanNow();
            getTotalQuantityCS();
        }

        public IActionResult getListModelName()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();
                    string query = @"SELECT ProductName FROM Product WHERE MachineCode='MCH1-02';";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataReader dataReader = command.ExecuteReader())
                        {
                            while (dataReader.Read())
                            {
                                ProductName productName = new ProductName();
                                productName.Name = dataReader.GetString(0);
                                listProducts.Add(productName);
                            }
                        }
                    }
                }
                ProductNames = JsonSerializer.Serialize(listProducts);
                return Page();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.ToString());
                return Page();
            }
        }

        public void getTotalQuantityCS()
        {
            CurrentDate = DateTime.Now.Date;
            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();

                    string query = @"SELECT SUM(Quantity) AS Qty FROM ProductionRecords
                                     INNER JOIN ProductionPlan ON ProductionRecords.PlanId = ProductionPlan.Id
                                     WHERE ProductionPlan.CurrentDate = @CurrentDate AND ProductionRecords.MachineCode = 'MCH1-02';";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@CurrentDate", CurrentDate);
                        using (SqlDataReader dataReader = command.ExecuteReader())
                        {
                            while (dataReader.Read())
                            {
                                TotalQuantityCS = dataReader.GetInt32(0).ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.ToString());
            }
        }

        public void InsertProductionPlanNow()
        {
            CurrentDate = DateTime.Now.Date;
            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();

                    string queryCheck = @"SELECT COUNT(1) FROM ProductionPlan WHERE CurrentDate = @CurrentDate;";
                    using (SqlCommand commandCheck = new SqlCommand(queryCheck, connection))
                    {
                        commandCheck.Parameters.AddWithValue("@CurrentDate", CurrentDate);
                        int count = (int)commandCheck.ExecuteScalar();
                        if (count == 0)
                        {
                            string queryInsert = @"INSERT INTO ProductionPlan (CurrentDate) VALUES (@CurrentDate);";
                            using (SqlCommand commandInsert = new SqlCommand(queryInsert, connection))
                            {
                                commandInsert.Parameters.AddWithValue("@CurrentDate", CurrentDate);
                                commandInsert.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            string querySelectAllData = @"SELECT ProductionRecords.Id, ProductionRecords.ProductName, ProductionRecords.Quantity, MasterData.QtyHour, 
                                                          ROUND(CAST(ProductionRecords.Quantity As float)/CAST(MasterData.QtyHour AS float), 2) AS Hour, ProductionRecords.Lot, 
                                                          ProductionRecords.Remark FROM ProductionRecords INNER JOIN MasterData ON ProductionRecords.ProductName = MasterData.ProductName
                                                          INNER JOIN ProductionPlan ON ProductionRecords.PlanId = ProductionPlan.Id WHERE ProductionPlan.CurrentDate = @CurrentDate 
                                                          AND ProductionRecords.MachineCode = 'MCH1-02' ORDER BY Id;";

                            using (SqlCommand commandSelectAll = new SqlCommand(querySelectAllData, connection))
                            {
                                commandSelectAll.Parameters.AddWithValue("@CurrentDate", CurrentDate);
                                using (SqlDataReader dataReader = commandSelectAll.ExecuteReader())
                                {
                                    while (dataReader.Read())
                                    {
                                        ProductionRecord record = new ProductionRecord();
                                        record.Id = dataReader.GetInt32(0);
                                        record.ModelName = dataReader.GetString(1);
                                        record.Quantity = dataReader.GetInt32(2);
                                        record.QtyHour = dataReader.GetInt32(3);
                                        record.Hour = dataReader.GetDouble(4);
                                        record.Lot = dataReader.GetString(5);
                                        record.Remark = dataReader.GetString(6);

                                        listRecords.Add(record);
                                    }
                                }
                            }

                            string querySelectComment = @"SELECT Comment_CS FROM ProductionPlan WHERE CurrentDate = @CurrentDate";
                            using (SqlCommand commandSelectComment = new SqlCommand(querySelectComment, connection))
                            {
                                commandSelectComment.Parameters.AddWithValue("@CurrentDate", CurrentDate);
                                using (SqlDataReader dataComment = commandSelectComment.ExecuteReader())
                                {
                                    while (dataComment.Read())
                                    {
                                        Comment = dataComment.GetString(0);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.ToString());
            }
        }

        public IActionResult OnPostInsertProduct()
        {
            string productName = Request.Form["ProductName"];

            if (string.IsNullOrEmpty(productName))
            {
                TempData["StatusMessage"] = "error";
                TempData["Message"] = "Model Name is required.";
                return RedirectToPage();
            }

            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();
                    string query = @"INSERT INTO Product (ProductName, MachineCode) VALUES (@ProductName, 'MCH1-02');";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@ProductName", productName);
                        command.ExecuteNonQuery();
                    }
                }
                TempData["StatusMessage"] = "success";
                TempData["Message"] = "Product Model successfully inserted!";
                return RedirectToPage();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.ToString());

                TempData["StatusMessage"] = "error";
                TempData["Message"] = "Error inserting data: " + ex.Message;
                return Page();
            }
        }

        public IActionResult OnPostInsertProductionRecord(
            List<int?> IdModel,
            List<string> ModelName,
            List<int?> Quantity,
            List<int?> QtyHour,
            List<string> Lot,
            List<string> Remark,
            string Comment
        )
        {
            int planId = 0;
            CurrentDate = DateTime.Now.Date;

            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();
                    string querySelectPlanId = @"SELECT TOP 1 Id FROM ProductionPlan WHERE CurrentDate = @CurrentDate;";
                    using (SqlCommand commandSelectId = new SqlCommand(querySelectPlanId, connection))
                    {
                        commandSelectId.Parameters.AddWithValue("@CurrentDate", CurrentDate);
                        using (SqlDataReader dataReader = commandSelectId.ExecuteReader())
                        {
                            while (dataReader.Read()) { planId = dataReader.GetInt32(0); }
                        }
                    }

                    if (!string.IsNullOrEmpty(Comment))
                    {
                        string queryUpdate = @"UPDATE ProductionPlan SET Comment_CS = @Comment WHERE Id = @Id;";
                        using (SqlCommand commandUpdate = new SqlCommand(queryUpdate, connection))
                        {
                            commandUpdate.Parameters.AddWithValue("@Id", planId);
                            commandUpdate.Parameters.AddWithValue("@Comment", Comment);

                            commandUpdate.ExecuteNonQuery();
                        }
                    }

                    for (int i = 0; i < ModelName.Count; i++)
                    {
                        bool isModelNameEmpty = string.IsNullOrEmpty(ModelName[i]);
                        bool isQuantityEmpty = !Quantity[i].HasValue;
                        bool isLotEmpty = string.IsNullOrEmpty(Lot[i]);
                        bool isRemarkEmpty = string.IsNullOrEmpty(Remark[i]);

                        if (isModelNameEmpty && isQuantityEmpty && isLotEmpty && isRemarkEmpty)
                        {
                            continue;
                        }

                        allFieldsEmpty = false;

                        if (QtyHour[i].HasValue)
                        {
                            string queryUpdate = @"UPDATE MasterData SET QtyHour = @QtyHour WHERE ProductName = @ProductName;";
                            using (SqlCommand commandUpdate = new SqlCommand(queryUpdate, connection))
                            {
                                commandUpdate.Parameters.AddWithValue("@QtyHour", QtyHour[i]);
                                commandUpdate.Parameters.AddWithValue("@ProductName", ModelName[i]);

                                commandUpdate.ExecuteNonQuery();
                            }
                        }

                        if (IdModel[i].HasValue)
                        {
                            string queryUpdateRecord = @"UPDATE ProductionRecords SET ProductName = @ProductName, Quantity = @Quantity, Lot = @Lot, Remark = @Remark 
                                                         WHERE Id = @Id";
                            using (SqlCommand commandUpdatePR = new SqlCommand(queryUpdateRecord, connection))
                            {
                                commandUpdatePR.Parameters.AddWithValue("@Id", IdModel[i]);
                                commandUpdatePR.Parameters.AddWithValue("@ProductName", ModelName[i]);
                                commandUpdatePR.Parameters.AddWithValue("@Quantity", Quantity[i]);
                                commandUpdatePR.Parameters.AddWithValue("@Lot", Lot[i] ?? string.Empty);
                                commandUpdatePR.Parameters.AddWithValue("@Remark", Remark[i] ?? string.Empty);

                                commandUpdatePR.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            string querySelectCode = @"SELECT MachineCode FROM Product WHERE ProductName = @ProductName";
                            using (SqlCommand commandMachineCode = new SqlCommand(querySelectCode, connection))
                            {
                                commandMachineCode.Parameters.AddWithValue("@ProductName", ModelName[i]);
                                using (SqlDataReader reader = commandMachineCode.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        MachineCode = reader.GetString(0);
                                    }
                                }
                            }
                            string queryInsertRecords = @"INSERT INTO ProductionRecords (PlanID, ProductName, MachineCode, Quantity, Lot, Remark) 
                                                        VALUES (@PlanID, @ProductName, @MachineCode, @Quantity, @Lot, @Remark);";
                            using (SqlCommand commandInsertPR = new SqlCommand(queryInsertRecords, connection))
                            {
                                commandInsertPR.Parameters.AddWithValue("@PlanID", planId);
                                commandInsertPR.Parameters.AddWithValue("@ProductName", ModelName[i]);
                                commandInsertPR.Parameters.AddWithValue("@MachineCode", MachineCode);
                                commandInsertPR.Parameters.AddWithValue("@Quantity", Quantity[i]);
                                commandInsertPR.Parameters.AddWithValue("@Lot", Lot[i] ?? string.Empty);
                                commandInsertPR.Parameters.AddWithValue("@Remark", Remark[i] ?? string.Empty);

                                commandInsertPR.ExecuteNonQuery();
                            }
                        }
                    }

                    if (allFieldsEmpty)
                    {
                        TempData["StatusMessage"] = "error";
                        TempData["Message"] = "Please input at least one model with quantity.";
                        return RedirectToPage();
                    }

                    TempData["StatusMessage"] = "success";
                    TempData["Message"] = "Production Plan successfully updated!";
                    return RedirectToPage();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.ToString());

                TempData["StatusMessage"] = "error";
                TempData["Message"] = "Error inserting data: " + ex.Message;
                return Page();
            }
        }

        public async Task<IActionResult> OnPostDeleteRecordAsync()
        {
            string recordId = Request.Form["RecordId"];
            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();
                    string queryDelete = @"DELETE FROM ProductionRecords WHERE Id = @RecordId;";
                    using (SqlCommand commandDelete = new SqlCommand(queryDelete, connection))
                    {
                        commandDelete.Parameters.AddWithValue("@RecordId", recordId);
                        int rowsAffected = await commandDelete.ExecuteNonQueryAsync();
                        if (rowsAffected > 0)
                        {
                            TempData["StatusMessage"] = "success";
                            TempData["Message"] = "Data deleted successfully";
                            return RedirectToPage();
                        }
                        else
                        {
                            TempData["StatusMessage"] = "error";
                            TempData["Message"] = "Data not found";
                            return RedirectToPage();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.ToString());

                TempData["StatusMessage"] = "error";
                TempData["Message"] = "Error deleting data: " + ex.Message;
                return Page();
            }
        }

        public async Task<IActionResult> OnPostDeleteAllRecord()
        {
            int planId = 0;
            CurrentDate = DateTime.Now.Date;
            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();
                    string queryGetId = "SELECT Id FROM ProductionPlan WHERE CurrentDate = @CurrentDate;";
                    using (SqlCommand commandGetId = new SqlCommand(queryGetId, connection))
                    {
                        commandGetId.Parameters.AddWithValue("@CurrentDate", CurrentDate);
                        using (SqlDataReader dataReader = commandGetId.ExecuteReader())
                        {
                            while (dataReader.Read()) { planId = dataReader.GetInt32(0); }
                        }
                    }

                    string queryDelete = "DELETE FROM ProductionRecords WHERE PlanId = @PlanId;";
                    using (SqlCommand commandDelete = new SqlCommand(queryDelete, connection))
                    {
                        commandDelete.Parameters.AddWithValue("@PlanId", planId);
                        int rowsAffected = await commandDelete.ExecuteNonQueryAsync();
                        if (rowsAffected > 0)
                        {
                            TempData["StatusMessage"] = "success";
                            TempData["Message"] = "Data deleted successfully";
                            return RedirectToPage();
                        }
                        else
                        {
                            TempData["StatusMessage"] = "error";
                            TempData["Message"] = "Data not found";
                            return RedirectToPage();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.ToString());

                TempData["StatusMessage"] = "error";
                TempData["Message"] = "Error deleting data: " + ex.Message;
                return Page();
            }
        }

        public IActionResult OnPostUpdateProduct()
        {
            string id = Request.Form["id"];
            string ProductName = Request.Form["ProductName"];
            string Quantity = Request.Form["Quantity"];
            string QtyHour = Request.Form["QtyHour"];
            string Lot = Request.Form["Lot"];
            string Remark = Request.Form["Remark"];

            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();

                    if (QtyHour != null)
                    {
                        string queryUpdate = @"UPDATE MasterData SET QtyHour = @QtyHour WHERE ProductName = @ProductName;";
                        using (SqlCommand commandUpdate = new SqlCommand(queryUpdate, connection))
                        {
                            commandUpdate.Parameters.AddWithValue("@QtyHour", QtyHour);
                            commandUpdate.Parameters.AddWithValue("@ProductName", ProductName);

                            commandUpdate.ExecuteNonQuery();
                        }

                    }

                    string query = @"UPDATE ProductionRecords SET ProductName = @ProductName, Quantity = @Quantity, Lot = @Lot, Remark = @Remark WHERE Id = @Id";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Id", id);
                        command.Parameters.AddWithValue("@ProductName", ProductName ?? string.Empty);
                        command.Parameters.AddWithValue("@Quantity", Quantity ?? string.Empty);
                        command.Parameters.AddWithValue("@Lot", Lot ?? string.Empty);
                        command.Parameters.AddWithValue("@Remark", Remark ?? string.Empty);

                        command.ExecuteNonQuery();
                    }
                }
                TempData["StatusMessage"] = "success";
                TempData["Message"] = "Production Plan successfully updated!";
                return RedirectToPage();
            }
            catch (Exception ex)
            {
                TempData["StatusMessage"] = "error";
                TempData["Message"] = "Error updating data: " + ex.Message;
                return Page();
            }
        }

        public class ProductName
        {
            public string? Name { get; set; }
        }

        public class ProductionRecord
        {
            public int Id { get; set; }
            public string? ModelName { get; set; }
            public int? Quantity { get; set; }
            public int? QtyHour { get; set; }
            public double? Hour { get; set; }
            public string? Lot { get; set; }
            public string? Remark { get; set; }
        }
    }
}
