using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using System.Text.Json;
using OfficeOpenXml;


namespace MonitoringSystem.Pages.Shared
{
    public class ProductionPlanCUModel : PageModel
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public List<ProductName> listProducts = new List<ProductName>();
        public List<ProductionRecord> listRecords = new List<ProductionRecord>();
        //public string dbcon = "Data Source=DESKTOP-2VG5S76\\VE_SERVER;Initial Catalog=PROMOSYS;User ID=sa;Password=gerrys0803;";
        public string dbcon = "Server=10.83.33.103;User Id=sa;Password=sa;Database=PROMOSYS;Trusted_Connection=False;TrustServerCertificate=True;Encrypt=False";
        //public string dbcon = "Data Source=DESKTOP-NBPATD6\\MSSQLSERVERR;trusted_connection=true;trustservercertificate=True;Database=PROMOSYS;Integrated Security=True;Encrypt=False";

        public string? ProductNames { get; set; }
        public string? MachineCode { get; set; }
        public string? TotalQuantity { get; set; }
        public string? TotalOvertime { get; set; }
        public string? GrandTotal { get; set; }
        public string? Comment { get; set; }
        public DateTime CurrentDate { get; set; }

        [BindProperty(SupportsGet = true)]
        public string? FilterMachineCode { get; set; }

        [BindProperty(SupportsGet = true)]
        public DateTime? FilterDate { get; set; }

        [BindProperty(SupportsGet = true)]
        public List<string>? FilterShifts { get; set; }


        bool allFieldsEmpty = true;

        public void OnGet()
        {
            if (string.IsNullOrEmpty(FilterMachineCode)) FilterMachineCode = "MCH1-01";

            CurrentDate = FilterDate.HasValue ? FilterDate.Value.Date : DateTime.Now.Date;

            getListModelName();
            InsertProductionPlanNow();
            getTotalQuantity();
        }

        public IActionResult getListModelName()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();
                    string query = @"SELECT ProductName FROM Product WHERE MachineCode = @MachineCode;";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@MachineCode", FilterMachineCode ?? "MCH1-01");
                        using (SqlDataReader dataReader = command.ExecuteReader())
                        {
                            while (dataReader.Read())
                            {
                                listProducts.Add(new ProductName { Name = dataReader.GetString(0) });
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

        public void getTotalQuantity()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();

                    // Ambil SUM Quantity dan SUM Overtime
                    string query = @"
                    SELECT 
                        SUM(PR.Quantity) as TotalNormal, 
                        SUM(PR.Overtime) as TotalOvt 
                    FROM ProductionRecords PR
                    INNER JOIN ProductionPlan PP ON PR.PlanId = PP.Id
                    WHERE PP.CurrentDate = @CurrentDate 
                    AND PR.MachineCode = @MachineCode;";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@CurrentDate", CurrentDate);
                        command.Parameters.AddWithValue("@MachineCode", FilterMachineCode ?? "MCH1-01");

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                // Ambil nilai (handle null dengan 0)
                                int normal = reader.IsDBNull(0) ? 0 : reader.GetInt32(0);
                                int ovt = reader.IsDBNull(1) ? 0 : reader.GetInt32(1);

                                // Set Properti
                                TotalQuantity = normal.ToString();
                                TotalOvertime = ovt.ToString();

                                // GABUNGKAN (JUMLAHKAN)
                                GrandTotal = (normal + ovt).ToString();
                            }
                            else
                            {
                                TotalQuantity = "0";
                                TotalOvertime = "0";
                                GrandTotal = "0";
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
            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();

                    // 1. Pastikan Plan Hari ini ada
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
                    }

                    // 2. AMBIL DATA (SELECT) DENGAN FILTER
                    // Kita akan membangun query dinamis berdasarkan FilterShifts
                    string shiftCondition = "";
                    if (FilterShifts != null && FilterShifts.Count > 0)
                    {
                        // Logic filter shift agak kompleks karena data di DB string (csv). 
                        // Untuk simpelnya, kita filter di Memory atau asumsikan user ingin melihat semua data mesin tersebut.
                        // Disini saya filter berdasarkan Machine Code dulu.
                    }

                    string querySelectAllData = @"
                        SELECT 
                            PR.Id, PR.ProductName, PR.Quantity, MD.QtyHour, 
                            ROUND(CAST(PR.Quantity As float)/CAST(MD.QtyHour AS float), 2) AS Hour, 
                            PR.Lot, PR.Remark,
                            PR.Overtime, PR.NoDirectOfWorker, PR.NoDirectOfWorkerOvertime, PR.Shift
                        FROM ProductionRecords PR
                        LEFT JOIN MasterData MD ON PR.ProductName = MD.ProductName
                        INNER JOIN ProductionPlan PP ON PR.PlanId = PP.Id 
                        WHERE PP.CurrentDate = @CurrentDate 
                        AND PR.MachineCode = @MachineCode
                        ORDER BY PR.Id DESC;"; // Order desc agar data baru diatas

                    using (SqlCommand commandSelectAll = new SqlCommand(querySelectAllData, connection))
                    {
                        commandSelectAll.Parameters.AddWithValue("@CurrentDate", CurrentDate);
                        commandSelectAll.Parameters.AddWithValue("@MachineCode", FilterMachineCode ?? "MCH1-01");

                        using (SqlDataReader dataReader = commandSelectAll.ExecuteReader())
                        {
                            while (dataReader.Read())
                            {
                                ProductionRecord record = new ProductionRecord();
                                record.Id = dataReader.GetInt32(0);
                                record.ModelName = dataReader.IsDBNull(1) ? "" : dataReader.GetString(1);
                                record.Quantity = dataReader.IsDBNull(2) ? 0 : dataReader.GetInt32(2);
                                record.QtyHour = dataReader.IsDBNull(3) ? 0 : dataReader.GetInt32(3);
                                record.Hour = dataReader.IsDBNull(4) ? 0 : dataReader.GetDouble(4);
                                record.Lot = dataReader.IsDBNull(5) ? "" : dataReader.GetString(5);
                                record.Remark = dataReader.IsDBNull(6) ? "" : dataReader.GetString(6);

                                // Mapping Kolom Baru
                                record.Overtime = dataReader.IsDBNull(7) ? null : dataReader.GetInt32(7);
                                record.NoDirectOfWorker = dataReader.IsDBNull(8) ? null : dataReader.GetInt32(8);
                                record.NoDirectOfWorkerOvertime = dataReader.IsDBNull(9) ? null : dataReader.GetInt32(9);
                                record.Shift = dataReader.IsDBNull(10) ? "" : dataReader.GetString(10);

                                listRecords.Add(record);
                            }
                        }
                    }

                    string commentColumn = (FilterMachineCode == "MCH1-02") ? "Comment_CS" : "Comment_CU";

                    // Gunakan variable commentColumn di dalam query
                    string querySelectComment = $"SELECT {commentColumn} FROM ProductionPlan WHERE CurrentDate = @CurrentDate";

                    using (SqlCommand commandSelectComment = new SqlCommand(querySelectComment, connection))
                    {
                        commandSelectComment.Parameters.AddWithValue("@CurrentDate", CurrentDate);
                        using (SqlDataReader dataComment = commandSelectComment.ExecuteReader())
                        {
                            if (dataComment.Read() && !dataComment.IsDBNull(0))
                            {
                                Comment = dataComment.GetString(0);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception Load Data: " + ex.ToString());
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
                    string query = @"INSERT INTO Product (ProductName, MachineCode) VALUES (@ProductName, 'MCH1-01');";
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
                TempData["Message"] = "Error inserting product: " + ex.Message;
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
    List<int?> Overtime, // Opsional
    List<int?> NoOfDirectWorker, // Wajib (Normal)
    List<int?> NoOfDirectWorkerOvertime, // Opsional
    string Comment,
    DateTime TargetDate
)
        {
            int planId = 0;
            CurrentDate = TargetDate != DateTime.MinValue ? TargetDate : DateTime.Now.Date;

            // Tangkap Filter dari Hidden Input
            string filterMachine = Request.Form["FilterMachineCode"];
            if (!string.IsNullOrEmpty(filterMachine)) FilterMachineCode = filterMachine;

            // Flag untuk tracking status simpan
            bool hasInvalidRows = false;
            int savedRowsCount = 0;

            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();

                    // 1. Dapatkan atau Buat Plan ID
                    string querySelectPlanId = @"SELECT TOP 1 Id FROM ProductionPlan WHERE CurrentDate = @CurrentDate;";
                    using (SqlCommand commandSelectId = new SqlCommand(querySelectPlanId, connection))
                    {
                        commandSelectId.Parameters.AddWithValue("@CurrentDate", CurrentDate);
                        var res = commandSelectId.ExecuteScalar();
                        if (res != null) planId = (int)res;
                        else
                        {
                            string qInsPlan = @"INSERT INTO ProductionPlan (CurrentDate) VALUES (@CurrentDate); SELECT SCOPE_IDENTITY();";
                            using (SqlCommand cIns = new SqlCommand(qInsPlan, connection))
                            {
                                cIns.Parameters.AddWithValue("@CurrentDate", CurrentDate);
                                planId = Convert.ToInt32(cIns.ExecuteScalar());
                            }
                        }
                    }

                    // 2. Update Komentar
                    if (!string.IsNullOrEmpty(Comment) && planId > 0)
                    {
                        string targetColumn = (FilterMachineCode == "MCH1-02") ? "Comment_CS" : "Comment_CU";
                        string queryUpdate = $"UPDATE ProductionPlan SET {targetColumn} = @Comment WHERE Id = @Id;";
                        using (SqlCommand cmd = new SqlCommand(queryUpdate, connection))
                        {
                            cmd.Parameters.AddWithValue("@Id", planId);
                            cmd.Parameters.AddWithValue("@Comment", Comment);
                            cmd.ExecuteNonQuery();
                        }
                    }

                    // 3. Loop Data (SAFE LOOP)
                    for (int i = 0; i < ModelName.Count; i++)
                    {
                        // --- SAFE ACCESSORS ---
                        string safeModelName = (ModelName != null && ModelName.Count > i) ? ModelName[i] : "";
                        int? safeQty = (Quantity != null && Quantity.Count > i) ? Quantity[i] : null;
                        int? safeWorker = (NoOfDirectWorker != null && NoOfDirectWorker.Count > i) ? NoOfDirectWorker[i] : null;

                        // Data Opsional / Lainnya
                        int? safeQtyHour = (QtyHour != null && QtyHour.Count > i) ? QtyHour[i] : null;
                        string safeLot = (Lot != null && Lot.Count > i) ? Lot[i] : null;
                        string safeRemark = (Remark != null && Remark.Count > i) ? Remark[i] : null;
                        int? safeOvertime = (Overtime != null && Overtime.Count > i) ? Overtime[i] : null;
                        int? safeWorkerOvt = (NoOfDirectWorkerOvertime != null && NoOfDirectWorkerOvertime.Count > i) ? NoOfDirectWorkerOvertime[i] : null;

                        // A. Cek Apakah Baris Kosong Total (User tidak isi apa-apa) -> Skip Silent
                        bool isRowEmpty = string.IsNullOrEmpty(safeModelName) &&
                                          (!safeQty.HasValue || safeQty == 0) &&
                                          (!safeWorker.HasValue);

                        if (isRowEmpty) continue;

                        // B. VALIDASI WAJIB: Product Name, Quantity, dan Worker Normal HARUS ADA
                        bool isRowValid = !string.IsNullOrEmpty(safeModelName) &&
                                          (safeQty.HasValue && safeQty > 0) &&
                                          safeWorker.HasValue;

                        if (!isRowValid)
                        {
                            // Jika data tidak lengkap, tandai error dan LEWATI baris ini
                            hasInvalidRows = true;
                            continue;
                        }

                        // C. LOGIC SHIFT OTOMATIS (Default "NS")
                        string shiftValue = "NS"; // Default value
                        string shiftKey = $"Shift[{i}]";
                        if (Request.Form.ContainsKey(shiftKey))
                        {
                            // Jika user mencentang checkbox, gunakan nilainya
                            shiftValue = string.Join(",", Request.Form[shiftKey]);
                        }
                        // Double check jika string kosong, paksa "NS"
                        if (string.IsNullOrEmpty(shiftValue)) shiftValue = "NS";


                        // D. Update Master Data QtyHour
                        if (safeQtyHour.HasValue && !string.IsNullOrEmpty(safeModelName))
                        {
                            string qUpdMaster = @"UPDATE MasterData SET QtyHour = @QtyHour WHERE ProductName = @ProductName;";
                            using (SqlCommand cmd = new SqlCommand(qUpdMaster, connection))
                            {
                                cmd.Parameters.AddWithValue("@QtyHour", safeQtyHour);
                                cmd.Parameters.AddWithValue("@ProductName", safeModelName);
                                cmd.ExecuteNonQuery();
                            }
                        }

                        // E. EKSEKUSI SQL (INSERT / UPDATE)
                        int? safeId = (IdModel != null && IdModel.Count > i) ? IdModel[i] : null;

                        // Tentukan Query
                        string querySQL = "";
                        if (safeId.HasValue && safeId > 0)
                        {
                            querySQL = @"UPDATE ProductionRecords 
                                 SET ProductName=@Pn, Quantity=@Qty, Lot=@Lot, Remark=@Rem, 
                                     Overtime=@Ovt, NoDirectOfWorker=@WNorm, NoDirectOfWorkerOvertime=@WOvt, Shift=@Sh
                                 WHERE Id=@Id";
                        }
                        else
                        {
                            querySQL = @"INSERT INTO ProductionRecords 
                                (PlanID, ProductName, MachineCode, Quantity, Lot, Remark, Overtime, NoDirectOfWorker, NoDirectOfWorkerOvertime, Shift) 
                                VALUES (@Pid, @Pn, @Mc, @Qty, @Lot, @Rem, @Ovt, @WNorm, @WOvt, @Sh);";
                        }

                        using (SqlCommand cmd = new SqlCommand(querySQL, connection))
                        {
                            // Parameter Wajib
                            cmd.Parameters.AddWithValue("@Pn", safeModelName);
                            cmd.Parameters.AddWithValue("@Qty", safeQty);
                            cmd.Parameters.AddWithValue("@WNorm", safeWorker);
                            cmd.Parameters.AddWithValue("@Sh", shiftValue); // Shift otomatis NS

                            // Parameter Opsional (Bisa Null)
                            cmd.Parameters.AddWithValue("@Ovt", (object)safeOvertime ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@WOvt", (object)safeWorkerOvt ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@Lot", (object)safeLot ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@Rem", (object)safeRemark ?? DBNull.Value);

                            // Parameter Kondisional (Id vs Insert)
                            if (safeId.HasValue && safeId > 0)
                            {
                                cmd.Parameters.AddWithValue("@Id", safeId);
                            }
                            else
                            {
                                cmd.Parameters.AddWithValue("@Pid", planId);
                                // Cari Machine Code
                                string mCode = FilterMachineCode ?? "MCH1-01";
                                string qM = "SELECT TOP 1 MachineCode FROM Product WHERE ProductName = @Pn";
                                using (SqlCommand cM = new SqlCommand(qM, connection))
                                {
                                    cM.Parameters.AddWithValue("@Pn", safeModelName);
                                    var resM = cM.ExecuteScalar();
                                    if (resM != null) mCode = resM.ToString();
                                }
                                cmd.Parameters.AddWithValue("@Mc", mCode);
                            }

                            cmd.ExecuteNonQuery();
                            savedRowsCount++;
                        }
                    } // End Loop

                    // 4. FEEDBACK MESSAGE
                    if (savedRowsCount > 0)
                    {
                        if (hasInvalidRows)
                        {
                            // Berhasil sebagian
                            TempData["StatusMessage"] = "warning"; // Icon Warning (Kuning)
                            TempData["Message"] = "Data Saved, but some rows were SKIPPED because Product Name, Quantity, or Normal Worker were empty.";
                        }
                        else
                        {
                            // Berhasil semua
                            TempData["StatusMessage"] = "success";
                            TempData["Message"] = "All Production Plan saved successfully!";
                        }
                    }
                    else
                    {
                        // Tidak ada yang tersimpan sama sekali
                        if (hasInvalidRows)
                        {
                            TempData["StatusMessage"] = "error";
                            TempData["Message"] = "Action Failed! Please fill in Product Name, Quantity, and Worker (Normal) for at least one row.";
                        }
                        else
                        {
                            TempData["StatusMessage"] = "info";
                            TempData["Message"] = "No data to save.";
                        }
                    }

                    return RedirectToPage(new { FilterDate = CurrentDate.ToString("yyyy-MM-dd"), FilterMachineCode = FilterMachineCode });
                }
            }
            catch (Exception ex)
            {
                TempData["StatusMessage"] = "error";
                TempData["Message"] = "Error: " + ex.Message;
                return RedirectToPage(new { FilterDate = CurrentDate.ToString("yyyy-MM-dd"), FilterMachineCode = FilterMachineCode });
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
                TempData["Message"] = "Error deleting records: " + ex.Message;
                return RedirectToPage(new
                {
                    FilterDate = CurrentDate.ToString("yyyy-MM-dd"),
                    FilterMachineCode = FilterMachineCode
                });
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
                            return RedirectToPage(new
                            {
                                FilterDate = CurrentDate.ToString("yyyy-MM-dd"),
                                FilterMachineCode = FilterMachineCode
                            });
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
            string id = Request.Form["Id"];
            string ProductName = Request.Form["ProductName"];
            string Quantity = Request.Form["Quantity"];
            string QtyHour = Request.Form["QtyHour"];
            string Lot = Request.Form["Lot"]; // Jika ada input Lot di modal
            string Remark = Request.Form["Remark"]; // Jika ada input Remark di modal

            // Tangkap Kolom Baru
            string Overtime = Request.Form["Overtime"];
            string NoOfDirectWorker = Request.Form["NoOfDirectWorker"];
            string NoOfDirectWorkerOvertime = Request.Form["NoOfDirectWorkerOvertime"];

            string targetDateString = Request.Form["TargetDate"];
            DateTime targetDate = DateTime.Now.Date; // Default fallback
            if (DateTime.TryParse(targetDateString, out DateTime parsedDate))
            {
                targetDate = parsedDate;
            }
            string shiftValue = "";
            if (Request.Form.ContainsKey("Shift"))
            {
                shiftValue = string.Join(",", Request.Form["Shift"]);
            }

            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();

                    // Update Master Data QtyHour
                    if (!string.IsNullOrEmpty(QtyHour))
                    {
                        string queryUpdate = @"UPDATE MasterData SET QtyHour = @QtyHour WHERE ProductName = @ProductName;";
                        using (SqlCommand commandUpdate = new SqlCommand(queryUpdate, connection))
                        {
                            commandUpdate.Parameters.AddWithValue("@QtyHour", QtyHour);
                            commandUpdate.Parameters.AddWithValue("@ProductName", ProductName);
                            commandUpdate.ExecuteNonQuery();
                        }
                    }

                    // Update ProductionRecords (LENGKAP)
                    string query = @"UPDATE ProductionRecords 
                             SET ProductName = @ProductName, 
                                 Quantity = @Quantity,
                                 Overtime = @Overtime,
                                 NoDirectOfWorker = @WNorm,
                                 NoDirectOfWorkerOvertime = @WOvt,
                                 Shift = @Shift
                             WHERE Id = @Id";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Id", id);
                        command.Parameters.AddWithValue("@ProductName", ProductName ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@Quantity", Quantity ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@Overtime", Overtime ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@WNorm", NoOfDirectWorker ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@WOvt", NoOfDirectWorkerOvertime ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@Shift", shiftValue);

                        command.ExecuteNonQuery();
                    }
                }
                TempData["StatusMessage"] = "success";
                TempData["Message"] = "Data successfully updated!";
                return RedirectToPage(new { FilterDate = targetDate.ToString("yyyy-MM-dd"), FilterMachineCode = FilterMachineCode });
            }
            catch (Exception ex)
            {
                TempData["StatusMessage"] = "error";
                TempData["Message"] = "Error updating data: " + ex.Message;
                return RedirectToPage(new
                {
                    FilterDate = targetDate.ToString("yyyy-MM-dd"),
                    FilterMachineCode = FilterMachineCode
                });
            }
        }

        [HttpPost]
        public async Task<IActionResult> OnPostSubmitCounter([FromBody] SubmitCount submitCount)
        {
            if (submitCount == null)
            {
                return BadRequest();
            }

            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();
                    string queryInsert = @"INSERT INTO SubmitCounts (SubmitCount, Timestamp) VALUES (1, GETDATE());";
                    using (SqlCommand commandInsert = new SqlCommand(queryInsert, connection))
                    {
                        await commandInsert.ExecuteNonQueryAsync();
                    }
                }

                // Ambil nilai terbaru setelah insert
                int updatedCount = 0;
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();
                    string queryCount = @"SELECT COUNT(*) FROM SubmitCounts WHERE CAST(Timestamp AS DATE) = @CurrentDate;";
                    using (SqlCommand commandCount = new SqlCommand(queryCount, connection))
                    {
                        commandCount.Parameters.AddWithValue("@CurrentDate", DateTime.Now.Date);
                        updatedCount = (int)commandCount.ExecuteScalar();
                    }
                }

                return new JsonResult(new { success = true, count = updatedCount });
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.ToString());
                return new JsonResult(new { success = false, message = "Internal server error" });
            }
        }


        [HttpGet]
        [Route("/OnGetGetSubmitCounter")]

        public async Task<IActionResult> OnGetGetSubmitCounter()
        {
            int submitCount = 0;
            CurrentDate = DateTime.Now.Date;

            try
            {
                using (SqlConnection connection = new SqlConnection(dbcon))
                {
                    connection.Open();
                    string query = @"SELECT COUNT(*) FROM SubmitCounts WHERE CAST(Timestamp AS DATE) = @CurrentDate;";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@CurrentDate", CurrentDate);
                        submitCount = (int)command.ExecuteScalar();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.ToString());
                return new JsonResult(new { success = false, message = "Error fetching submit count" });
            }

            return new JsonResult(new { success = true, count = submitCount });
        }

        public async Task<IActionResult> OnPostUploadAsync(IFormFile UploadedFile, string TargetMachine, int TargetMonth, int TargetYear)
        {
            if (UploadedFile == null || UploadedFile.Length == 0)
            {
                TempData["StatusMessage"] = "error";
                TempData["Message"] = "File Excel tidak ditemukan.";
                return RedirectToPage();
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            int totalSaved = 0;
            int daysInMonth = DateTime.DaysInMonth(TargetYear, TargetMonth);

            try
            {
                using (var stream = new MemoryStream())
                {
                    await UploadedFile.CopyToAsync(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.Rows;

                        using (SqlConnection connection = new SqlConnection(dbcon))
                        {
                            connection.Open();
                            using (SqlTransaction transaction = connection.BeginTransaction())
                            {
                                try
                                {
                                    // Loop Baris (Mulai baris 3)
                                    for (int row = 3; row <= rowCount; row++)
                                    {
                                        // PERUBAHAN 1: Baca Model Name dari Kolom 2 (B)
                                        string modelName = worksheet.Cells[row, 2].Value?.ToString()?.Trim();

                                        if (string.IsNullOrEmpty(modelName)) continue;

                                        for (int day = 1; day <= daysInMonth; day++)
                                        {
                                            // PERUBAHAN 2: Geser pembacaan Data Data
                                            // Sebelumnya: 2 + ((day - 1) * 2)  -> Mulai Kolom 2 (B)
                                            // Sekarang:   3 + ((day - 1) * 2)  -> Mulai Kolom 3 (C)
                                            // Karena Kolom A = No, Kolom B = Model Name, Data mulai Kolom C

                                            int colNormal = 3 + ((day - 1) * 2);
                                            int colOvertime = colNormal + 1;

                                            var valNormal = worksheet.Cells[row, colNormal].Value;
                                            var valOvertime = worksheet.Cells[row, colOvertime].Value;

                                            int qtyNormal = 0;
                                            int qtyOvertime = 0;

                                            if (valNormal != null) int.TryParse(valNormal.ToString(), out qtyNormal);
                                            if (valOvertime != null) int.TryParse(valOvertime.ToString(), out qtyOvertime);

                                            if (qtyNormal > 0 || qtyOvertime > 0)
                                            {
                                                DateTime currentDate = new DateTime(TargetYear, TargetMonth, day);
                                                int planId = GetOrCreatePlanId(connection, transaction, currentDate);
                                                InsertRecordFromExcel(connection, transaction, planId, modelName, TargetMachine, qtyNormal, qtyOvertime);
                                                totalSaved++;
                                            }
                                        }
                                    }

                                    transaction.Commit();
                                    TempData["StatusMessage"] = "success";
                                    TempData["Message"] = $"Upload Berhasil! {totalSaved} record produksi berhasil disimpan.";
                                }
                                catch (Exception ex)
                                {
                                    transaction.Rollback();
                                    throw ex;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Upload Error: " + ex.ToString());
                TempData["StatusMessage"] = "error";
                TempData["Message"] = "Gagal memproses file: " + ex.Message;
            }

            DateTime redirectDate = new DateTime(TargetYear, TargetMonth, 1);
            return RedirectToPage(new { FilterDate = redirectDate.ToString("yyyy-MM-dd"), FilterMachineCode = TargetMachine });
        }

        // --- HELPER METHODS (Supaya kode rapi) ---

        private int GetOrCreatePlanId(SqlConnection conn, SqlTransaction trans, DateTime date)
        {
            // Cek apakah Plan ID untuk tanggal ini sudah ada?
            string queryCheck = "SELECT Id FROM ProductionPlan WHERE CurrentDate = @Date";
            using (SqlCommand cmd = new SqlCommand(queryCheck, conn, trans))
            {
                cmd.Parameters.AddWithValue("@Date", date);
                var res = cmd.ExecuteScalar();
                if (res != null) return (int)res;
            }

            // Jika belum ada, buat baru
            string queryInsert = "INSERT INTO ProductionPlan (CurrentDate) VALUES (@Date); SELECT SCOPE_IDENTITY();";
            using (SqlCommand cmd = new SqlCommand(queryInsert, conn, trans))
            {
                cmd.Parameters.AddWithValue("@Date", date);
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private void InsertRecordFromExcel(SqlConnection conn, SqlTransaction trans, int planId, string modelName, string machineCode, int qty, int ovt)
        {
            string query = @"
        IF EXISTS (SELECT 1 FROM ProductionRecords WHERE PlanId = @PlanId AND ProductName = @Pn AND MachineCode = @Mc)
        BEGIN
            -- Data Sudah Ada: UPDATE
            UPDATE ProductionRecords 
            SET Quantity = @Qty, 
                Overtime = @Ovt
            WHERE PlanId = @PlanId AND ProductName = @Pn AND MachineCode = @Mc;
        END
        ELSE
        BEGIN
            -- Data Belum Ada: INSERT
            INSERT INTO ProductionRecords 
            (PlanId, ProductName, MachineCode, Quantity, Overtime, NoDirectOfWorker, NoDirectOfWorkerOvertime, Shift) 
            VALUES 
            (@PlanId, @Pn, @Mc, @Qty, @Ovt, 0, 0, 'NS');
        END";

            using (SqlCommand cmd = new SqlCommand(query, conn, trans))
            {
                cmd.Parameters.AddWithValue("@PlanId", planId);
                cmd.Parameters.AddWithValue("@Pn", modelName);
                cmd.Parameters.AddWithValue("@Mc", machineCode);
                cmd.Parameters.AddWithValue("@Qty", qty);

                // Handle Overtime Nullable
                if (ovt > 0) cmd.Parameters.AddWithValue("@Ovt", ovt);
                else cmd.Parameters.AddWithValue("@Ovt", DBNull.Value);

                cmd.ExecuteNonQuery();
            }
        }

        public IActionResult OnGetDownloadTemplate()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "productionplan", "ProductionPlan_Template.xlsx");
            if (!System.IO.File.Exists(filePath)) return NotFound("File template tidak ditemukan di server.");
            var bytes = System.IO.File.ReadAllBytes(filePath);
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ProductionPlan_Template.xlsx");
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
            public int? Overtime { get; set; }
            public int? NoDirectOfWorker { get; set; }
            public int? NoDirectOfWorkerOvertime { get; set; }
            public string? Shift { get; set; }
            public int? QtyHour { get; set; }
            public double? Hour { get; set; }
            public string? Lot { get; set; }
            public string? Remark { get; set; }
        }

        public class SubmitCount
        {
            public int Id { get; set; }
            public int SubmitCounter { get; set; }
            public DateTime Timestamp { get; set; }
        }
    }
}
