using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpakSoftware.Models;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;

namespace OpakSoftware.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        public string tableName = "UploadedData";
        private readonly string _connectionString;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _connectionString = configuration.GetConnectionString("DefaultConnection");
        }

        public IActionResult Index()
        {
            var dataTable = GetDataFromDatabase();
            var model = new UploadedDataViewModel
            {
                ColumnNames = dataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList(),
                Rows = dataTable.Rows.Cast<DataRow>().Select(r => r.ItemArray.Select(i => i.ToString()).ToList()).ToList()
            };
            return View(model);
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        #region Excel Operations
        [HttpPost]
        public IActionResult Upload(IFormFile file)
        {
            if (file != null && file.Length > 0)
            {
                var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", file.FileName);
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }
                ProcessExcelFile(filePath);
            }
            return RedirectToAction("Index");
        }
        public IActionResult ExportToExcel()
        {
            var dataTable = GetDataFromDatabase();
            var memoryStream = ExportDataTableToExcel(dataTable);

            return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExportedData.xlsx");
        }
        private MemoryStream ExportDataTableToExcel(DataTable dataTable)
        {
            using (var workbook = new XSSFWorkbook())
            {
                ISheet sheet = workbook.CreateSheet("ExportedData");
                IRow headerRow = sheet.CreateRow(0);

                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    headerRow.CreateCell(col).SetCellValue(dataTable.Columns[col].ColumnName);
                }

                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    IRow dataRow = sheet.CreateRow(row + 1);
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        dataRow.CreateCell(col).SetCellValue(dataTable.Rows[row][col].ToString());
                    }
                }

                var memoryStream = new MemoryStream();
                workbook.Write(memoryStream);
                memoryStream.Flush();
                //memoryStream.Position = 0;
                return memoryStream;
            }
        }

        private void ProcessExcelFile(string filePath)
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(stream);
                ISheet sheet = workbook.GetSheetAt(0);
                var columnNames = new List<string>();

                IRow headerRow = sheet.GetRow(0);
                for (int col = 0; col < headerRow.LastCellNum; col++)
                {
                    columnNames.Add(headerRow.GetCell(col).StringCellValue);
                }

                CreateDatabaseTable(columnNames);

                for (int row = 1; row <= sheet.LastRowNum; row++)
                {
                    IRow dataRow = sheet.GetRow(row);
                    var values = new List<string>();
                    for (int col = 0; col < dataRow.LastCellNum; col++)
                    {
                        values.Add(dataRow.GetCell(col)?.ToString() ?? string.Empty);
                    }
                    InsertDataIntoTable(columnNames, values);
                }
            }
        }
        #endregion

        #region Sql Operations
        private void CreateDatabaseTable(List<string> columnNames)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                var commandText = $"IF OBJECT_ID('{tableName}', 'U') IS NOT NULL DROP TABLE {tableName}; CREATE TABLE {tableName} (";
                var filteredColumnNames = columnNames.Where(c => c != "Id").ToList();
                commandText += string.Join(",", filteredColumnNames.ConvertAll(c => $"[{c}] NVARCHAR(MAX)"));
                commandText += ", [Id] INT IDENTITY(1,1) PRIMARY KEY)";

                using (var command = new SqlCommand(commandText, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        private void InsertDataIntoTable(List<string> columnNames, List<string> values)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                var commandText = $"INSERT INTO {tableName} (";
                commandText += string.Join(",", columnNames.Where(c => c != "Id").Select(c => $"[{c}]"));
                commandText += ") VALUES (";
                commandText += string.Join(",", values.Where((v, index) => columnNames[index] != "Id").Select(v => $"'{v.Replace("'", "''")}'"));
                commandText += ");";
                using (var command = new SqlCommand(commandText, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        [HttpPost]
        public IActionResult TruncateTable()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                using (var command = new SqlCommand($"TRUNCATE TABLE {tableName}", connection))
                {
                    command.ExecuteNonQuery();
                }
            }
            return Ok();
        }

        private DataTable GetDataFromDatabase()
        {
            var dataTable = new DataTable();

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                var tableExistsQuery = $"IF OBJECT_ID('{tableName}', 'U') IS NOT NULL SELECT 1 ELSE SELECT 0";
                var tableExists = false;
                using (var checkTableCommand = new SqlCommand(tableExistsQuery, connection))
                {
                    tableExists = (int)checkTableCommand.ExecuteScalar() == 1;
                }

                if (!tableExists)
                {
                    return dataTable;
                }

                var schemaCommandText = $"SELECT TOP 0 * FROM {tableName}";
                using (var schemaCommand = new SqlCommand(schemaCommandText, connection))
                {
                    using (var schemaAdapter = new SqlDataAdapter(schemaCommand))
                    {
                        schemaAdapter.Fill(dataTable);
                    }
                }

                var selectCommandText = $"SELECT * FROM {tableName}";
                using (var selectCommand = new SqlCommand(selectCommandText, connection))
                {
                    using (var adapter = new SqlDataAdapter(selectCommand))
                    {
                        adapter.Fill(dataTable);
                    }
                }
            }
            return dataTable;
        }
        #endregion

        #region CRUD Operations
        [HttpPost]
        public IActionResult UpdateRow(string keyColumnName, string keyValue, string columnNames)
        {
            var columns = columnNames.Split(',').ToList();

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                var setClause = string.Join(",", columns.Where(c => c != keyColumnName)
                    .Select(col => $"[{col}] = @{col}"));

                var commandText = $"UPDATE {tableName} SET {setClause} WHERE [{keyColumnName}] = @keyValue";

                using (var command = new SqlCommand(commandText, connection))
                {
                    foreach (var column in columns)
                    {
                        if (column != keyColumnName)
                        {
                            var value = Request.Form[column].ToString();
                            command.Parameters.AddWithValue($"@{column}", value);
                        }
                    }
                    command.Parameters.AddWithValue("@keyValue", keyValue);
                    command.ExecuteNonQuery();
                }
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public IActionResult DeleteRow(string keyColumnName, string keyValue)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                var commandText = $"DELETE FROM {tableName} WHERE [{keyColumnName}] = @keyValue";
                using (var command = new SqlCommand(commandText, connection))
                {
                    command.Parameters.AddWithValue("@keyValue", keyValue);
                    command.ExecuteNonQuery();
                }
            }
            return RedirectToAction("Index");
        }

        public IActionResult AddRow(string columnNames, string[] values)
        {
            var columns = columnNames.Split(',').ToList();
            var tableName = "UploadedData";

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                columns = columns.Where(c => c != "Id").ToList();

                var columnPlaceholders = string.Join(",", columns.Select(c => $"[{c}]"));
                var valuePlaceholders = string.Join(",", columns.Select((c, index) => $"@value{index}"));
                var commandText = $"INSERT INTO {tableName} ({columnPlaceholders}) VALUES ({valuePlaceholders})";

                using (var command = new SqlCommand(commandText, connection))
                {
                    for (int i = 0; i < columns.Count; i++)
                    {
                        command.Parameters.AddWithValue($"@value{i}", values[i]);
                    }
                    command.ExecuteNonQuery();
                }
            }
            return RedirectToAction("Index");
        }
        #endregion
    }
}
