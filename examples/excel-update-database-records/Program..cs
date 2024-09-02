using IronXL;
using System.Data;
using System.Data.SqlClient;

// Supported for XLSX, XLS, XLSM, XLTX, CSV and TSV
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Convert the workbook to ToDataSet
DataSet dataSet = workBook.ToDataSet();

// Your sql query
string sql = "SELECT * FROM Users";

// Your connection string
string connectionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=usersdb;Integrated Security=True";
using (SqlConnection connection = new SqlConnection(connectionString))
{
    // Open connections to the database
    connection.Open();
    SqlDataAdapter adapter = new SqlDataAdapter(sql, connection);

    // Update the values in database using the values in Excel
    adapter.Update(dataSet);
}
