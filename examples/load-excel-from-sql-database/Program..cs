using IronXL;
using System.Data;
using System.Data.SqlClient;

// Your sql query
string sql = "SELECT * FROM Users";

// Your connection string
string connectionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=usersdb;Integrated Security=True";

using (SqlConnection connection = new SqlConnection(connectionString))
{
    // Open connections to the database
    connection.Open();
    SqlDataAdapter adapter = new SqlDataAdapter(sql, connection);
    DataSet ds = new DataSet();
    // Fill DataSet with data
    adapter.Fill(ds);

    // Create an Excel workbook from the SQL DataSet
    WorkBook workBook = WorkBook.Load(ds);
}
