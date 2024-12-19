using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet
{
    public static class Section6
    {
        public static void Run()
        {
            // Create database objects to populate data from database
            string contring;
            string sql;
            DataSet ds = new DataSet("DataSetName");
            SqlConnection con;
            SqlDataAdapter da;
            
            // Set Database Connection string
            contring = @"Data Source=Server_Name;Initial Catalog=Database_Name;User ID=User_ID;Password=Password";
            
            // SQL Query to obtain data
            sql = "SELECT Field_Names FROM Table_Name";
            
            // Open Connection & Fill DataSet
            con = new SqlConnection(contring);
            da = new SqlDataAdapter(sql, con);
            con.Open();
            da.Fill(ds);
            
            // Loop through contents of dataset
            foreach (DataTable table in ds.Tables)
            {
                int Count = table.Rows.Count - 1;
                for (int j = 12; j <= 21; j++)
                {
                    workSheet["A" + j].Value = table.Rows[Count]["Field_Name_1"].ToString();
                    workSheet["B" + j].Value = table.Rows[Count]["Field_Name_2"].ToString();
                    workSheet["C" + j].Value = table.Rows[Count]["Field_Name_3"].ToString();
                    workSheet["D" + j].Value = table.Rows[Count]["Field_Name_4"].ToString();
                    workSheet["E" + j].Value = table.Rows[Count]["Field_Name_5"].ToString();
                    workSheet["F" + j].Value = table.Rows[Count]["Field_Name_6"].ToString();
                    workSheet["G" + j].Value = table.Rows[Count]["Field_Name_7"].ToString();
                    workSheet["H" + j].Value = table.Rows[Count]["Field_Name_8"].ToString();
                    workSheet["I" + j].Value = table.Rows[Count]["Field_Name_9"].ToString();
                    workSheet["J" + j].Value = table.Rows[Count]["Field_Name_10"].ToString();
                    workSheet["K" + j].Value = table.Rows[Count]["Field_Name_11"].ToString();
                    workSheet["L" + j].Value = table.Rows[Count]["Field_Name_12"].ToString();
                }
                Count++;
            }
        }
    }
}