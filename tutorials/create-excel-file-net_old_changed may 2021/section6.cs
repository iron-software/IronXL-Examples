using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet_Old_Changed May 2021
{
    public static class Section6
    {
        public static void Run()
        {
            /**
            Add Cells from Database
            anchor-add-directly-from-a-database
            **/
            //Create database objects to populate data from database
            string contring;
            string sql;
            DataSet ds = new DataSet("DataSetName");
            SqlConnection con;
            SqlDataAdapter da;
            
            //Set Database Connection string
            contring = @"Data Source=Server_Name;Initial Catalog=Database_Name;User ID=User_ID;Password=Password";
            
            //SQL Query to obtain data
            sql = "SELECT Field_Names FROM Table_Name";
            
            //Open Connection & Fill DataSet
            con = new SqlConnection(contring);
            da = new SqlDataAdapter(sql, con);
            
            con.Open();
            da.Fill(ds);
            
            //Loop through contents of dataset
            foreach (DataTable table in ds.Tables)
            {
                 int Count = table.Rows.Count - 1;
            
                 for (int j = 12; j <= 21; j++)
                 {
                   sheet ["A" + j].Value = table.Rows [Count]["Field_Name_1"].ToString();
                   sheet ["B" + j].Value = table.Rows [Count]["Field_Name_2"].ToString();
                   sheet ["C" + j].Value = table.Rows [Count]["Field_Name_3"].ToString();
                   sheet ["D" + j].Value = table.Rows [Count]["Field_Name_4"].ToString();
                   sheet ["E" + j].Value = table.Rows [Count]["Field_Name_5"].ToString();
                   sheet ["F" + j].Value = table.Rows [Count]["Field_Name_6"].ToString();
                   sheet ["G" + j].Value = table.Rows [Count]["Field_Name_7"].ToString();
                   sheet ["H" + j].Value = table.Rows [Count]["Field_Name_8"].ToString();
                   sheet ["I" + j].Value = table.Rows [Count]["Field_Name_9"].ToString();
                   sheet ["J" + j].Value = table.Rows [Count]["Field_Name_10"].ToString();
                   sheet ["K" + j].Value = table.Rows [Count]["Field_Name_11"].ToString();
                   sheet ["L" + j].Value = table.Rows [Count]["Field_Name_12"].ToString();
                 }
                 Count++;
            }
        }
    }
}