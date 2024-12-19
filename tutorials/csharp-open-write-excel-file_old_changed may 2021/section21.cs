using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section21
    {
        public static void Run()
        {
            /**
            Import Data to Database
            anchor-fill-database-with-data-from-excel-sheet
            **/
            TestDbEntities dbContext = new TestDbEntities ();
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            WorkSheet sheet = workbook.GetWorkSheet("Sheet3");
            System.Data.DataTable dataTable = sheet.ToDataTable(true);
            foreach (DataRow row in dataTable.Rows)
            {
                Country c = new Country();
                c.CountryName = row [1].ToString();
                dbContext.Countries.Add(c);
            }
            dbContext.SaveChanges();
        }
    }
}