TestDbEntities dbContext = new TestDbEntities();
var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet3");
System.Data.DataTable dataTable = sheet.ToDataTable(true);
foreach (DataRow row in dataTable.Rows)
{
    Country c = new Country();
    c.CountryName = row[1].ToString();
    dbContext.Countries.Add(c);
}
dbContext.SaveChanges();