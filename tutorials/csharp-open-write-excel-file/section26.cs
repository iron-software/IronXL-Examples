using IronXL;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section26
    {
        public static void Run()
        {
            // This snippet belongs to the tutorial's Entity Framework sample, which the page names but does not print.
            // Kept verbatim; see README.md for the full context.
            // TestDbEntities dbContext = new TestDbEntities();
            // var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            // WorkSheet sheet = workbook.CreateWorkSheet("FromDb");
            // List<Country> countryList = dbContext.Countries.ToList();
            // sheet.SetCellValue(0, 0, "Id");
            // sheet.SetCellValue(0, 1, "Country Name");
            // int row = 1;
            // foreach (var item in countryList)
            // {
            // sheet.SetCellValue(row, 0, item.id);
            // sheet.SetCellValue(row, 1, item.CountryName);
            // row++;
            // }
            // workbook.SaveAs("FilledFile.xlsx");
        }
    }
}