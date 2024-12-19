using IronXL.Excel;
namespace IronXL.Examples.HowTo.SetPasswordWorkbook
{
    public static class Section1
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            
            // Open protected spreadsheet file
            WorkBook protectedWorkBook = WorkBook.Load("sample.xlsx", "IronSoftware");
            
            // Set protection for spreadsheet file
            workBook.Encrypt("IronSoftware");
            
            workBook.Save();
        }
    }
}