using IronXL.Excel;
namespace ironxl.SetPasswordWorkbook
{
    public class Section1
    {
        public void Run()
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