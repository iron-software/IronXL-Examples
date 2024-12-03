using IronXL;
using IronXL.Excel;
namespace ironxl.SetPasswordWorksheet
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Set protection for selected worksheet
            workSheet.ProtectSheet("IronXL");
            
            workBook.Save();
        }
    }
}