using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.SetPasswordWorksheet
{
    public static class Section1
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Set protection for selected worksheet
            workSheet.ProtectSheet("IronXL");
            
            workBook.Save();
        }
    }
}