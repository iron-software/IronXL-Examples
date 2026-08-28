using IronXL;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section8
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet WorkSheet = workBook.DefaultWorkSheet;

            //to find the sum of specific cell range 
            decimal total = WorkSheet ["Starting Cell Address : Ending Cell Address"].Sum();
        }
    }
}