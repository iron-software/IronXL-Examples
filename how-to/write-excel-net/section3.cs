using IronXL;
namespace IronXL.Examples.HowTo.WriteExcelNet
{
    public static class Section3
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            workSheet["Cell Address"].Value="Assign the Value";
        }
    }
}