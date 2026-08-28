using IronXL;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp
{
    public static class Section8
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            Range range = workSheet["D2:D101"];
        }
    }
}