using IronXL;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet
{
    public static class Section3
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");

            WorkSheet workSheet = workBook.CreateWorkSheet("2020 Budget");
        }
    }
}