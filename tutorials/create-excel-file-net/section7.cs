using IronXL;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet
{
    public static class Section7
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            workSheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
        }
    }
}