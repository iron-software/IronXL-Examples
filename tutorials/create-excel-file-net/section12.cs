using IronXL;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet
{
    public static class Section12
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");

            workBook.SaveAs("Budget.xlsx");
        }
    }
}