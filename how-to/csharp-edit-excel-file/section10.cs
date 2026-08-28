using IronXL;
namespace IronXL.Examples.HowTo.CsharpEditExcelFile
{
    public static class Section10
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook wb = IronXL.WorkBook.Load("sample.xlsx");

            wb.RemoveWorkSheet("Sheet1"); //by sheet name
        }
    }
}