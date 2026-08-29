using IronXL;
namespace IronXL.Examples.HowTo.CSharpReadExcelFileExample
{
    public static class Section2
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook wb = IronXL.WorkBook.Load("sample.xlsx");

            //Open Excel WorkSheet
            WorkSheet ws = wb.GetWorkSheet("SheetName");
        }
    }
}