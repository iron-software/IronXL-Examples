using IronXL;
namespace IronXL.Examples.HowTo.CSharpExcelInterop
{
    public static class Section2
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet ws = workBook.DefaultWorkSheet;

            ws ["A3:C3"].Value = "New Value";
        }
    }
}