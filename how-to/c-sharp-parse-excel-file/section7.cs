using System.Data;
using IronXL;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section7
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.

            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook WorkBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet WorkSheet = WorkBook.DefaultWorkSheet;

            DataTable dt = WorkSheet.ToDataTable(true);
        }
    }
}