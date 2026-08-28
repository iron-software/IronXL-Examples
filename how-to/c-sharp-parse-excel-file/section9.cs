using System.Data;
using IronXL;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section9
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook WorkBook = IronXL.WorkBook.Load("sample.xlsx");

            DataSet ds = WorkBook.ToDataSet();
        }
    }
}