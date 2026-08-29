using System.Data;
using IronXL;
namespace IronXL.Examples.GettingStarted.CSharpExcelInterop
{
    public static class Section1
    {
        public static void Run()
        {
            // Access WorkBook and WorkSheet
            WorkBook wb = WorkBook.Load("sample.xlsx");
            WorkSheet ws = wb.GetWorkSheet("Sheet1");
            
            // Convert workbook to DataSet
            DataSet ds = wb.ToDataSet();
            
            // Convert worksheet to DataTable
            DataTable dt = ws.ToDataTable(true);
        }
    }
}