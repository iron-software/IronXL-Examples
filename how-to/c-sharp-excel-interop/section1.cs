using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpExcelInterop
{
    public static class Section1
    {
        public static void Run()
        {
            /**
            DataSet and DataTables
            anchor-dataset-and-datatables
            **/
            //Access WorkBook.          
            WorkBook wb = WorkBook.Load("sample.xlsx");
            //Access WorkSheet.
             WorkSheet ws = wb.GetWorkSheet("Sheet1");
            //Behave with a workbook as Dataset.
            DataSet ds = wb.ToDataSet(); 
            //Behave with workbook as DataTable
            DataTable dt = ws.ToDataTable(true);
        }
    }
}