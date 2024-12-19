using System.Data;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.ExportDatasetDatatable
{
    public static class Section1
    {
        public static void Run()
        {
            // Create dataset
            DataSet dataSet = new DataSet();
            
            // Create workbook
            WorkBook workBook = WorkBook.Create();
            
            // Load DataSet to workBook
            WorkBook.LoadWorkSheetsFromDataSet(dataSet, workBook);
        }
    }
}