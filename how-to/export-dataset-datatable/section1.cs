using System.Data;
using IronXL.Excel;
namespace ironxl.ExportDatasetDatatable
{
    public class Section1
    {
        public void Run()
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