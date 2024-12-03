using System.Data;
using IronXL.Excel;
namespace ironxl.LoadSpreadsheet
{
    public class Section3
    {
        public void Run()
        {
            // Create dataset
            DataSet dataSet = new DataSet();
            
            // Create workbook
            WorkBook workBook = WorkBook.Create();
            
            // Load DataSet
            WorkBook.LoadWorkSheetsFromDataSet(dataSet, workBook);
        }
    }
}