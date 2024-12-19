using System.Data;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.LoadSpreadsheet
{
    public static class Section3
    {
        public static void Run()
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