using System.Data;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.ExportDatasetDatatable
{
    public static class Section2
    {
        public static void Run()
        {
            // Create new Excel WorkBook document
            WorkBook workBook = WorkBook.Create();
            
            // Create a blank WorkSheet
            WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");
            
            // Export as DataSet
            DataSet dataSet = workBook.ToDataSet();
        }
    }
}