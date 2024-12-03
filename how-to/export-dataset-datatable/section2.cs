using System.Data;
using IronXL.Excel;
namespace ironxl.ExportDatasetDatatable
{
    public class Section2
    {
        public void Run()
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