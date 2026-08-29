using System.Data;
using IronXL;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section12
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");

            //import WorkBook into dataset
            DataSet ds = workBook.ToDataSet(true);
        }
    }
}