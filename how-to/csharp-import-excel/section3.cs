using System.Linq;
using IronXL;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section3
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            int SheetIndex = 0;

            /**
            Import WorkSheet 
            anchor-access-worksheet-for-project
            **/
            //by sheet indexing
            WorkSheet byIndex = workBook.WorkSheets [SheetIndex];
            //get default  WorkSheet
            WorkSheet defaultSheet = workBook.DefaultWorkSheet;
            //get first WorkSheet
            WorkSheet firstSheet = workBook.WorkSheets.First();
            //for the first or default sheet:
            WorkSheet firstOrDefaultSheet = workBook.WorkSheets.FirstOrDefault();
        }
    }
}