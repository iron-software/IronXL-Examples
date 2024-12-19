using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section3
    {
        public static void Run()
        {
            /**
            Import WorkSheet 
            anchor-access-worksheet-for-project
            **/
            //by sheet indexing
            WorkBook.WorkSheets [SheetIndex];
            //get default  WorkSheet
            WorkBook.DefaultWorkSheet;
            //get first WorkSheet
            WorkBook.WorkSheets.First();
            //for the first or default sheet:
            WorkBook.WorkSheets.FirstOrDefault();
        }
    }
}