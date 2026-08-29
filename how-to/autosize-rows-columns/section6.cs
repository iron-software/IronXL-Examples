using IronXL;
namespace IronXL.Examples.HowTo.AutosizeRowsColumns
{
    public static class Section6
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            workSheet.Merge("A1:B1");
            
            workSheet.AutoSizeColumn(0, false);
            workSheet.AutoSizeColumn(1, false);
        }
    }
}