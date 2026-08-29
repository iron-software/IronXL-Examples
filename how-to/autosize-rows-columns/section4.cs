using IronXL;
namespace IronXL.Examples.HowTo.AutosizeRowsColumns
{
    public static class Section4
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            workSheet.Merge("A1:A3");
            
            workSheet.AutoSizeRow(0, false);
            workSheet.AutoSizeRow(1, false);
            workSheet.AutoSizeRow(2, false);
        }
    }
}