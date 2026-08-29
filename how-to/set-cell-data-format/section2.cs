using IronXL;
namespace IronXL.Examples.HowTo.SetCellDataFormat
{
    public static class Section2
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            // Assign value as string
            workSheet["A1"].StringValue = "4402-12";
        }
    }
}