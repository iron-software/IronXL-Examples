using IronXL;
namespace IronXL.Examples.HowTo.CreateXlsxFileCSharp
{
    public static class Section5
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet worksheet = workBook.DefaultWorkSheet;

            worksheet ["CellAddress"].Value = "MyValue";
        }
    }
}