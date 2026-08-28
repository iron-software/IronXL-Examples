using IronXL;
namespace IronXL.Examples.HowTo.CsharpEditExcelFile
{
    public static class Section6
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet ws = workBook.DefaultWorkSheet;

            ws ["B4:E4"].Replace("old value", "new value");
        }
    }
}