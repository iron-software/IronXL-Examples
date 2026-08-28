using IronXL;
namespace IronXL.Examples.HowTo.CsharpEditExcelFile
{
    public static class Section1
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet ws = workBook.DefaultWorkSheet;

            ws ["B4"].Value = "New Value"; //alternative way to access specific cell and apply changes
        }
    }
}