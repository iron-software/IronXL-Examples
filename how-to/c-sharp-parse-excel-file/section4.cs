using IronXL;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section4
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet ws = workBook.DefaultWorkSheet;

            /**
            Parse into Boolean Values
            anchor-parse-excel-data-into-boolean-values
            **/
            bool Val = ws ["Cell Address"].BoolValue;
        }
    }
}