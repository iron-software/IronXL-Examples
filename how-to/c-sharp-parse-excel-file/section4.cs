using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section4
    {
        public static void Run()
        {
            /**
            Parse into Boolean Values
            anchor-parse-excel-data-into-boolean-values
            **/
            bool Val = ws ["Cell Address"].BoolValue;
        }
    }
}