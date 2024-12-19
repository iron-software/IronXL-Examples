using IronXL.Excel;
namespace IronXL.Examples.HowTo.CreateXlsxFileCSharp
{
    public static class Section6
    {
        public static void Run()
        {
            /**
            Insert WorkSheet Data
            anchor-insert-data-into-worksheets
            **/
            ws1 ["A1"].Value = "Hello World";
        }
    }
}