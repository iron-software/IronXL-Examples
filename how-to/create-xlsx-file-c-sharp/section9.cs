using IronXL.Excel;
namespace IronXL.Examples.HowTo.CreateXlsxFileCSharp
{
    public static class Section9
    {
        public static void Run()
        {
            /**
            Set Font Style
            anchor-set-font-style
            **/
            WorkSheet ["CellAddress"].Style.Font.Bold =true;
            WorkSheet ["CellAddress"].Style.Font.Italic =true;
        }
    }
}