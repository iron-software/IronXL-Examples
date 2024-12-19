using IronXL.Excel;
namespace IronXL.Examples.HowTo.CreateXlsxFileCSharp
{
    public static class Section10
    {
        public static void Run()
        {
            /**
            Add Strikeout
            anchor-add-strikeout
            **/
            WorkSheet ["CellAddress"].Style.Font.Strikeout = true;
        }
    }
}