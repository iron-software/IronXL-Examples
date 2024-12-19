using IronXL.Excel;
namespace IronXL.Examples.HowTo.CreateXlsxFileCSharp
{
    public static class Section11
    {
        public static void Run()
        {
            /**
            Set Border Style
            anchor-set-border-style
            **/
            WorkSheet ["CellAddress"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dotted;
        }
    }
}