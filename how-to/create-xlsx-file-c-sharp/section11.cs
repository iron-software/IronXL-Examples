using IronXL.Styles;
using IronXL;
namespace IronXL.Examples.HowTo.CreateXlsxFileCSharp
{
    public static class Section11
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet WorkSheet = workBook.DefaultWorkSheet;

            /**
            Set Border Style
            anchor-set-border-style
            **/
            WorkSheet ["CellAddress"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dotted;
        }
    }
}