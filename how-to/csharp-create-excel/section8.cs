using IronXL.Styles;
using IronXL;
namespace IronXL.Examples.HowTo.CsharpCreateExcel
{
    public static class Section8
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet WorkSheet = workBook.DefaultWorkSheet;

            //bold the text of specified cell
            WorkSheet ["CellAddress"].Style.Font.Bold =true;
            
            //Italic the text of specified cell
            WorkSheet ["CellAddress"].Style.Font.Italic =true;
            
            //Strikeout the text of specified cell
            WorkSheet ["CellAddress"].Style.Font.Strikeout = true;
            
            //border style of specific cell 
            WorkSheet ["CellAddress"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dotted;
            
            //border color of specific cell 
            WorkSheet ["CellAddress"].Style.BottomBorder.SetColor("color value");
        }
    }
}