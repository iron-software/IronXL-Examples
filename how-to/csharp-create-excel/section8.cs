using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpCreateExcel
{
    public static class Section8
    {
        public static void Run()
        {
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