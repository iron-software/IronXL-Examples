using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpCreateExcel
{
    public static class Section9
    {
        public static void Run()
        {
            //bold the text of specified range cells
            WorkSheet ["FromCellAddress : ToCellAddress"].Style.Font.Bold =true;
            
            //Italic the text of specified range cells
            WorkSheet ["FromCellAddress : ToCellAddress"].Style.Font.Italic =true;
            
            //Strikeout the text of specified range cells
            WorkSheet ["FromCellAddress : ToCellAddress"].Style.Font.Strikeout = true;
            
            //border style of specified range cells 
            WorkSheet ["FromCellAddress : ToCellAddress"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dotted;
            
            //border color of specified range cells 
            WorkSheet ["FromCellAddress : ToCellAddress"].Style.BottomBorder.SetColor("color value");
        }
    }
}