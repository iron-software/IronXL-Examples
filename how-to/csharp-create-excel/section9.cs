using IronXL.Styles;
using IronXL;
namespace IronXL.Examples.HowTo.CsharpCreateExcel
{
    public static class Section9
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet WorkSheet = workBook.DefaultWorkSheet;

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