using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet_Old_Changed May 2021
{
    public static class Section8
    {
        public static void Run()
        {
            /**
            Create Borders
            anchor-create-borders
            **/
            sheet ["A1:L1"].Style.TopBorder.SetColor("#000000");
            sheet ["A1:L1"].Style.BottomBorder.SetColor("#000000");
            
            sheet ["L2:L11"].Style.RightBorder.SetColor("#000000");
            sheet ["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
            
            sheet ["A11:L11"].Style.BottomBorder.SetColor("#000000");
            sheet ["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
        }
    }
}