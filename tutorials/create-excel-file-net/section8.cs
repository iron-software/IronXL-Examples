using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet
{
    public static class Section8
    {
        public static void Run()
        {
            workSheet["A1:L1"].Style.TopBorder.SetColor("#000000");
            workSheet["A1:L1"].Style.BottomBorder.SetColor("#000000");
            workSheet["L2:L11"].Style.RightBorder.SetColor("#000000");
            workSheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
            workSheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
            workSheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
        }
    }
}