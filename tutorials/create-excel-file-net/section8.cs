using IronXL.Styles;
using IronXL;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet
{
    public static class Section8
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            workSheet["A1:L1"].Style.TopBorder.SetColor("#000000");
            workSheet["A1:L1"].Style.BottomBorder.SetColor("#000000");
            workSheet["L2:L11"].Style.RightBorder.SetColor("#000000");
            workSheet["L2:L11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
            workSheet["A11:L11"].Style.BottomBorder.SetColor("#000000");
            workSheet["A11:L11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;
        }
    }
}