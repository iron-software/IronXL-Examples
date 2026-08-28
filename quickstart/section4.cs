using IronXL.Styles;
using IronXL;
namespace IronXL.Examples.Overview.Quickstart
{
    public static class Section4
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            // Set cell's value and styles
            workSheet["A1"].Value = "Hello World";
            workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;
        }
    }
}