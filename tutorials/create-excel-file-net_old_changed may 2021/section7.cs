using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet_Old_Changed May 2021
{
    public static class Section7
    {
        public static void Run()
        {
            /**
            Set Cell Background Color
            anchor-set-background-colors-of-cells
            **/
            sheet ["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
        }
    }
}