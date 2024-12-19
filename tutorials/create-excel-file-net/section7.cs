using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet
{
    public static class Section7
    {
        public static void Run()
        {
            workSheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
        }
    }
}