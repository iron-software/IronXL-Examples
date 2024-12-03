using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section7
    {
        public void Run()
        {
            workSheet["A1:L1"].Style.SetBackgroundColor("#d3d3d3");
        }
    }
}