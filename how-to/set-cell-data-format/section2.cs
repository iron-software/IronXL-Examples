using IronXL.Excel;
namespace ironxl.SetCellDataFormat
{
    public class Section2
    {
        public void Run()
        {
            // Assign value as string
            workSheet["A1"].StringValue = "4402-12";
        }
    }
}