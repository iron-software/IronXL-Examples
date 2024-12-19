using IronXL.Excel;
namespace IronXL.Examples.HowTo.SetCellDataFormat
{
    public static class Section2
    {
        public static void Run()
        {
            // Assign value as string
            workSheet["A1"].StringValue = "4402-12";
        }
    }
}