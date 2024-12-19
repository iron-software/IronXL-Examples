using IronXL.Excel;
namespace IronXL.Examples.HowTo.WriteExcelNet
{
    public static class Section5
    {
        public static void Run()
        {
            // Assign value as string
            workSheet["A1"].StringValue = "4402-12";
        }
    }
}