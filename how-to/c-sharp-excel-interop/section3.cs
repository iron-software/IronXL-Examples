using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpExcelInterop
{
    public static class Section3
    {
        public static void Run()
        {
            WorkSheet.Replace("old value", "new value");
        }
    }
}