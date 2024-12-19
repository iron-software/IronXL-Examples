using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpExcelInterop
{
    public static class Section6
    {
        public static void Run()
        {
            WorkSheet ["From:To"].Replace("old value", "new value");
        }
    }
}