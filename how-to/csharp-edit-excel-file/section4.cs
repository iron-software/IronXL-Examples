using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpEditExcelFile
{
    public static class Section4
    {
        public static void Run()
        {
            ws.Rows [2].Replace("old value", "new value");
        }
    }
}