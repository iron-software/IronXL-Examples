using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section5
    {
        public static void Run()
        {
            var array = WorkSheet ["From:To"].ToArray();
        }
    }
}