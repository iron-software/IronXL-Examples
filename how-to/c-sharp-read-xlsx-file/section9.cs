using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section9
    {
        public static void Run()
        {
            decimal sum = ws ["From:To"].Sum();
            decimal min = ws ["From:To"].Min();
            decimal max = ws ["From:To"].Max();
        }
    }
}