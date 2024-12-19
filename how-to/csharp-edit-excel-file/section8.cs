using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpEditExcelFile
{
    public static class Section8
    {
        public static void Run()
        {
            ws ["B5:B10"].Replace("old value", "new value");
        }
    }
}