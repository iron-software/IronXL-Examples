using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpEditExcelFile
{
    public static class Section6
    {
        public static void Run()
        {
            ws ["B4:E4"].Replace("old value", "new value");
        }
    }
}