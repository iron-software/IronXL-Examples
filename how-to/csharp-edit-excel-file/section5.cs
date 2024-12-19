using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpEditExcelFile
{
    public static class Section5
    {
        public static void Run()
        {
            ws ["From Cell Address : To Cell Address"].Replace("old value", "new value");
        }
    }
}