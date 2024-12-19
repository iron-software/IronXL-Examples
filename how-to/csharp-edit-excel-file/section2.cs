using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpEditExcelFile
{
    public static class Section2
    {
        public static void Run()
        {
            ws ["A3:E3"].Value = "New Value";
        }
    }
}