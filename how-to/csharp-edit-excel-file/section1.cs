using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpEditExcelFile
{
    public static class Section1
    {
        public static void Run()
        {
            ws ["B4"].Value = "New Value"; //alternative way to access specific cell and apply changes
        }
    }
}