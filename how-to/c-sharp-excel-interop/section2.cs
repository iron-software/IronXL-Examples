using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpExcelInterop
{
    public static class Section2
    {
        public static void Run()
        {
            ws ["A3:C3"].Value = "New Value";
        }
    }
}