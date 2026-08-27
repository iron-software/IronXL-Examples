using IronXL;
namespace IronXL.Examples.GettingStarted.CSharpExcelInterop
{
    public static class Section6
    {
        public static void Run()
        {
            ws["From:To"].Replace("old value", "new value");
        }
    }
}