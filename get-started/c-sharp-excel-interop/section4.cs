using IronXL;
namespace IronXL.Examples.GettingStarted.CSharpExcelInterop
{
    public static class Section4
    {
        public static void Run()
        {
            ws.Rows[RowIndex].Replace("old value", "new value");
        }
    }
}