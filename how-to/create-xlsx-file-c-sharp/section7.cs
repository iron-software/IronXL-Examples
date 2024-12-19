using IronXL.Excel;
namespace IronXL.Examples.HowTo.CreateXlsxFileCSharp
{
    public static class Section7
    {
        public static void Run()
        {
            ws1 ["A3:A8"].Value = "NewValue";
        }
    }
}