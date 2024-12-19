using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpEditExcelFile
{
    public static class Section3
    {
        public static void Run()
        {
            /**
            Replace Cell Values
            anchor-replace-specific-value-of-complete-worksheet
            **/
            ws.Replace("old value", "new value");
        }
    }
}