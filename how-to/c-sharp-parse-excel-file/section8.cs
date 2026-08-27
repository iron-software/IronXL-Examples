using IronXL;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section8
    {
        public static void Run()
        {
            DataTable dt=WorkSheet.ToDataTable(True);
        }
    }
}