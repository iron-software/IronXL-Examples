using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section7
    {
        public static void Run()
        {
            DataTable dt = WorkSheet.ToDataTable();
        }
    }
}