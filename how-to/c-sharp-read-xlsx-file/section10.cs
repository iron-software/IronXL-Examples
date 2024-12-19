using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section10
    {
        public static void Run()
        {
            DataTable dt=WorkSheet.ToDataTable();
        }
    }
}