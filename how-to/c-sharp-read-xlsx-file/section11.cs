using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section11
    {
        public static void Run()
        {
            DataTable dt=WorkSheet.ToDataTable(True);
        }
    }
}