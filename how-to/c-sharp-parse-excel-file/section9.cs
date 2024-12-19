using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section9
    {
        public static void Run()
        {
            DataSet ds = WorkBook.ToDataSet();
        }
    }
}