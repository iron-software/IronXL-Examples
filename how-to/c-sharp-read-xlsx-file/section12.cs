using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section12
    {
        public static void Run()
        {
            DataSet ds = WorkBook.ToDataSet();
        }
    }
}