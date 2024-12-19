using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadExcelFileExample
{
    public static class Section1
    {
        public static void Run()
        {
            //Load Excel file
            WorkBook wb = WorkBook.Load("Path");
        }
    }
}