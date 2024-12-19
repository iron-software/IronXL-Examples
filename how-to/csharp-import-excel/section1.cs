using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section1
    {
        public static void Run()
        {
            //load Excel file
            WorkBook wb = WorkBook.Load("Path");
        }
    }
}