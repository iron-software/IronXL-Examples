using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section2
    {
        public static void Run()
        {
            //specify sheet name of Excel WorkBook
            WorkSheet ws = wb.GetWorkSheet("SheetName");
        }
    }
}