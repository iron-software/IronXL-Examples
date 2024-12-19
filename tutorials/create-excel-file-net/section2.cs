using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
        }
    }
}