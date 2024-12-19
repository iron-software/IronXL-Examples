using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet_Old_Changed May 2021
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
        }
    }
}