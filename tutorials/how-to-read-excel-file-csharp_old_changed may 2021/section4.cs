using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp_Old_Changed May 2021
{
    public static class Section4
    {
        public static void Run()
        {
            /**
            Create WorkBook
            anchor-create-a-workbook
            **/
            var workbook = new WorkBook(ExcelFileFormat.XLSX);
        }
    }
}