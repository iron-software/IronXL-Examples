using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section19
    {
        public static void Run()
        {
            /**
            Add New Sheet
            anchor-add-new-sheet-to-a-workbook
            **/
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            var newSheet = workbook.CreateWorkSheet("new_sheet");
            newSheet ["A1"].Value = "Hello World";
            workbook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx");
        }
    }
}