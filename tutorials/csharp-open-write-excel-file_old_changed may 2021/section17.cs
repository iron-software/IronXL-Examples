using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section17
    {
        public static void Run()
        {
            /**
            Function Trim
            anchor-trim-example
            **/
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
            var sheet = workbook.WorkSheets.First();
            int i = 1;
            foreach (var cell in sheet ["f1:f4"])
            {
                cell.Formula = "=trim(D"+i+")";
                i++;
            }
            workbook.SaveAs("editedFile.xlsx");
        }
    }
}