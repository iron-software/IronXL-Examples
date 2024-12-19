using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section23
    {
        public static void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            int i = 1;
            foreach (var cell in workSheet["f1:f4"])
            {
                cell.Formula = "=trim(D" + i + ")";
                i++;
            }
            workBook.SaveAs("editedFile.xlsx");
        }
    }
}