using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section25
    {
        public void Run()
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