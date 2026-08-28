using IronXL.Printing;
using IronXL;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet
{
    public static class Section11
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            workSheet.SetPrintArea("A1:L12");
            workSheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;
            workSheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;
        }
    }
}