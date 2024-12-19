using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet
{
    public static class Section11
    {
        public static void Run()
        {
            workSheet.SetPrintArea("A1:L12");
            workSheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;
            workSheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;
        }
    }
}