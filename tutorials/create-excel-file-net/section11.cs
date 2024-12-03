using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section11
    {
        public void Run()
        {
            workSheet.SetPrintArea("A1:L12");
            workSheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;
            workSheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;
        }
    }
}