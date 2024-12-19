using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet_Old_Changed May 2021
{
    public static class Section11
    {
        public static void Run()
        {
            /**
            Page & Print Properties
            anchor-set-page-and-print-properties
            **/
            sheet.SetPrintArea("A1:L12");
            sheet.PrintSetup.PrintOrientation = IronXL.Printing.PrintOrientation.Landscape;
            sheet.PrintSetup.PaperSize = IronXL.Printing.PaperSize.A4;
        }
    }
}