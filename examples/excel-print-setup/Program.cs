using IronXL;
using IronXL.Printing;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Set the print header and footer of the worksheet
workSheet.Header.Center = "My document";
workSheet.Footer.Center = "Page &P of &N";

// Set the header margin
workSheet.PrintSetup.HeaderMargin = 2.33;

// Set the size of the paper
// Paper size enum represents different sizes of paper
workSheet.PrintSetup.PaperSize = PaperSize.B4;

// Set the print orientation of the worksheet
workSheet.PrintSetup.PrintOrientation = PrintOrientation.Portrait;

// Set black and white printing
workSheet.PrintSetup.NoColor = true;

workBook.SaveAs("PrintSetup.xlsx");
