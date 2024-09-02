using IronXL;
using IronXL.Styles;
using IronSoftware.Drawing;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

workSheet["B2"].Style.LeftBorder.Type = BorderType.Thick;
workSheet["B2"].Style.RightBorder.Type = BorderType.Thick;

// Set cell border color
workSheet["B2"].Style.LeftBorder.SetColor(Color.Aquamarine);
workSheet["B2"].Style.RightBorder.SetColor("#FF7F50");

workBook.SaveAs("setBorderColor.xlsx");