using IronXL;

WorkBook workBook = WorkBook.Load("insertImages.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Remove image
workSheet.RemoveImage(3);

workBook.SaveAs("removeImage.xlsx");