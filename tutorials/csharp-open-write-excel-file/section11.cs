using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section11
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            workBook.Metadata.Title = "IronXL New File";
            
            WorkSheet workSheet = workBook.CreateWorkSheet("1stWorkSheet");
            workSheet["A1"].Value = "Hello World";
            workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
            
            workBook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
        }
    }
}