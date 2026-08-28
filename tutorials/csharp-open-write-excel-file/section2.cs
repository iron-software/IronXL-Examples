using IronXL;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section2
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            workSheet["B1"].Value = 11.54;
            
            // Save Changes
            workBook.SaveAs("test.xlsx");
        }
    }
}