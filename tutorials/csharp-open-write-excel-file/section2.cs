using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section2
    {
        public static void Run()
        {
            workSheet["B1"].Value = 11.54;
            
            // Save Changes
            workBook.SaveAs("test.xlsx");
        }
    }
}