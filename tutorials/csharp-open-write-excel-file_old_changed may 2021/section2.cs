using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section2
    {
        public static void Run()
        {
            sheet ["B1"].Value = 11.54;
            
            //Save Changes
            workbook.SaveAs("test.xlsx");
        }
    }
}