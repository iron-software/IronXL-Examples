using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section2
    {
        public void Run()
        {
            workSheet["B1"].Value = 11.54;
            
            // Save Changes
            workBook.SaveAs("test.xlsx");
        }
    }
}