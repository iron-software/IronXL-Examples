using IronXL.Excel;
namespace ironxl.SetPasswordWorkbook
{
    public class Section2
    {
        public void Run()
        {
            // Remove protection for opened workbook. Original password is required.
            workBook.Password = null;
        }
    }
}