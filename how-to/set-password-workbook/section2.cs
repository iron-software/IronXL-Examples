using IronXL.Excel;
namespace IronXL.Examples.HowTo.SetPasswordWorkbook
{
    public static class Section2
    {
        public static void Run()
        {
            // Remove protection for opened workbook. Original password is required.
            workBook.Password = null;
        }
    }
}