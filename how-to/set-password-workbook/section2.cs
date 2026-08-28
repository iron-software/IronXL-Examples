using IronXL;
namespace IronXL.Examples.HowTo.SetPasswordWorkbook
{
    public static class Section2
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");

            // Remove protection for opened workbook. Original password is required.
            workBook.Password = null;
        }
    }
}