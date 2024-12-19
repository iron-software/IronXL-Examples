using IronXL.Excel;
namespace IronXL.Examples.HowTo.SetPasswordWorksheet
{
    public static class Section2
    {
        public static void Run()
        {
            // Remove protection for selected worksheet. It works without password!
            workSheet.UnprotectSheet();
        }
    }
}