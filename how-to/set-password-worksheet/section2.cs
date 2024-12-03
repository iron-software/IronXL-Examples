using IronXL.Excel;
namespace ironxl.SetPasswordWorksheet
{
    public class Section2
    {
        public void Run()
        {
            // Remove protection for selected worksheet. It works without password!
            workSheet.UnprotectSheet();
        }
    }
}