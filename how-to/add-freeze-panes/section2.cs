using IronXL.Excel;
namespace ironxl.AddFreezePanes
{
    public class Section2
    {
        public void Run()
        {
            // Remove all existing freeze or split pane
            workSheet.RemovePane();
        }
    }
}