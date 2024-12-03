using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section10
    {
        public void Run()
        {
            workSheet.ProtectSheet("Password");
            workSheet.CreateFreezePane(0, 1);
        }
    }
}