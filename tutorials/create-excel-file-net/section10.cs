using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet
{
    public static class Section10
    {
        public static void Run()
        {
            workSheet.ProtectSheet("Password");
            workSheet.CreateFreezePane(0, 1);
        }
    }
}