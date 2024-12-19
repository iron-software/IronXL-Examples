using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet_Old_Changed May 2021
{
    public static class Section10
    {
        public static void Run()
        {
            /**
            Set Worksheet Properties
            anchor-set-worksheet-and-print-properties
            **/
            sheet.ProtectSheet("Password");
            sheet.CreateFreezePane(0, 1);
        }
    }
}