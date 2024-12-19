using IronXL.Excel;
namespace IronXL.Examples.HowTo.CreateXlsxFileCSharp
{
    public static class Section1
    {
        public static void Run()
        {
            /**
            Create XLSX File
            anchor-create-a-workbook
            **/
            WorkBook wb = WorkBook.Create();
        }
    }
}