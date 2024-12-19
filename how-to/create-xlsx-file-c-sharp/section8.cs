using IronXL.Excel;
namespace IronXL.Examples.HowTo.CreateXlsxFileCSharp
{
    public static class Section8
    {
        public static void Run()
        {
            /**
            Set Metadata
            anchor-set-metadata-for-excel-files
            **/
            WorkBook wb = WorkBook.Create();
            wb.Metadata.Author = "AuthorName";
            wb.Metadata.Title="TitleValue";
        }
    }
}