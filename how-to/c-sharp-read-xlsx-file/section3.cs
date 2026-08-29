using IronXL;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section3
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook wb = IronXL.WorkBook.Load("sample.xlsx");

            /**
            Access Sheet by Index
            anchor-access-specific-worksheet
            **/
            WorkSheet ws = wb.WorkSheets [0]; //by sheet index
        }
    }
}