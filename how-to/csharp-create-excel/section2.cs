using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpCreateExcel
{
    public static class Section2
    {
        public static void Run()
        {
            /**
            Create Csharp WorkBook 
            anchor-c-num-create-excel-workbook
            **/
            //for creating .xlsx extension file
            WorkBook wb = WorkBook.Create(ExcelFileFormat.XLSX);
            //for creating .xls extension file
            WorkBook wb = WorkBook.Create(ExcelFileFormat.XLS);
        }
    }
}