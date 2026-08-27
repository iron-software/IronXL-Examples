using System.Data;
using IronXL;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section13
    {
        public static void Run()
        {
            // This snippet is a member of a larger component from the accompanying README, not a standalone program.
            // Kept verbatim; see README.md for the full context.
            // /**
            // WorkSheet Cell Values
            // anchor-read-excel-file-as-dataset
            // **/
            // static void Main(string [] args)
            // {
            // WorkBook wb = WorkBook.Load("sample.xlsx");
            // DataSet ds = wb.ToDataSet();//behave complete Excel file as DataSet
            // foreach (DataTable dt in ds.Tables)//behave Excel WorkSheet as DataTable.
            // {
            // foreach (DataRow row in dt.Rows)//corresponding Sheet's Rows
            // {
            // for (int i = 0; i < dt.Columns.Count; i++)//Sheet columns of corresponding row
            // {
            // Console.Write(row [i]);
            // }
            // }
            // }
            // }
        }
    }
}