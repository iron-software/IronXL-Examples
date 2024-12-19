using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section13
    {
        public static void Run()
        {
            /**
            WorkSheet Cell Values
            anchor-read-excel-file-as-dataset
            **/
            using IronXL;
            using System.Data; 
            static void Main(string [] args)
            { 
            WorkBook wb = WorkBook.Load("sample.xlsx");
            DataSet ds = wb.ToDataSet();//behave complete Excel file as DataSet
            foreach (DataTable dt in ds.Tables)//behave Excel WorkSheet as DataTable. 
            {
                foreach (DataRow row in dt.Rows)//corresponding Sheet's Rows
                {
                    for (int i = 0; i < dt.Columns.Count; i++)//Sheet columns of corresponding row
                    {
                        Console.Write(row [i]);
                    }
                }
            }
            }
        }
    }
}