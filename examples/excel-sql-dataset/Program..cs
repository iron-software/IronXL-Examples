using IronXL;
using System;
using System.Data;

// Supported for XLSX, XLS, XLSM, XLTX, CSV and TSV
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Convert the whole Excel WorkBook to a DataSet
DataSet dataSet = workBook.ToDataSet();

foreach (DataTable table in dataSet.Tables)
{
    Console.WriteLine(table.TableName);

    // Enumerate by rows or columns first at your preference
    foreach (DataRow row in table.Rows)
    {
        for (int i = 0 ; i < table.Columns.Count ; i++)
        {
            Console.Write(row[i]);
        }
    }
}
