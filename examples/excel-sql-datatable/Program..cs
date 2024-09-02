using IronXL;
using System;
using System.Data;

// Supported for XLSX, XLS, XLSM, XLTX, CSV and TSV
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Select default sheet
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Convert the worksheet to DataTable
DataTable dataTable = workSheet.ToDataTable(true);

// Enumerate by rows or columns first at your preference
foreach (DataRow row in dataTable.Rows)
{
    for (int i = 0 ; i < dataTable.Columns.Count ; i++)
    {
        Console.Write(row[i]);
    }
}
