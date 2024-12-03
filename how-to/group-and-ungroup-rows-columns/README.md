# How to Group and Ungroup Rows & Columns in Excel

***Based on <https://ironsoftware.com/how-to/group-and-ungroup-rows-columns/>***


## Introduction

The grouping feature in Excel allows for better organization of data by enabling users to create collapsible sections across rows or columns. This functionality is particularly useful for managing and analyzing extensive datasets. On the flip side, the ungrouping feature reverses these groups to their initial layout, facilitating straightforward data manipulation and detailed inspection of particular areas of the spreadsheet.

IronXL supports the ability to programmatically group and ungroup rows and columns in C# .NET environments, eliminating the need for Interop.

## Group & Ungroup Rows Example

Remember that all indexes specified are zero-based.

Grouping and ungrouping operations are limited to cells that are not empty.

### Group Rows

To create row groups, use the `GroupRows` method, specifying the start and end index positions. You can execute multiple groupings on the same or different rows by repeating this method.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.GroupAndUngroupRowsColumns
{
    public class Section1
    {
        public void Run()
        {
            // Load the existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Group rows from index 0 to 7 (equivalent to Excel rows 1-8)
            workSheet.GroupRows(0, 7);
            
            // Save the new spreadsheet
            workBook.SaveAs("groupRow.xlsx");
        }
    }
}
```

#### Output

![Group Rows](https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-group-row.png)

### Ungroup Rows

To ungroup previously grouped rows, use the `UngroupRows` method. This method provides the flexibility to segment a group of rows into smaller subgroups, although these won't be recognized as separate groupings. For instance, ungrouping rows 3 to 5 from a group consisting of rows 1 to 9 results in two ungrouped sections: rows 1-2 and 6-9.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.GroupAndUngroupRowsColumns
{
    public class Section2
    {
        public void Run()
        {
            // Load the existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Ungroup rows 3 to 5
            workSheet.UngroupRows(2, 4);
            
            // Save the spreadsheet after ungrouping
            workBook.SaveAs("ungroupRow.xlsx");
        }
    }
}
```

#### Output

![Before Ungrouping Rows](https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-group-row.png)
![After Ungrouping Rows](https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-ungroup-row.png)

## Group & Ungroup Columns Example

### Group Columns

Grouping columns operates similarly to rows. Utilize the `GroupColumns` method by specifying index numbers or alphabetical identifiers for columns. You can create multiple column groups by calling this method repeatedly.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.GroupAndUngroupRowsColumns
{
    public class Section3
    {
        public void Run()
        {
            // Load the existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Group columns from A to F
            workSheet.GroupColumns(0, 5);
            
            // Save the spreadsheet with grouped columns
            workBook.SaveAs("groupColumn.xlsx");
        }
    }
}
```

#### Output

![Group Columns](https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-group-column.png)

### Ungroup Columns

To ungroup columns, you can use `UngroupColumn` for individual column identifiers or `UngroupColumns` for column indexes. This process effectively splits a group into two. For example, ungrouping columns C to D from a group of A to F results in columns A-B and E-F left as separate groups.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.GroupAndUngroupRowsColumns
{
    public class Section4
    {
        public void Run()
        {
            // Load the existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Ungroup columns between C and D
            workSheet.UngroupColumn("C", "D");
            
            // Save the modified spreadsheet
            workBook.SaveAs("ungroupColumn.xlsx");
        }
    }
}
```

#### Output

![Before Ungrouping Columns](https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-group-column.png)
![After Ungrouping Columns](https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-ungroup-column.png)