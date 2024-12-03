# Mastering Worksheet Management

***Based on <https://ironsoftware.com/how-to/manage-worksheet/>***


The **IronXL** library offers a streamlined approach to managing worksheets through C# coding. Leveraging IronXL negates the necessity for Office Interop, empowering you with capabilities such as sheet creation, deletion, repositioning, and activation within an Excel workbook.

## Examples of Worksheet Management

IronXL equips you with straightforward commands for creating, moving, and removing worksheets efficiently.

All index references utilize zero-based indexing for clarity.

## Create Worksheet

Utilize the `CreateWorkSheet` method for instantiating a new sheet, which only requires a unique sheet name as an argument. The method also conveniently returns the worksheet instance for immediate further actions such as [merging cells](https://ironsoftware.com/csharp/excel/how-to/csharp-excel-merge-cells/).

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.ManageWorksheet
{
    public class Section1
    {
        public void Run()
        {
            // Initialize a new Excel workbook
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            
            // Add multiple worksheets
            WorkSheet workSheet1 = workBook.CreateWorkSheet("workSheet1");
            WorkSheet workSheet2 = workBook.CreateWorkSheet("workSheet2");
            WorkSheet workSheet3 = workBook.CreateWorkSheet("workSheet3");
            WorkSheet workSheet4 = workBook.CreateWorkSheet("workSheet4");
            
            // Save the workbook with new sheets
            workBook.SaveAs("createNewWorkSheets.xlsx");
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-create-worksheet.png" alt="Create Worksheets" class="img-responsive add-shadow">
    </div>
</div>

<hr>

## Adjust Worksheet Position

Reordering a sheet within the workbook is achieved using the `SetSheetPosition` method. It requires the name of the worksheet and its new zero-based index.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.ManageWorksheet
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");
            
            // Rearrange the position of a worksheet
            workBook.SetSheetPosition("workSheet2", 0);
            
            // Store changes
            workBook.SaveAs("setWorksheetPosition.xlsx");
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-set-sheet-position.png" alt="Change Worksheet Position" class="img-responsive add-shadow">
    </div>
</div>

<hr>

## Activate a Worksheet

Designate an active worksheet so that it opens by default when the workbook is accessed. This is done using the `SetActiveTab` method, specifying the sheet's index.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.ManageWorksheet
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");
            
            // Make workSheet3 the default active sheet
            workBook.SetActiveTab(2);
            
            // Save the configuration
            workBook.SaveAs("setActiveTab.xlsx");
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-set-active-tab.png" alt="Set Active Worksheet" class="img-responsive add-shadow">
    </div>
</div>

<hr>

## Remove Worksheet

To remove a worksheet, employ the `RemoveWorksheet` method by specifying either the sheet's index or its name.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.ManageWorksheet
{
    public class Section4
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");
            
            // Delete workSheet1 by index
            workBook.RemoveWorkSheet(1);
            
            // Delete workSheet2 by name
            workBook.RemoveWorkSheet("workSheet2");
            
            // Persist changes
            workBook.SaveAs("removeWorksheet.xlsx");
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-remove-worksheet.png" alt="Remove Worksheet" class="img-responsive add-shadow">
    </div>
</div>

## Duplicate Worksheet

You can clone a worksheet either within the same workbook or to another workbook by utilizing the `CopySheet` method for intra-workbook duplication or `CopyTo` for inter-workbook operations.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.ManageWorksheet
{
    public class Section5
    {
        public void Run()
        {
            WorkBook firstBook = WorkBook.Create(ExcelFileFormat.XLSX);
            WorkBook secondBook = WorkBook.Create();
            
            // Grab the initial sheet
            WorkSheet workSheet = firstBook.DefaultWorkSheet;
            
            // Clone within the same workbook
            workSheet.CopySheet("Copied Sheet");
            
            // Clone to a different workbook
            workSheet.CopyTo(secondBook, "Copied Sheet");
            
            // Save both workbooks
            firstBook.SaveAs("firstWorksheet.xlsx");
            secondBook.SaveAs("secondWorksheet.xlsx");
        }
    }
}
```

<div class="competitors-section__wrapper-even-1">
    <div class="competitors__card" style="width: 54%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-copy-worksheet-first.png" alt="First Worksheet" class="img-responsive add-shadow" style="margin-bottom: 20px;"/>
        <p class="competitors__download-link" style="color: #181818; font-style: italic; margin-bottom: 30px;">firstWorksheet.xlsx</p>
    </div>
    <div class="competitors__card" style="width: 44%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-copy-worksheet-second.png" alt="Second Worksheet" class="img-responsive add-shadow" style="margin-bottom: 20px;"/>
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">secondWorksheet.xlsx</p>
    </div>
</div>