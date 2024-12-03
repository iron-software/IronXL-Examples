# Setting Cell Background Patterns and Colors in Excel

***Based on <https://ironsoftware.com/how-to/background-pattern-color/>***


In Excel, defining background patterns involves applying a decorative or textured fill to the background of a cell. Moreover, setting a background color refers to filling a cell's background with a uniform color. 

IronXL allows you to harness these attributes effectively to enhance the visual impact of your Excel spreadsheets. This feature is particularly useful for improving data visualization and emphasizing crucial data points.

## Example: Setting Cell Background Pattern and Color

To apply a background pattern to a [selected cell, column, row, or range](https://ironsoftware.com/csharp/excel/how-to/select-range/), utilize the `FillPattern` property from the `IronXL.Styles.FillPattern` enum. Subsequently, set the background color by using the `SetBackgroundColor` method or by assigning a value to the `BackgroundColor` property. Colors can be specified using the `Color` class or by entering a HEX color code, such as SeaGreen represented by "#FFF5EE".

Currently, it is not possible to alter the color of the fill pattern itself.

```cs
using IronSoftware.Drawing;
using IronXL.Excel;
namespace ironxl.BackgroundPatternColor
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workbook = WorkBook.Create();
            WorkSheet worksheet = workbook.DefaultWorkSheet;
            
            // Apply background pattern
            worksheet["A1"].Style.FillPattern = FillPattern.AltBars;
            worksheet["A2"].Style.FillPattern = FillPattern.ThickVerticalBands;
            
            // Apply background color
            worksheet["A1"].Style.SetBackgroundColor(Color.Aquamarine);
            worksheet["A2"].Style.BackgroundColor = "#ADFF2F";
            
            workbook.SaveAs("setBackgroundPattern.xlsx");
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/background-pattern-color/background-pattern-color.png" alt="Output" class="img-responsive add-shadow">
    </div>
</div>

## Available Fill Patterns

Explore the diverse range of fill patterns provided by the `IronXL.Styles.FillPattern` enum. These patterns can bring distinct visual styles to your Excel worksheets. Below is a visual guide to all the fill patterns available in IronXL:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/background-pattern-color/fill-pattern.png" alt="Available Fill Pattern" class="img-responsive add-shadow">
    </div>
</div