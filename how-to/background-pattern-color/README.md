# How to Set Cell Background Pattern & Color

In Excel, defining a background cell pattern involves adding a specific fill design to the back of a cell. Additionally, setting a background cell color involves selecting a solid shade that covers the cell's background.

By incorporating these functionalities, users have the ability to design enticing cell backgrounds using an array of patterns, hues, and textures. With IronXL, you gain the capability to customize these background aspects in your Excel sheets, which serves to enhance data presentation and draw attention to crucial data points.

***

***

## Example: Setting Cell Background Pattern & Color

To configure a background pattern for a [specific cell, column, row, or range](https://ironsoftware.com/csharp/excel/how-to/select-range/), you should utilize the **FillPattern** property available within the **IronXL.Styles.FillPattern** enumeration. Subsequently, either use the `SetBackgroundColor` method or alter the **BackgroundColor** property to apply your desired color. The **Color** class provides various options, or you can directly use a Hex color code, such as SeaGreen represented by "#FFF5EE".

Currently, it is not possible to modify the color of the fill pattern itself.

```cs
using IronXL;
using IronXL.Styles;
using IronSoftware.Drawing;

// Create a new workbook
WorkBook workbook = WorkBook.Create();

// Access the default worksheet
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Apply a background pattern
worksheet["A1"].Style.FillPattern = FillPattern.AltBars;
worksheet["A2"].Style.FillPattern = FillPattern.ThickVerticalBands;

// Apply a background color
worksheet["A1"].Style.SetBackgroundColor(Color.Aquamarine);
worksheet["A2"].Style.BackgroundColor = "#ADFF2F";

// Save the workbook
workbook.SaveAs("CustomBackgroundExcel.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/background-pattern-color/background-pattern-color.png" alt="Output" class="img-responsive add-shadow">
    </div>
</div>

## Available Fill Patterns

Explore the diverse range of fill patterns provided by the **IronXL.Styles.FillPattern** enum to personalize the fill pattern in your Excel documents. Below is a display showcasing all the fill patterns available through IronXL:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/background-pattern-color/fill-pattern.png" alt="Available Fill Pattern" class="img-responsive add-shadow">
    </div>
</div