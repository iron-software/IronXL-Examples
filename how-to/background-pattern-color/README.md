# How to Configure Cell Background Pattern & Color in Excel

***Based on <https://ironsoftware.com/how-to/background-pattern-color/>***


In Excel, the term "background cell pattern" refers to the texture or visual pattern added to the background of a cell. Similarly, "background cell color" pertains to the flat, uniform color that fills a cell's background.

By utilizing both these attributes together, users are able to craft visually engaging backgrounds incorporating diverse patterns, colors, and textures. IronXL provides the capabilities to embellish Excel cell backgrounds, thus enhancing data visibility and emphasizing key data points within your spreadsheets.

***

***

<h3>Getting Started with IronXL</h3>

----------------------------------

## Example: Configuring Cell Background Patterns & Colors

To apply a background pattern to a [selected cell, column, row, or range](https://ironsoftware.com/csharp/excel/how-to/select-range/), adjust the **FillPattern** property using values from **IronXL.Styles.FillPattern** enum. Next, set the background color using the `SetBackgroundColor` method or the **BackgroundColor** property. You can select a predefined color from the **Color** class or use a Hex color code, such as "#FFF5EE" for SeaGreen.

Currently, altering the color of the fill pattern is not supported.

```cs
using IronXL;
using IronXL.Styles;
using IronSoftware.Drawing;

WorkBook workbook = WorkBook.Create();
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Applying background pattern
worksheet["A1"].Style.FillPattern = FillPattern.AltBars;
worksheet["A2"].Style.FillPattern = FillPattern.ThickVerticalBands;

// Applying background color
worksheet["A1"].Style.SetBackgroundColor(Color.Aquamarine);
worksheet["A2"].Style.BackgroundColor = "#ADFF2F";

workbook.SaveAs("customBackgroundPattern.xlsx");
```

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/background-pattern-color/background-pattern-color.png" alt="Output" class="img-responsive add-shadow">
    </div>
</div>

## Fill Patterns Available

Leverage the distinct fill patterns obtainable via the **IronXL.Styles.FillPattern** enum to customize your Excel documents. Below is a portrayal of all the fill patterns that IronXL offers:

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/background-pattern-color/fill-pattern.png" alt="Available Fill Pattern" class="img-responsive add-shadow">
    </div>
</div