# How to Configure Cell Border and Text Alignment

***Based on <https://ironsoftware.com/how-to/border-alignment/>***


Cell borders in Excel are the distinct lines that can be added around cells or groups of cells. Text alignment, conversely, determines the placement of text within a cell both on the vertical and horizontal axes.

Utilizing IronXL, you can elevate data visualization, augment readability, and forge spreadsheets that look and feel professional by personalizing border styles, thickness, colors, and text alignments for enhanced data representation.

<h3>Getting Started with IronXL</h3>

----------------------------------

## Example: Setting Cell Border and Alignment

Modify the visual style of a [specific cell, column, row, or block of cells](https://ironsoftware.com/csharp/excel/how-to/select-range/) by applying borders through the **TopBorder**, **RightBorder**, **BottomBorder**, and **LeftBorder** properties. IronXL offers various border styles through the **IronXL.Styles.BorderType** enumeration. Explore [all possible border styles](#anchor-available-border-type) to select the ideal one for your needs.

For precise control over text positioning, tweak the **HorizontalAlignment** and **VerticalAlignment** properties in the Style object according to your requirements. Use the enumerations **IronXL.Styles.HorizontalAlignment** and **IronXL.Styles.VerticalAlignment** for setting desired text orientations. Discover [all alignment options](#anchor-available-border-type) to perfectly present your data.

```cs
using IronXL;
using IronXL.Styles;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

workSheet["B2"].Value = "Cell B2";

// Apply cell border
workSheet["B2"].Style.LeftBorder.Type = BorderType.MediumDashed;
workSheet["B2"].Style.RightBorder.Type = BorderType.MediumDashed;

// Adjust text alignment
workSheet["B2"].Style.HorizontalAlignment = HorizontalAlignment.Center;

workBook.SaveAs("customizedBorderAndAlignment.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/set-border-alignment.webp" alt="Border And Alignment Configuration" class="img-responsive add-shadow">
    </div>
</div>

## Advanced Cell Border and Alignment Customization

### Border Color Customization

While the default color for borders is black, IronXL enables customization to any preferable color using the **Color** class or Hex color codes. Set the border color via the **Color** property by specifying your desired color or Hex code. Remember, the border color remains invisible until a border type is specified.

```cs
using IronXL;
using IronXL.Styles;
using IronSoftware.Drawing;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

workSheet["B2"].Style.LeftBorder.Type = BorderType.Thick;
workSheet["B2"].Style.RightBorder.Type = BorderType.Thick;

// Customize border color
workSheet["B2"].Style.LeftBorder.SetColor(Color.Aquamarine);
workSheet["B2"].Style.RightBorder.SetColor("#FF7F50");

workBook.SaveAs("customBorderColors.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/set-border-color.webp" alt="Custom Border Color" class="img-responsive add-shadow">
    </div>
</div>

### Border Lines and Patterns

IronXL supports six border line positions with multiple patterns or types including top, right, bottom, left, diagonal forward, diagonal backward, and diagonal in both directions.

```cs
using IronXL;
using IronXL.Styles;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

workSheet["B2"].StringValue = "Set Top Border";
workSheet["B4"].StringValue = "Set Diagonal Forward Border";

// Apply top border
workSheet["B2"].Style.TopBorder.Type = BorderType.Thick;

// Apply diagonal border and set direction
workSheet["B4"].Style.DiagonalBorder.Type = BorderType.Thick;
workSheet["B4"].Style.DiagonalBorderDirection = DiagonalBorderDirection.Forward;

workBook.SaveAs("differentBorderTypes.xlsx");
```

#### Displaying Border Lines

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/border-line.webp" alt="Border Line Types" class="img-responsive add-shadow">
    </div>
</div>

#### Showcasing Border Patterns

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/border-types.webp" alt="Border Patterns Example" class="img-responsive add-shadow">
    </div>
</div>

### Illustration of Alignment Types Offered by IronXL

Get to know the extensive alignment functionalities provided by IronXL through this graphical representation:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/alignment-types.webp" alt="Alignment Types Example" class="img-responsive add-shadow">
    </div>
</div>