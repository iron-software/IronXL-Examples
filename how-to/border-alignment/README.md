# Configuring Cell Border and Text Alignment Using IronXL

In Excel, enhancing the appearance of cells with customized borders and fixed text alignment greatly improves the visual appeal and readability of your spreadsheets. With the IronXL library, setting border styles, thickness, and colors, as well as aligning text, can dramatically enhance data presentation and visibility.

## Example of Setting Cell Borders and Alignment

You can modify the visual layout of cells in your Excel document by adjusting properties such as **TopBorder**, **RightBorder**, **BottomBorder**, and **LeftBorder**. Utilize the styles available in the `IronXL.Styles.BorderType` enumeration to personalize cell borders. Further details on available border styles can be found by navigating to our [border styles list](https://ironsoftware.com/csharp/excel/how-to/select-range/).

For precise control over text placement, modify the **HorizontalAlignment** and **VerticalAlignment** settings using the respective properties in the `Style` class, choosing from the options provided in the `IronXL.Styles.HorizontalAlignment` and `IronXL.Styles.VerticalAlignment` enumerations. Explore the full range of alignment options to ensure your data is displayed perfectly.

```cs
using IronXL;
using IronXL.Styles;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

workSheet["B2"].Value = "Sample Text";

// Setting borders
workSheet["B2"].Style.LeftBorder.Type = BorderType.MediumDashed;
workSheet["B2"].Style.RightBorder.Type = BorderType.MediumDashed;

// Adjusting text alignment
workSheet["B2"].Style.HorizontalAlignment = HorizontalAlignment.Center;

workBook.SaveAs("EnhancedSpreadsheet.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/set-border-alignment.webp" alt="Border And Alignment Example" class="img-responsive add-shadow">
    </div>
</div>

## Advanced Customization: Border Color and Line Patterns

### Customize Border Color

The default border color is typically black, but IronXL allows for customization to any desired color. You can apply specific colors using the `Color` class or Hex color codes with the `SetColor` method. Note that you must set a border type to see the color applied.

```cs
using IronXL;
using IronXL.Styles;
using IronSoftware.Drawing;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

workSheet["B2"].Style.LeftBorder.Type = BorderType.Thick;
workSheet["B2"].Style.RightBorder.Type = BorderType.Thick;

// Customizing border colors
workSheet["B2"].Style.LeftBorder.SetColor(Color.Aquamarine);
workSheet["B2"].Style.RightBorder.SetColor("#FF7F50");

workBook.SaveAs("CustomBorderColors.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/set-border-color.webp" alt="Custom Border Colors" class="img-responsive add-shadow">
    </div>
</div>

### Border Line Positions and Patterns

IronXL supports various border positions and line styles, including top, right, bottom, left, and diagonal borders. This versatility enables sophisticated formatting and custom designs.

```cs
using IronXL;
using IronXL.Styles;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

workSheet["B2"].StringValue = "Top Border Example";
workSheet["B4"].StringValue = "Diagonal Border Example";

// Applying top border line
workSheet["B2"].Style.TopBorder.Type = BorderType.Thick;

// Applying a diagonal border line and setting its direction
workSheet["B4"].Style.DiagonalBorder.Type = BorderType.Thick;
workSheet["B4"].Style.DiagonalBorderDirection = DiagonalBorderDirection.Forward;

workBook.SaveAs("BorderStyles.xlsx");
```

#### Visual Reference:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/border-types.webp" alt="Border Styles Reference" class="img-responsive add-shadow">
    </div>
</div>

### Alignment Types

IronXL offers a comprehensive set of horizontal and vertical alignment types to cater to various design needs. Whether it's centering text, justifying content, or distributing words evenly, these options ensure that your cell content is precisely positioned for optimal readability and aesthetic balance.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/alignment-types.webp" alt="Text Alignment Options" class="img-responsive add-shadow">
    </div>
</div>

By utilizing these sophisticated functions, you can easily create professional, visually pleasing worksheets that clearly depict your data, making information both easy to read and eye-catching.