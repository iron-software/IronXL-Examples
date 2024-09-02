# How to Modify Text Style and Size in Excel Cells

Adjusting text attributes like font type, size, color, underline, bold, italic, script, and strikeout enhances document formatting significantly. These features give you the flexibility to highlight important sections of the content, enhance readability, and design an attractive document layout. Using IronXL, you can modify these font properties directly in C# .NET, bypassing the need for COM interop, thus streamlining your process for producing refined and professional documents.

## Example: Modifying Font Style and Size

To customize the font of a [cell, column, or range](https://www.ironsoftware.com/csharp/excel/how-to/select-range/), alter the **Font** attributes within the **Style** object. Define the font family with the **Name** attribute, adjust the size with the **Height** attribute, and bold the text using the **Bold** attribute. For added emphasis, the **Underline** attribute can be applied.

```cs
using IronXL;
using IronXL.Styles;

WorkBook workbook = WorkBook.Create();
WorkSheet worksheet = workbook.DefaultWorkSheet;

worksheet["B2"].StringValue = "Text Attributes";

// Applying font family
worksheet["B2"].Style.Font.Name = "Arial";

// Modifying font size
worksheet["B2"].Style.Font.Height = 16;

// Making text bold
worksheet["B2"].Style.Font.Bold = true;

// Adding underline
worksheet["B2"].Style.Font.Underline = FontUnderlineType.Single;

workbook.SaveAs("customTextStyles.xlsx");
```

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://www.ironsoftware.com/static-assets/excel/how-to/cell-font-size/set-font-and-size.webp" alt="Custom Font and Size" class="img-responsive add-shadow">
    </div>
</div>

## Advanced Example: Comprehensive Font Customization

For more refined customization, IronXL supports additional font styling options, including Italic, Strikeout, FontScript for superscripts and subscripts, and specific font colors. Here is an example demonstrating these extensive font styling capabilities.

The **Name** property should match the exact font name as formatted, including spaces and capitalization.

```cs
using IronXL;
using IronXL.Styles;

WorkBook workbook = WorkBook.Create();
WorkSheet worksheet = workbook.DefaultWorkSheet;

worksheet["B2"].StringValue = "Advanced Text Styles";

// Choosing font style
worksheet["B2"].Style.Font.Name = "Calibri";

// Adjusting script type
worksheet["B2"].Style.Font.FontScript = FontScript.None;

// Applying double underline
worksheet["B2"].Style.Font.Underline = FontUnderlineType.Double;

// Setting bold text
worksheet["B2"].Style.Font.Bold = true;

// Turning off italic style
worksheet["B2"].Style.Font.Italic = false;

// Disabling strikeout
worksheet["B2"].Style.Font.Strikeout = false;

// Customizing font color
worksheet["B2"].Style.Font.Color = "#FF6347";

workbook.SaveAs("advancedTextStyles.xlsx");
```

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://www.ironsoftware.com/static-assets/excel/how-to/cell-font-size/set-font-and-size-advanced.webp" alt="Advanced Font Styling" class="img-responsive add-shadow">
    </div>
</div>

### Underline Options

Excel offers various underline styles suitable for enhancing text formatting. The Accounting underline, for example, offers additional space between the characters and the outline, especially extending beyond the text for entries but limited to the numeric values in [data formatted cells](https://www.ironsoftware.com/csharp/excel/how-to/set-cell-data-format/).

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://www.ironsoftware.com/static-assets/excel/how-to/cell-font-size/underline.webp" alt="Underline Varieties" class="img-responsive add-shadow">
    </div>
</div>

### Font Script Styling

IronXL font script provides three styling options:
- **none**: Default, aligns text along the baseline.
- **super**: Elevates text above the baseline, useful for notation like exponents.
- **sub**: Lowers text beneath the baseline, ideal for chemical formulas or mathematical expressions.

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://www.ironsoftware.com/static-assets/excel/how-to/cell-font-size/font-script.webp" alt="Script Options" class="img-responsive add-shadow">
    </div>
</div>

### Defining Font Color

The color of the font can be set using either the `Color` property or using the `SetColor` method, which supports inputs as IronSoftware.Drawing.Color or a Hex color code.

```cs
using IronXL;
using IronSoftware.Drawing;

WorkBook workbook = WorkBook.Create();
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Setting font color directly
worksheet["B2"].Style.Font.Color = "#FF6347";

// Setting font color using Hex code
worksheet["B2"].Style.Font.SetColor("#FF6347");

// Applying color through IronSoftware.Drawing
worksheet["B2"].Style.Font.SetColor(Color.Blue);
```