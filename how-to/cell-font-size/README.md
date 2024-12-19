# Customizing Cell Font and Size in Excel

***Based on <https://ironsoftware.com/how-to/cell-font-size/>***


Customizing the font attributes such as the font type, size, color, underline, bolding, italicizing, subscripting, and strike-through enhances the readability, emphasizes key details, and augments the visual aesthetics of your documents. IronXL allows you to modify these font properties easily in your C# .NET applications, streamlining the process to help you produce professional and polished outputs.

### Start Using IronXL

---

## Example: Adjusting Font and Size

To tailor the font characteristics of a [specific cell, column, row, or range](https://ironsoftware.com/csharp/excel/how-to/select-range/), you need to alter the **Font** attributes found in the **Style** of the cell. You can change the **Name** to pick the desired font family, adjust the **Height** for changing font size, and set **Bold** to highlight the text. Moreover, the **Underline** property is useful for adding an underline to make certain details stand out more.

```cs
using IronXL;
using IronXL.Styles;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

workSheet["B2"].StringValue = "Font and Size";

// Change font family
workSheet["B2"].Style.Font.Name = "Arial";  // Updated to Arial

// Adjust font size
workSheet["B2"].Style.Font.Height = 18;  // Increased size for better readability

// Highlight text using bold
workSheet["B2"].Style.Font.Bold = true;

// Add single underline
workSheet["B2"].Style.Font.Underline = FontUnderlineType.Single;

workBook.SaveAs("CustomizedFontAndSize.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/cell-font-size/set-font-and-size.webp" alt="Customize Font And Size" class="img-responsive add-shadow">
    </div>
</div>

## Advanced Font Customization Example

Beyond the basic settings, IronXL allows extensive customization of the text appearance in Excel. Here, you can apply **Italic**, utilize **Strikeout**, set **FontScript** for superscript or subscript, and choose a particular **color**. The following example showcases how these advanced options can personalize the styles of your cells.

Ensure that the **Name** property includes the exact formatting and spacing of the desired font, as shown below.

```cs
using IronXL;
using IronXL.Styles;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

workSheet["B2"].StringValue = "Enhanced Styles";

// Using a different font family
workSheet["B2"].Style.Font.Name = "Courier New";  // Changed to Courier New

// Applying subscript
workSheet["B2"].Style.Font.FontScript = FontScript.Sub;

// Double underline for emphasis
workSheet["B2"].Style.Font.Underline = FontUnderlineType.Double;

// Retain bold effect
workSheet["B2"].Style.Font.Bold = true;

// Applying italicization
workSheet["B2"].Style.Font.Italic = true;

// Using strikeout
workSheet["B2"].Style.Font.Strikeout = true;

// Choosing font color
workSheet["B2"].Style.Font.Color = "#FF6347";  // Tomato red for a vibrant effect

workBook.SaveAs("EnhancedFontStyles.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/cell-font-size/set-font-and-size-advanced.webp" alt="Advanced Font Customization" class="img-responsive add-shadow">
    </div>
</div>

### Exploring Underline Options

Excel provides different styles of underlining, such as the Accounting underline which offers extra space around the text and confines underlines to numerical values, making numbers stand out in your documents.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/cell-font-size/underline.webp" alt="Excel Underline Options" class="img-responsive add-shadow">
    </div>
</div>

### Understanding Font Script

IronXL supports modifying font script to adjust text placement relative to the baseline: normal, superscript, and subscript, enabling precise formatting for diverse purposes like annotations or formulas.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/cell-font-size/font-script.webp" alt="Font Script Options" class="img-responsive add-shadow">
    </div>
</div>

### Setting Font Color

You can define font colors directly through the **Color** property or the `SetColor` method, accepting either a Hex code or values from the **IronSoftware.Drawing.Color** class.

```cs
using IronXL;
using IronSoftware.Drawing;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Direct assignment using Hex code
workSheet["B2"].Style.Font.Color = "#32CD32"; // Lime Green for freshness

// Set color using IronSoftware Drawing library
workSheet["B2"].Style.Font.SetColor(Color.Blue); // Blue for serenity
```