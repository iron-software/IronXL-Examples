# Adjusting Cell Font and Size with IronXL

***Based on <https://ironsoftware.com/how-to/cell-font-size/>***


Enhancing document aesthetics and readability through font adjustments is vital in professional document preparation. IronXL provides a straightforward solution for modifying font styles, sizes, colors, and decorations such as underline, bold, italic, super/subscript, and strikeout in C# .NET applications. This capability facilitates the production of polished and professional-looking documents conveniently without the need for interop services.

## Example: Setting Cell Font and Size

Adjusting the font of a [specific cell, column, row, or range](https://ironsoftware.com/csharp/excel/how-to/select-range/) is made simple by modifying the **Font** properties within the **Style**. To choose a font family, modify the **Name** property. For changing the font size, adjust the **Height** property, and to make text bold, use the **Bold** property. Additionally, the **Underline** property can be set to underline text, adding further emphasis.

```cs
using IronXL.Styles;
using IronXL.Excel;
namespace ironxl.CellFontSize
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Text content for the cell
            workSheet["B2"].StringValue = "Font and Size";
            
            // Choosing font family
            workSheet["B2"].Style.Font.Name = "Times New Roman";
            
            // Adjusting font size
            workSheet["B2"].Style.Font.Height = 15;
            
            // Making font bold
            workSheet["B2"].Style.Font.Bold = true;
            
            // Applying underline
            workSheet["B2"].Style.Font.Underline = FontUnderlineType.Single;
            
            // Saving the workbook with changes
            workBook.SaveAs("fontAndSize.xlsx");
        }
    }
}
```

![Set Font And Size](https://ironsoftware.com/static-assets/excel/how-to/cell-font-size/set-font-and-size.webp)

## Advanced Example: Extensive Font Customization

Building on the basics, further customization can be achieved by setting italic and strikeout styles, choosing specific colors, and applying font scripts for super or subscript text styling.

When specifying the **Name** property, ensure the font name matches the exact format, like "Times New Roman" with spaces and capitalization.

```cs
using IronXL.Styles;
using IronXL.Excel;
namespace ironxl.CellFontSize
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            workSheet["B2"].StringValue = "Advanced";
            
            // Setting font and its properties
            workSheet["B2"].Style.Font.Name = "Lucida Handwriting";
            workSheet["B2"].Style.Font.FontScript = FontScript.None;
            workSheet["B2"].Style.Font.Underline = FontUnderlineType.Double;
            workSheet["B2"].Style.Font.Bold = true;
            workSheet["B2"].Style.Font.Italic = false;
            workSheet["B2"].Style.Font.Strikeout = false;
            workSheet["B2"].Style.Font.Color = "#00FFFF";
            
            // Save advanced styled workbook
            workBook.SaveAs("fontAndSizeAdvanced.xlsx");
        }
    }
}
```

![Set Font And Size Advanced](https://ironsoftware.com/static-assets/excel/how-to/cell-font-size/set-font-and-size-advanced.webp)

### Exploring Underline Options

Different underline styles are available in Excel, including Accounting, which adds extra space around the underline, suitable for different types of cell content, particularly for cells containing numbers.

![Available Underline Options](https://ironsoftware.com/static-assets/excel/how-to/cell-font-size/underline.webp)

### Utilizing Font Script

Font scripts in IronXL enable text placement customization:
- **none**: Default setting, aligns text on the baseline.
- **super**: Raises text above the baseline.
- **sub**: Lowers text below the baseline for nuanced mathematical and chemical notation.

![Available Font Script Options](https://ironsoftware.com/static-assets/excel/how-to/cell-font-size/font-script.webp)

### Setting Font Color

Font colors can be set via the **Color** property or through the `SetColor` method using either an **IronSoftware.Drawing.Color** object or Hex color codes.

```cs
using IronSoftware.Drawing;
using IronXL.Excel;
namespace ironxl.CellFontSize
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Set font color directly
            workSheet["B2"].Style.Font.Color = "#00FFFF";
            
            // Apply Hex color code for font
            workSheet["B2"].Style.Font.SetColor("#00FFFF");
            
            // Apply color using IronSoftware.Drawing
            workSheet["B2"].Style.Font.SetColor(Color.Red);
        }
    }
}
```