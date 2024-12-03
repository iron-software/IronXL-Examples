# Customizing Cell Borders and Text Alignment in Excel Spreadsheets

***Based on <https://ironsoftware.com/how-to/border-alignment/>***


Excel spreadsheets provide cell borders, which are visual lines outlining individual or groups of cells, and text alignment capabilities which determine the positioning of the text within those cells both vertically and horizontally.

Utilizing IronXL, you can enhance your data presentation by customizing these aspects. Styling borders and adjusting text alignments can increase readability and give a polished, professional appearance to your spreadsheets.

## Example: Setting Cell Borders and Text Alignment

Enhance a [specific cell, column, row, or range](https://ironsoftware.com/csharp/excel/how-to/select-range/) by applying borders and setting alignment through the **TopBorder**, **RightBorder**, **BottomBorder**, and **LeftBorder** properties. Choose from a variety of styles offered in the **IronXL.Styles.BorderType** enum. Check out [all available border styles](#anchor-available-border-type) to find one that suits your needs.

For precise placement of your text, adjust **HorizontalAlignment** and **VerticalAlignment** in the Style properties. The **IronXL.Styles.HorizontalAlignment** and **IronXL.Styles.VerticalAlignment** enums allow you to position text exactly how you want it within your spreadsheet. Explore [all alignment options](#anchor-available-border-type) to perfectly showcase your data.

```cs
using IronXL.Styles;
using IronXL.Excel;
namespace ironxl.BorderAlignment
{
    public class Section1
    {
        public void Execute()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            workSheet["B2"].Value = "Data Point";
            
            // Applying cell border
            workSheet["B2"].Style.LeftBorder.Type = BorderType.MediumDashed;
            workSheet["B2"].Style.RightBorder.Type = BorderType.MediumDashed;
            
            // Configuring text alignment
            workSheet["B2"].Style.HorizontalAlignment = HorizontalAlignment.Center;
            
            workBook.SaveAs("FormattedExcel.xlsx");
        }
    }
}
```

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/set-border-alignment.webp" alt="Border And Alignment" class="img-responsive add-shadow">
    </div>
</div>

## Advanced Customization: Border Color and Line Patterns

### Customizing Border Color

While the default border color is black, IronXL allows customization to any desired color using either the **Color** class or Hex color codes. Before setting a color, ensure the border type is defined.

```cs
using IronSoftware.Drawing;
using IronXL.Excel;
namespace ironxl.BorderAlignment
{
    public class Section2
    {
        public void Execute()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            workSheet["B2"].Style.LeftBorder.Type = BorderType.Thick;
            workSheet["B2"].Style.RightBorder.Type = BorderType.Thick;
            
            // Customizing border color
            workSheet["B2"].Style.LeftBorder.SetColor(Color.Aquamarine);
            workSheet["B2"].Style.RightBorder.SetColor("#FF7F50");
            
            workBook.SaveAs("CustomBorderColor.xlsx");
        }
    }
}
```

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/set-border-color.webp" alt="Border Color" class="img-responsive add-shadow">
    </div>
</div>

### Border Lines and Patterns

IronXL supports a range of border lines including top, right, bottom, left, and diagonal directions. Each of these can be set with various patterns or types as needed.

```cs
using IronXL.Styles;
using IronXL.Excel;
namespace ironxl.BorderAlignment
{
    public class Section3
    {
        public void Execute()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            workSheet["B2"].StringValue = "Peak";
            workSheet["B4"].StringValue = "Diagonal";
            
            // Establishing top border line
            workSheet["B2"].Style.TopBorder.Type = BorderType.Thick;
            
            // Designing a diagonal border
            workSheet["B4"].Style.DiagonalBorder.Type = BorderType.Thick;
            workSheet["B4"].Style.DiagonalBorderDirection = DiagonalBorderDirection.Forward;
            
            workBook.SaveAs("FormattedBorderPatterns.xlsx");
        }
    }
}
```

#### Visualizing Border Lines

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/border-line.webp" alt="Available Border Lines" class="img-responsive add-shadow">
    </div>
</div>

#### Visualizing Border Patterns

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/border-types.webp" alt="Available Border Types" class="img-responsive add-shadow">
    </div>
</div>

### Comprehensive Guide to Alignment Types

Explore IronXL's complete set of text alignment options through graphical representations below:

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/border-alignment/alignment-types.webp" alt="Available Alignment Types" class="img-responsive add-shadow">
    </div>
</div>