using IronSoftware.Drawing;
using IronXL;
using IronXL.Styles;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

var range = workSheet["A1:H10"];

var cell = range.First();

// Set background color of the cell with an rgb string
cell.Style.SetBackgroundColor("#428D65");

// Apply styling to the whole range.

// Set underline property to the font
// FontUnderlineType is enum that stands for different types of font underlying
range.Style.Font.Underline = FontUnderlineType.SingleAccounting;

// Define whether to use horizontal line through the text or not
range.Style.Font.Strikeout = false;

// Define whether the font is bold or not
range.Style.Font.Bold = true;

// Define whether the font is italic or not
range.Style.Font.Italic = false;

// Get or set script property of a font
// Font script enum stands for available options
range.Style.Font.FontScript = FontScript.Super;

// Set the type of the border line
// There are also TopBorder,LeftBorder,RightBorder,DiagonalBorder properties
// BorderType enum indicates the line style of a border in a cell
range.Style.BottomBorder.Type = BorderType.MediumDashed;

// Indicate whether the cell should be auto-sized
range.Style.ShrinkToFit = true;

// Set alignment of the cell
range.Style.VerticalAlignment = VerticalAlignment.Bottom;

// Set border color
range.Style.DiagonalBorder.SetColor("#20C96F");

// Define border type and border direction as well
range.Style.DiagonalBorder.Type = BorderType.Thick;

// DiagonalBorderDirection enum stands for direction of diagonal border inside cell
range.Style.DiagonalBorderDirection = DiagonalBorderDirection.Forward;

// Set background color of cells
range.Style.SetBackgroundColor(Color.Aquamarine);

// Set fill pattern of the cell
// FillPattern enum indicates the style of fill pattern
range.Style.FillPattern = FillPattern.Diamonds;

// Set the number of spaces to intend the text
range.Style.Indention = 5;

// Indicate if the text is wrapped
range.Style.WrapText = true;

workBook.SaveAs("stylingOptions.xls");
