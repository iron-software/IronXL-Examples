# Using Conditional Formatting with IronXL

***Based on <https://ironsoftware.com/how-to/conditional-formatting/>***


Conditional formatting is a powerful tool in spreadsheet and data processing software. It allows users to apply specific styles or parameters to cells based on predefined conditions, enhancing the visibility and comprehension of data in tables or spreadsheets.

With IronXL, it's effortless to add, retrieve, and remove conditional formatting. Users can tweak styles such as [font and size](https://ironsoftware.com/csharp/excel/how-to/cell-font-size/), [borders and alignment](https://ironsoftware.com/csharp/excel/how-to/border-alignment/), and [background patterns and colors](https://ironsoftware.com/csharp/excel/how-to/background-pattern-color/) directly through simple methods.

## Example of Adding Conditional Formatting

Conditional formatting in IronXL involves creating rules and applying styles when data fulfills these rules. For instance, adjustments to [font and size](https://ironsoftware.com/csharp/excel/how-to/cell-font-size/), [borders and settings for alignment](https://ironsoftware.com/csharp/excel/how-to/border-alignment/), and even [background colors and patterns](https://ironsoftware.com/csharp/excel/how-to/background-pattern-color/) can be defined easily.

Here's how you can create a conditional formatting rule:

```cs
using IronXL.Formatting.Enums;
using IronXL.Excel;
namespace ironxl.ConditionalFormatting
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Creating a conditional formatting rule for values less than 8
            var rule = workSheet.ConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.LessThan, "8");
            
            // Setting the background color style for the rule
            rule.PatternFormatting.BackgroundColor = "#54BDD9";
            
            // Applying the rule to a range of cells
            workSheet.ConditionalFormatting.AddConditionalFormatting("A1:A10", rule);
            
            // Save the workbook with changes
            workBook.SaveAs("addConditionalFormatting.xlsx");
        }
    }
}
```

Below are the visual changes before and after applying the formatting:

<div class="competitors-section__wrapper-even-1">
    <div class="competitors__card" style="width: 49%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/conditional-formatting/before.png" alt="Before" class="img-responsive add-shadow" style="margin: auto;">
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">Before</p>
    </div>
    <div class="competitors__card" style="width: 49%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/conditional-formatting/after.png" alt="After" class="img-responsive add-shadow" style="margin: auto;">
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">After</p>
    </div>
</div>

List of available rules includes:
- NoComparison: Default
- Between
- NotBetween
- Equal
- NotEqual
- GreaterThan
- LessThan
- GreaterThanOrEqual
- LessThanOrEqual

## Retrieving Conditional Formatting

To access and modify an existing conditional formatting, you can retrieve it using the `GetConditionalFormattingAt` method. Adjust properties like **BackgroundColor** through the **PatternFormatting** attribute.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.ConditionalFormatting
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("addConditionalFormatting.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Accessing the first conditional formatting rule
            var ruleCollection = workSheet.ConditionalFormatting.GetConditionalFormattingAt(0);
            var rule = ruleCollection.GetRule(0);
            
            // Modifying the background color
            rule.PatternFormatting.BackgroundColor = "#B6CFB6";
            
            // Saving the workbook
            workBook.SaveAs("editedConditionalFormatting.xlsx");
        }
    }
}
```

The process allows a before (unaltered state) and after (modified state) view as demonstrated:

<div class="competitors-section__wrapper-even-1">
    <div class="competitors__card" style="width: 49%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/conditional-formatting/after.png" alt="Before" class="img-responsive add-shadow" style="margin: auto;">
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">Before</p>
    </div>
    <div class="competitors__card" style="width: 49%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/conditional-formatting/edit-style.png" alt="After" class="img-responsive add-shadow" style="margin: auto;">
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">After</p>
    </div>
</div>

## Removing Conditional Formatting

If you need to remove conditional formatting from a spreadsheet, use the `RemoveConditionalFormatting` method by specifying the rule index.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.ConditionalFormatting
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("addConditionalFormatting.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Removing the first conditional formatting rule
            workSheet.ConditionalFormatting.RemoveConditionalFormatting(0);
            
            // Saving the changes
            workBook.SaveAs("removedConditionalFormatting.xlsx");
        }
    }
}
```