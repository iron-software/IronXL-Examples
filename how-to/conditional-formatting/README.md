# How to Apply Conditional Formatting to Cells

Conditional formatting is an essential tool in spreadsheet and data processing software that enables users to apply specific formatting styles to cells or data based on defined criteria or conditions. This functionality allows for the visual distinction of data that adheres to certain conditions, thereby simplifying data analysis and interpretation in a spreadsheet or table.

Effortlessly Add, Retrieve, and Clear Conditional Formatting with IronXL. With this tool, adjustments can be made to [font and size](https://ironsoftware.com/csharp/excel/how-to/cell-font-size/), [borders and alignment](https://ironsoftware.com/csharp/excel/how-to/border-alignment/), and [background patterns and colors](https://ironsoftware.com/csharp/excel/how-to/background-pattern-color/).

## Example of Adding Conditional Formatting

Conditional formatting involves rules and styles applied to cells when they satisfy certain conditions. The styles can range from [font and size adjustments](https://ironsoftware.com/csharp/excel/how-to/cell-font-size/), [borders and alignment configurations](https://ironsoftware.com/csharp/excel/how-to/border-alignment/), to [background patterns and colors](https://ironsoftware.com/csharp/excel/how-to/background-pattern-color/).

To create a rule, utilize the `CreateConditionalFormattingRule` method from `ConditionalFormatting`. Assign the result of this method to a variable to facilitate the styling application. Then, employ the `AddConditionalFormatting` method, inputting the established rule and the cell range where it should be implemented.

```cs
using IronXL;
using IronXL.Formatting.Enums;

// Load a workbook
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Establish a conditional formatting rule
var rule = workSheet.ConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.LessThan, "8");

// Configure style options
rule.PatternFormatting.BackgroundColor = "#54BDD9";

// Implement the conditional formatting rule
workSheet.ConditionalFormatting.AddConditionalFormatting("A1:A10", rule);

// Save changes
workBook.SaveAs("addConditionalFormatting.xlsx");
```

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

Below are all the available rules:
- NoComparison: The default value.
- Between: 'Between' operator
- NotBetween: 'Not between' operator
- Equal: 'Equal to' operator
- NotEqual: 'Not equal to' operator
- GreaterThan: 'Greater than' operator
- LessThan: 'Less than' operator
- GreaterThanOrEqual: 'Greater than or equal to' operator
- LessThanOrEqual: 'Less than or equal to' operator

<hr>

## Example of Retrieving Conditional Formatting

To extract a conditional formatting rule, employ the `GetConditionalFormattingAt` method from the workbook. This method retrieves a collection of rules, from which the `GetRule` method extracts a specific one. Although most attributes of the rule cannot be altered, the **BackgroundColor** can be modified by accessing the **PatternFormatting** property, as illustrated below.

```cs
using IronXL;

// Load the workbook
WorkBook workBook = WorkBook.Load("addConditionalFormatting.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Acquire the conditional formatting rule collection
var ruleCollection = workSheet.ConditionalFormatting.GetConditionalFormattingAt(0);
var rule = ruleCollection.GetRule(0);

// Modify the styling
rule.PatternFormatting.BackgroundColor = "#B6CFB6";

// Save the updated formatting
workBook.SaveAs("editedConditionalFormatting.xlsx");
```

<div class="competitors-section__wrapper-even-1">
    <div class="competitors__card" style="width: 49%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/conditional-formatting/after.png" alt="Before" class="img-responsive add-shadow" style="margin: auto;">
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">Before</p>
    </div>
    <div class="competitors__card" style="width: 49%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/conditional-formatting/edit-style.png" alt="After" class="img-responsive add-shadow" style="margin:auto;">
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">After</p>
    </div>
</div>

<hr>

## Example of Removing Conditional Formatting

To eliminate a conditional formatting rule, utilize the `RemoveConditionalFormatting` method. Just specify the index of the rule you wish to remove.

```cs
using IronXL;

// Load the workbook
WorkBook workBook = WorkBook.Load("addConditionalFormatting.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Delete the conditional formatting rule
workSheet.ConditionalFormatting.RemoveConditionalFormatting(0);

// Save the workbook without the rule
workBook.SaveAs("removedConditionalFormatting.xlsx");
```