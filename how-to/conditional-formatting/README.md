# Implementing Conditional Formatting with Cells

***Based on <https://ironsoftware.com/how-to/conditional-formatting/>***


Conditional Formatting is an invaluable tool for spreadsheet and data management applications, facilitating the visual differentiation of data based on specific conditions or criteria. This functionality is crucial for highlighting significant data points within a spreadsheet or table, allowing for a streamlined data analysis and interpretation process.

IronXL enhances this experience by offering straightforward methods to add, retrieve, and erase conditional formatting. This includes customization capabilities like [font and size adjustments](https://ironsoftware.com/csharp/excel/how-to/cell-font-size/), [borders and alignment](https://ironsoftware.com/csharp/excel/how-to/border-alignment/), and [background patterns and colors](https://ironsoftware.com/csharp/excel/how-to/background-pattern-color/).

### Getting Started with IronXL

---

## Example: Adding Conditional Formatting

Conditional formatting in IronXL involves creating rules and applying specific styles when cells meet these rules. Styles may encompass [font and size adjustments](https://ironsoftware.com/csharp/excel/how-to/cell-font-size/), [border and alignment settings](https://ironsoftware.com/csharp/excel/how-to/border-alignment/), and [background patterns and colors](https://ironsoftware.com/csharp/excel/how-to/background-pattern-color/).

To set up a rule, harness the `CreateConditionalFormattingRule` method from the `ConditionalFormatting` object. Assigning the resulting object to a variable allows you to impose the desired styling. To implement the styling, use the `AddConditionalFormatting` method and specify both the rule and the affected cell range.

```cs
using IronXL;
using IronXL.Formatting.Enums;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Define a conditional formatting rule
var rule = workSheet.ConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.LessThan, "8");

// Configure style settings
rule.PatternFormatting.BackgroundColor = "#54BDD9";

// Implement the conditional formatting rule
workSheet.ConditionalFormatting.AddConditionalFormatting("A1:A10", rule);

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

Here are the types of available conditional formatting rules:
- **NoComparison**: Default rule.
- **Between**: Checks if values lie within a specific range.
- **NotBetween**: Ensures values do not lie within a specified range.
- **Equal**: Tests for equality.
- **NotEqual**: Checks for non-equality.
- **GreaterThan**: Evaluates if greater than a value.
- **LessThan**: Checks if less than a value.
- **GreaterThanOrEqual**: Determines if values are greater than or equal to a certain number.
- **LessThanOrEqual**: Assesses if values are less than or equal to a specified number.

<hr>

## Example: Retrieving Conditional Formatting

To access an applied conditional formatting rule, utilize the `GetConditionalFormattingAt` method. This method might return a set of rules, from which the `GetRule` method can be used to select an individual rule. Although many attributes of a returned rule are immutable, the **BackgroundColor** can be altered via the **PatternFormatting** property.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("addConditionalFormatting.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Access the first conditional formatting collection and retrieve its first rule
var ruleCollection = workSheet.ConditionalFormatting.GetConditionalFormattingAt(0);
var rule = ruleCollection.GetRule(0);

// Modify the styling
rule.PatternFormatting.BackgroundColor = "#B6CFB6";

workBook.SaveAs("editedConditionalFormatting.xlsx");
```

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

<hr>

## Example: Removing Conditional Formatting

To erase a particular conditional formatting rule, employ the `RemoveConditionalFormatting` method, which requires the index of the rule to be removed.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("addConditionalFormatting.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Erase the first conditional formatting rule
workSheet.ConditionalFormatting.RemoveConditionalFormatting(0);

workBook.SaveAs("removedConditionalFormatting.xlsx");
```