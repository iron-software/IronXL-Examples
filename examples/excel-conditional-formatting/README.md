***Based on <https://ironsoftware.com/examples/excel-conditional-formatting/>***

The IronXL library offers the feature of **Conditional Formatting** for cells and ranges. This functionality enables users to dynamically alter cell styles, including background colors and text styles, based on specific logical or programmatic conditions.

You can establish a conditional formatting rule by using the method `CreateConditionalFormattingRule(string formula)`. This rule is triggered when the specified Boolean formula evaluates to true, leading to the highlighting of the cell. It's crucial to ensure that the formula provided is a Boolean function.

The version of the `CreateConditionalFormattingRule` method that accepts three parameters is restricted to using `ComparisonOperator.Between` and `ComparisonOperator.NotBetween` for the first parameter.

Conditional formatting is a powerful tool used to visually enhance cells and ranges based on specified color and format schemes, contingent on the true/false outcomes of defined rules. This feature is particularly useful for data analysis, problem detection, and the identification of patterns and trends in data.