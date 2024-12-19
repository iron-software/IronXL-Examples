***Based on <https://ironsoftware.com/examples/excel-conditional-formatting/>***

The IronXL library enables **Conditional Formatting** for cells and ranges, allowing dynamic changes to cell styles like background color or text style based on specified logical or programmatic rules.

You can initiate a conditional formatting rule using the `CreateConditionalFormattingRule(string formula)` method, wherein a rule is applied based on a Boolean formula. When the formula evaluates to true, the respective cell is visually accentuated. It is crucial to ensure that the formula provided is a Boolean function.

For the `CreateConditionalFormattingRule` method that accepts three parameters, only `ComparisonOperator.Between` and `ComparisonOperator.NotBetween` can be used as the first parameter.

Conditional formatting is a powerful tool to enhance visual analysis by highlighting cells and ranges with specific color and format combinations. These highlights are based on the truth value of the cell, determined by predefined rules, enabling easier data analysis and helping in the detection of issues, recognition of patterns, and identification of trends.