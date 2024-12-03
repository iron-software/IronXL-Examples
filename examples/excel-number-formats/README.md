***Based on <https://ironsoftware.com/examples/excel-number-formats/>***

To customize the display format of values in any Excel `Cell` or `Range` within C#, you can leverage the `FormatString` property when using IronXL.

The statement `workSheet["A2"]` specifically targets a `Range` at the provided address. To access an individual `Cell`, you would typically use the `First()` method. However, the application of the `FormatString` property can be done for both `Cell` and `Range`, making use of `First()` unnecessary in this context. A variety of Excel number formats can be implemented using this approach, as demonstrated in the given code example.