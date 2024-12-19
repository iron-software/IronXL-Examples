***Based on <https://ironsoftware.com/examples/excel-number-formats/>***

You can utilize the `FormatString` property in C# with IronXL to format the display value of any Excel `Cell` or `Range`.

Selecting `workSheet["A2"]` targets the `Range` at the specified address. To directly manipulate a `Cell`, you might think to employ the `First()` method. Nonetheless, since the `FormatString` setting can be adjusted for both `Cell` and `Range`, employing `First()` is not obligatory. Additional Excel number formatting options are available and can be implemented similarly as illustrated in the provided code example.