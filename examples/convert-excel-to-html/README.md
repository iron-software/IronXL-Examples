***Based on <https://ironsoftware.com/examples/convert-excel-to-html/>***

The code snippet provided demonstrates the method for transforming Excel documents into HTML format using C#. To tailor the HTML output, employ the `HtmlExportOptions` class which comprises several customizable features:

- The **`OutputRowNumbers`** property determines if row numbers should be displayed in the resulting file.
- The **`OutputColumnHeaders`** property specifies whether column headers should be included in the output file.
- The **`OutputHiddenRows`** property controls the visibility of hidden rows in the output.
- The **`OutputHiddenColumns`** property decides whether hidden columns should be visible in the output.
- Finally, the **`OutputLeadingSpacesAsNonBreaking`** property indicates if leading spaces (extra spaces before the first character in a cell) should appear as non-breaking spaces in the resulting file.