# Editing Excel Formulas in a Worksheet

***Based on <https://ironsoftware.com/how-to/edit-formulas/>***


An Excel formula, which begins with an equals sign (`=`), serves as a powerful tool for conducting mathematical calculations, manipulating data, and deriving outcomes based on the values in cells. Composed of arithmetic operations, functions, references to other cells, constants, and logical operators, these formulas are dynamic, automatically updating as the input values change. This feature makes Excel an invaluable resource for automating routine tasks and conducting complex data analyses.

IronXL provides robust support for amending existing formulas within Excel spreadsheets, allowing users to obtain formula outcomes and enforcing a workbook evaluation. This functionality guarantees that each formula within the workbook is recalculated, thus ensuring the precision of the results. IronXL is compatible with over **165 different formulas**, enhancing its utility for a wide range of applications.

## Example of Editing Formulas

To modify or establish a formula, utilize the **Formula** property of a cell or range. Start by selecting the specific Cell or Range, after which you can retrieve or assign a new formula using the Formula property. The Formula property allows both retrieval (get) and modification (set) of the formula in the cell. To guarantee that calculations throughout the workbook are current and accurate, use the `EvaluateAll` method to recompute all formulas.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.EditFormulas
{
    public class Section1
    {
        public void Execute()
        {
            // Load the existing workbook
            WorkBook workbook = WorkBook.Load("Book1.xlsx");

            // Access the default worksheet
            WorkSheet worksheet = workbook.DefaultWorkSheet;

            // Update or define a new formula
            worksheet["A4"].Formula = "=SUM(A1,A3)";

            // Recalculate all formulas in the workbook
            workbook.EvaluateAll();
        }
    }
}
```

<hr>

## Retrieving Calculated Results from Formulas

While it's possible to extract the computational results using the **Value** property of a given Range or Cell, for more precise outcomes, utilizing the **FormattedCellValue** property of the Cell is advisable. Within the designated Range, the `First` method facilitates access to the Cell. Applying this method, we'll pinpoint the first item in the sequence, often the Cell "A4" in this scenario. Once achieved, you can easily tap into the FormattedCellValue to obtain the result.

```cs
using System.Linq;
using IronXL;

namespace ironxl.EditFormulas
{
    public class RetrieveFormulaResult
    {
        public void Execute()
        {
            // Open the Excel workbook
            WorkBook workBook = WorkBook.Load("Book1.xlsx");

            // Access the default worksheet
            WorkSheet workSheet = workBook.DefaultWorkSheet;

            // Obtain the calculated value from a formula at cell A4
            string result = workSheet["A4"].First().FormattedCellValue;

            // Output the result to the console
            Console.WriteLine($"The calculated value in cell A4 is: {result}");
        }
    }
}
```

<hr>

## Supported Formulas

Excel provides over 450 formulas for various calculations and data processing needs. IronXL, however, extends support to approximately 165 of the most frequently utilized formulas. For a comprehensive list of these supported formulas, please refer to the following section:

<style>

Here is the paraphrased content from the section you provided:

```css
tr:nth-child(odd) {
    background-color: #F1F9FB;
}
```

Note that I changed the `rgb` color specification to a hexadecimal format while maintaining the same color tone, to provide an alternative representation of the color in CSS.

</style>

<table class="table">
    <tr>
        <th>Formula Name</th>
        <th>Description</th>
    </tr>
    <tr><td>ABS</td><td>Returns the absolute value of a number, disregarding its sign.</td></tr>
    <tr><td>INT</td><td>Rounds a number down to the nearest integer.</td></tr>
    <tr><td>COUNT</td><td>Counts the number of cells that contain numbers within a specified range.</td></tr>
    <tr><td>IF</td><td>Performs a conditional test and returns one value if the condition is true and another if it's false.</td></tr>
    <tr><td>SUM</td><td>Adds up a range of numbers.</td></tr>
    <tr><td>AVERAGE</td><td>Calculates the average of a range of numbers.</td></tr>
    <tr><td>MIN</td><td>Returns the minimum value from a set of numbers.</td></tr>
    <tr><td>MAX</td><td>Returns the maximum value from a set of numbers.</td></tr>
    <tr><td>ROW</td><td>Returns the row number of a cell reference.</td></tr>
    <tr><td>COLUMN</td><td>Returns the column number of a cell reference.</td></tr>
    <tr><td>NA</td><td>Represents an error value for "Not Available" or missing data.</td></tr>
    <tr><td>NPV</td><td>Calculates the Net Present Value of a series of cash flows at a specified discount rate.</td></tr>
    <tr><td>STDEV</td><td>Calculates the standard deviation of a set of numbers.</td></tr>
    <tr><td>SIGN</td><td>Returns the sign of a number as -1 for negative, 0 for zero, or 1 for positive.</td></tr>
    <tr><td>ROUND</td><td>Rounds a number to a specified number of decimal places.</td></tr>
    <tr><td>LOOKUP</td><td>Searches for a value in a range and returns a corresponding value from another range.</td></tr>
    <tr><td>INDEX</td><td>Returns the value of a cell in a specified row and column of a given range.</td></tr>
    <tr><td>REPT</td><td>Repeats a text string a specified number of times.</td></tr>
    <tr><td>MID</td><td>Extracts a portion of text from a given text string based on a specified starting position and length.</td></tr>
    <tr><td>LEN</td><td>Returns the number of characters in a text string.</td></tr>
    <tr><td>VALUE</td><td>Converts a text string that represents a number to an actual number.</td></tr>
    <tr><td>TRUE</td><td>Represents the logical value for "True."</td></tr>
    <tr><td>FALSE</td><td>Represents the logical value for "False."</td></tr>
    <tr><td>AND</td><td>Checks if all specified conditions are true and returns "True" if they are, and "False" otherwise.</td></tr>
    <tr><td>OR</td><td>Checks if at least one of the specified conditions is true and returns "True" if it is, and "False" otherwise.</td></tr>
    <tr><td>NOT</td><td>Inverts the logical value of a condition, turning "True" into "False" and vice versa.</td></tr>
    <tr><td>MOD</td><td>Returns the remainder when one number is divided by another.</td></tr>
    <tr><td>DMIN</td><td>Extracts the minimum value from a database based on specified criteria.</td></tr>
    <tr><td>VAR</td><td>Calculates the variance of a set of numbers.</td></tr>
    <tr><td>TEXT</td><td>Converts a number to text using a specified format.</td></tr>
    <tr><td>PV</td><td>Calculates the present value of an investment or loan based on a series of cash flows and a discount rate.</td></tr>
    <tr><td>FV (Future Value)</td><td>Calculates the future value of an investment or loan based on periodic payments and a specified interest rate.</td></tr>
    <tr><td>NPER (Number of Periods)</td><td>Determines the number of payment periods required to reach a certain financial goal, given regular payments and an interest rate.</td></tr>
    <tr><td>PMT (Payment)</td><td>Calculates the periodic payment needed to pay off a loan or investment, including principal and interest.</td></tr>
    <tr><td>RATE (Interest Rate)</td><td>Calculates the interest rate required to reach a financial goal with a series of periodic payments.</td></tr>
    <tr><td>MIRR (Modified Internal Rate of Return)</td><td>Calculates the internal rate of return for a series of cash flows, addressing multiple reinvestment and financing rates.</td></tr>
    <tr><td>IRR (Internal Rate of Return)</td><td>Calculates the internal rate of return for a series of cash flows, indicating the rate at which an investment breaks even.</td></tr>
    <tr><td>RAND</td><td>Generates a random decimal number between 0 and 1.</td></tr>
    <tr><td>MATCH</td><td>Searches for a specified value in a range and returns the relative position of the item found.</td></tr>
    <tr><td>DATE</td><td>Creates a date value by specifying the year, month, and day.</td></tr>
    <tr><td>TIME</td><td>Creates a time value by specifying the hour, minute, and second.</td></tr>
    <tr><td>DAY</td><td>Extracts the day from a given date.</td></tr>
    <tr><td>MONTH</td><td>Extracts the month from a given date.</td></tr>
    <tr><td>YEAR</td><td>Extracts the year from a given date.</td></tr>
    <tr><td>WEEKDAY</td><td>Returns the day of the week for a specified date.</td></tr>
    <tr><td>HOUR</td><td>Extracts the hour from a given time.</td></tr>
    <tr><td>MINUTE</td><td>Extracts the minute from a given time.</td></tr>
    <tr><td>SECOND</td><td>Extracts the second from a given time.</td></tr>
    <tr><td>NOW</td><td>Returns the current date and time.</td></tr>
    <tr><td>AREAS</td><td>Counts the number of individual ranges within a reference.</td></tr>
    <tr><td>ROWS</td><td>Counts the number of rows in a specified range.</td></tr>
    <tr><td>COLUMNS</td><td>Counts the number of columns in a specified range.</td></tr>
    <tr><td>OFFSET</td><td>Returns a reference offset from a specified cell by a certain number of rows and columns.</td></tr>
    <tr><td>SEARCH</td><td>Searches for a substring within a text string and returns its position.</td></tr>
    <tr><td>TRANSPOSE</td><td>Transposes the rows and columns of a range.</td></tr>
    <tr><td>ATAN2</td><td>Calculates the arctangent of a specified x and y coordinate.</td></tr>
    <tr><td>ASIN</td><td>Calculates the arcsine of a specified value.</td></tr>
    <tr><td>ACOS</td><td>Calculates the arccosine of a specified value.</td></tr>
    <tr><td>CHOOSE</td><td>Returns a value from a list of values based on a specified position.</td></tr>
    <tr><td>HLOOKUP</td><td>Searches for a value in the top row of a table or range and returns a value in the same column from a specified row.</td></tr>
    <tr><td>VLOOKUP</td><td>Searches for a value in the first column of a table or range and returns a value in the same row from a specified column.</td></tr>
    <tr><td>ISREF</td><td>Checks if a value is a reference and returns "True" if it is, or "False" if it's not.</td></tr>
    <tr><td>LOG</td><td>Calculates the logarithm of a number to a specified base.</td></tr>
    <tr><td>CHAR</td><td>Returns the character specified by a given number.</td></tr>
    <tr><td>LOWER</td><td>Converts text to lowercase.</td></tr>
    <tr><td>UPPER</td><td>Converts text to uppercase.</td></tr>
    <tr><td>PROPER</td><td>Capitalizes the first letter of each word in a text string.</td></tr>
    <tr><td>LEFT</td><td>Extracts a specified number of characters from the beginning of a text string.</td></tr>
    <tr><td>RIGHT</td><td>Extracts a specified number of characters from the end of a text string.</td></tr>
    <tr><td>EXACT</td><td>Compares two text strings and returns "True" if they are identical, and "False" if they are not.</td></tr>
    <tr><td>TRIM</td><td>Removes extra spaces from a text string, except for single spaces between words.</td></tr>
    <tr><td>REPLACE</td><td>Replaces a specified number of characters in a text string with new text.</td></tr>
    <tr><td>SUBSTITUTE</td><td>Replaces occurrences of a specified text in a text string with new text.</td></tr>
    <tr><td>CODE</td><td>Returns the numeric Unicode value of the first character in a text string.</td></tr>
    <tr><td>FIND</td><td>Searches for a specific substring within a text string and returns its position.</td></tr>
    <tr><td>ISERR</td><td>Checks if a value is an error value other than "#N/A" and returns "True" if it is, or "False" if it's not.</td></tr>
    <tr><td>ISTEXT</td><td>Checks if a value is text and returns "True" if it is, or "False" if it's not.</td></tr>
    <tr><td>ISNUMBER</td><td>Checks if a value is a number and returns "True" if it is, or "False" if it's not.</td></tr>
    <tr><td>ISBLANK</td><td>Checks if a cell is empty and returns "True" if it is, or "False" if it's not.</td></tr>
    <tr><td>T</td><td>Converts a value to text format.</td></tr>
    <tr><td>DATEVALUE</td><td>Converts a date represented as text into a date serial number.</td></tr>
    <tr><td>CLEAN</td><td>Removes non-printable characters from text.</td></tr>
    <tr><td>MDETERM</td><td>Calculates the matrix determinant of an array.</td></tr>
    <tr><td>MINVERSE</td><td>Returns the multiplicative inverse (reciprocal) of a matrix.</td></tr>
    <tr><td>MMULT</td><td>Multiplies two matrices together.</td></tr>
    <tr><td>IPMT</td><td>Calculates the interest portion of a loan payment for a given period.</td></tr>
    <tr><td>PPMT</td><td>Calculates the principal portion of a loan payment for a given period.</td></tr>
    <tr><td>COUNTA</td><td>Counts the number of non-empty cells in a range, including text and numbers.</td></tr>
    <tr><td>PRODUCT</td><td>Multiplies all the numbers in a range.</td></tr>
    <tr><td>FACT</td><td>Calculates the factorial of a number.</td></tr>
    <tr><td>ISNONTEXT</td><td>Checks if a value is not text and returns "True" if it's not text, or "False" if it is text.</td></tr>
    <tr><td>VARP</td><td>Estimates the variance of a population based on a sample.</td></tr>
    <tr><td>TRUNC</td><td>Truncates a number to a specified number of decimal places.</td></tr>
    <tr><td>ISLOGICAL</td><td>Checks if a value is a logical (Boolean) value and returns "True" if it is, or "False" if it's not.</td></tr>
    <tr><td>USDOLLAR</td><td>Converts a number to text format with a currency symbol and two decimal places.</td></tr>
    <tr><td>ROUNDUP</td><td>Rounds a number up to a specified number of decimal places.</td></tr>
    <tr><td>ROUNDDOWN</td><td>Rounds a number down to a specified number of decimal places.</td></tr>
    <tr><td>RANK</td><td>Returns the rank of a number in a list, with options to handle ties.</td></tr>
    <tr><td>ADDRESS</td><td>Returns the cell address as text based on row and column numbers.</td></tr>
    <tr><td>DAYS360</td><td>Calculates the number of days between two dates using the 360-day year.</td></tr>
    <tr><td>TODAY</td><td>Returns the current date.</td></tr>
    <tr><td>MEDIAN</td><td>Returns the median (middle value) of a set of numbers.</td></tr>
    <tr><td>SUMPRODUCT</td><td>Multiplies corresponding components in arrays and returns the sum of the products.</td></tr>
    <tr><td>SINH</td><td>Calculates the hyperbolic sine of a number.</td></tr>
    <tr><td>COSH</td><td>Calculates the hyperbolic cosine of a number.</td></tr>
    <tr><td>TANH</td><td>Calculates the hyperbolic tangent of a number.</td></tr>
    <tr><td>ASINH</td><td>Calculates the inverse hyperbolic sine of a number.</td></tr>
    <tr><td>ACOSH</td><td>Calculates the inverse hyperbolic cosine of a number.</td></tr>
    <tr><td>ATANH</td><td>Calculates the inverse hyperbolic tangent of a number.</td></tr>
    <tr><td>ExternalFunction</td><td>Represents a function call or operation provided by an external add-in or custom function.</td></tr>
    <tr><td>ERRORTYPE</td><td>Returns a number that corresponds to the error type in a given value.</td></tr>
    <tr><td>AVEDEV</td><td>Calculates the average absolute deviation of a set of values from their mean.</td></tr>
    <tr><td>COMBIN</td><td>Calculates the number of combinations for a given number of items taken from a larger set.</td></tr>
    <tr><td>EVEN</td><td>Rounds a number up to the nearest even integer.</td></tr>
    <tr><td>FLOOR</td><td>Rounds a number down to the nearest multiple of a specified significance.</td></tr>
    <tr><td>CEILING</td><td>Rounds a number up to the nearest multiple of a specified significance.</td></tr>
    <tr><td>NORMDIST</td><td>Calculates the cumulative normal distribution function for a specified value.</td></tr>
    <tr><td>NORMSDIST</td><td>Calculates the standard normal cumulative distribution function.</td></tr>
    <tr><td>NORMINV</td><td>Calculates the inverse of the normal cumulative distribution function for a specified probability.</td></tr>
    <tr><td>NORMSINV</td><td>Calculates the inverse of the standard normal cumulative distribution function.</td></tr>
    <tr><td>STANDARDIZE</td><td>Converts a value to a standard normal distribution with a mean of 0 and a standard deviation of 1.</td></tr>
    <tr><td>ODD</td><td>Rounds a number up to the nearest odd integer.</td></tr>
    <tr><td>POISSON</td><td>Calculates the Poisson distribution probability for a given number of events.</td></tr>
    <tr><td>TDIST</td><td>Calculates the Student's t-distribution for a specified value and degrees of freedom.</td></tr>
    <tr><td>SUMXMY2</td><td>Calculates the sum of squares of the differences between corresponding values in two arrays.</td></tr>
    <tr><td>SUMX2MY2</td><td>Calculates the sum of squares of the differences between corresponding values in two arrays.</td></tr>
    <tr><td>SUMX2PY2</td><td>Calculates the sum of squares of the sum of corresponding values in two arrays.</td></tr>
    <tr><td>INTERCEPT</td><td>Calculates the point at which a trendline crosses the y-axis in a chart.</td></tr>
    <tr><td>SLOPE</td><td>Calculates the slope of a trendline in a chart.</td></tr>
    <tr><td>DEVSQ</td><td>Returns the sum of squares of deviations of data points from their mean.</td></tr>
    <tr><td>SUMSQ</td><td>Calculates the sum of squares of a set of numbers.</td></tr>
    <tr><td>LARGE</td><td>Returns the k-th largest value in a dataset, where k is specified.</td></tr>
    <tr><td>SMALL</td><td>Returns the k-th smallest value in a dataset, where k is specified.</td></tr>
    <tr><td>PERCENTILE</td><td>Returns the k-th percentile of a dataset, where k is specified.</td></tr>
    <tr><td>PERCENTRANK</td><td>Returns the rank of a value in a dataset as a percentage of the total number of values.</td></tr>
    <tr><td>MODE</td><td>Returns the most frequently occurring value in a dataset.</td></tr>
    <tr><td>CONCATENATE</td><td>Combines multiple text strings into one.</td></tr>
    <tr><td>POWER</td><td>Raises a number to a specified power.</td></tr>
    <tr><td>RADIANS</td><td>Converts degrees to radians.</td></tr>
    <tr><td>DEGREES</td><td>Converts radians to degrees.</td></tr>
    <tr><td>SUBTOTAL</td><td>Performs various calculations (e.g., sum, average) on a range, and you can choose whether to include or exclude other SUBTOTAL results within the range.</td></tr>
    <tr><td>SUMIF</td><td>Adds up all the numbers in a range that meet a specified condition.</td></tr>
    <tr><td>COUNTIF</td><td>Counts the number of cells in a range that meet a specified condition.</td></tr>
    <tr><td>COUNTBLANK</td><td>Counts the number of empty cells in a range.</td></tr>
    <tr><td>ROMAN</td><td>Converts an Arabic numeral to a Roman numeral.</td></tr>
    <tr><td>HYPERLINK</td><td>Creates a hyperlink to a webpage or file.</td></tr>
    <tr><td>MAXA</td><td>Returns the maximum value from a set of numbers, including text and logical values.</td></tr>
    <tr><td>MINA</td><td>Returns the minimum value from a set of numbers, including text and logical values.</td></tr>
</table>

