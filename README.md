# PackXLSX.js
Simple .xlsx writer with basic features:
* Numbers
* Strings
* Formulas (strings starting with `=`)
* Multiple sheets.

Input: array of sheets  
Sheet: optional name (string), array of rows  
Row: empty row (`undefined`) or array of cells  
Cell: empty cell (`undefined`) or number/string/formula.  
Automatic names of sheets are `Sheet1`, `Sheet2`...
## Example 1.
    const xl = packxlsx([[["Hello", , "Word!"], , [2, 3.5, "=A3*B3"]]]);
Single sheet with automatic name (`Sheet1`):
|| A | B | C |
|-|---|---|---|
|1|Hello||World|
|2|
|3|2|3.5|`=A3*B3`|
## Example 2.
    const xl = packxlsx([
        ["Trallala", ["Hello", , "Word!"], , [2, 3.5, "=A3*B3"]],
        [["Lorem"], [4, 5, "=A2*B3-A3*B2"], [6, 7]],
        ["Ipsum", "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged.".split("")]
    ]);
Three sheets with a bit more content:  
1. same content as above, but custom name `Trallala`
2. default name (`Sheet2`), `Lorem` text and simple determinant of a 2x2 matrix
3. custom name `Ipsum`, long row of cells containing letters of a longer text.
## Live example
The simple demo page providing 10x10 sheets and a download button can be visited here: https://tevemadar.github.io/PackXLSX.js/
