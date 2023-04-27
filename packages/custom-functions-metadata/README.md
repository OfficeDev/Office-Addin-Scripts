# Custom-Functions-Metadata

This package allows metadata for custom functions to be generated automatically from JSDoc tags and the function parameter types.

## Command-Line Interface
* [generate](#generate)

#

## generate 
Generates metadata for custom functions from source code. 

Syntax:

`custom-functions-metadata generate <sourceFile> [outputFile] [options]`

`sourceFile`: path to the source file (.ts or .js).
`outputFile`: If specified, the metadata is written to the file. Otherwise, it is written to the console.

Options:

`--allow-custom-data-for-data-type-any`

Allow a custom function to return a 
[FormattedNumberCellValue](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-data-types-concepts#output-a-formatted-number-value) 
data type.