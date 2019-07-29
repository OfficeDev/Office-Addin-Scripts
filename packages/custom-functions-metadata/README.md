# Custom-Functions-Metadata

This package allows metadata for custom functions to be generated automatically from JSDoc tags and the function parameter types.

## Command-Line Interface
* [generate](#generate)

#

## generate 
Generates metadata for custom functions from source code. 

Syntax:

`custom-functions-metadata generate <sourceFile> <metadataFile>`

`sourceFile`: path to the source file (.ts or .js).
`metadataFile`: path to the metadata file (i.e functions.json).

Notes:

* The metadata file is written if there are no errors.
* Otherwise, errors are displayed in the console.

