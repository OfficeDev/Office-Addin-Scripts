# Custom-Functions-Metadata

Provides the ability to auto generate the custom functions metadata.

## Command-Line Interface
* [generate](#generate)

#

## generate 
Generates metadata for custom functions from source code. 

Syntax:

`custom-functions-metadata generate <inputFile> <output>`

`inputFile`: path to custom functions file (.ts or .js).
`output`: filename of the metadata file (i.e functions.json).

Notes:

* Output file is generated, if no errors found during processing.
* If errors are found, they will be displayed in the console.
