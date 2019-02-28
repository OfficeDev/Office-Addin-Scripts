# Custom-Functions-Metadata

Provides the ability to auto generate the custom functions metadata.

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

