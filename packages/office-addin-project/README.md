# Office-Addin-Manifest

This package provides the ability to do project wide commands for Office Add-ins, such as conversion.

For more information, see the [documentation](
https://docs.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests).

## Command-Line Interface
* [convert](#info)

#

### convert

Converts the Office Addin Code from xml to json based manifest.

Syntax:

`office-addin-cli convert [options]`

`manifest-path`: path to manifest file.

Options:

`-m <manifest-path>`<br>
`--manifest <manifest-path>`

Specify the location of the manifest file. If the path is not provided, `./manifest.xml` is used.
