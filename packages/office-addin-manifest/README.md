# Office-Addin-Manifest

This package provides the ability to parse, display, modify, and validate the manifest file for Office Add-ins.

For more information, see the [documentation](
https://docs.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests).

## Command-Line Interface
* [info](#info)
* [modify](#modify)
* [validate](#validate)

#

### info 
Display the information in the Office Add-in manifest. 

Syntax:

`office-addin-manifest info <manifest> [options]`

`manifest`: path to manifest file.

#

### modify
Modify values in the Office Add-in manifest file.

Syntax:

`office-addin-manifest modify <manifest> [options]`

`manifest`: path to manifest file. 

Options:

`-g [guid]`<br>
`--guid [guid]`

Update the unique id for the Office Add-in. If the guid is not provided, a random guid is used.

This value is the `<Id>` element of `<OfficeApp>`.

For more info, see [OfficeApp documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/manifest/officeapp)

`-d <name>`<br>
`--displayName <name>`

Update the display name for the Office Add-in.

This value is the `<DisplayName>` element of `<OfficeApp>`. 

For more info, see [OfficeApp documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/manifest/officeapp).

#

### validate 
Determines whether the Office Add-in manifest is valid.

Syntax:

`office-addin-manifest validate <manifest>`

`manifest`: path to manifest file.

