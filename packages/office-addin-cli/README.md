# Office-Addin-CLI

A command-line interface for Office Add-ins.
This package provides the ability to upgrade your Office Addin code.

## Command-Line Interface

* [upgrade](#upgrade)

#

### convert

Converts the Office Addin Code.

Syntax:

`office-addin-cli convert [options]`

`manifest-path`: path to manifest file.

Options:

`-m <manifest-path>`<br>
`--manifest <manifest-path>`

Specify the location of the manifest file. If the path is not provided, `./manifest.xml` is used.

`-p <packageJson-path>`<br>
`--packageJson <packageJson-path>`

Specify the location of the manifest file. If the path is not provided, `./package.json` is used.
