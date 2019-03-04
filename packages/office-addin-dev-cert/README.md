# Office-Addin-Debugging

Provides the ability to start and stop debugging Office Add-ins.

## Command-Line Interface
* [generate](#generate)
* [install](#install)
* [verify](#verify)
* [uninstall](#uninstall)
* [clean](#clean)

#


### generate
Generate self-signed ca and localhost certificate.

Syntax:

`office addin-dev-cert generate <manifest> [options]`

`manifest`: path to manifest file.

Options:

`--path <command>`

Optional path of generated certficate files. By default, certificates are generated in .certs directory.
 
#

### install
Install the certificate.

Syntax:

`office addin-dev-cert install <manifest>`

`manifest`: path to manifest file.
 
#

### verify
Verify the certificate.

Syntax:

`office addin-dev-cert verify <manifest>`

`manifest`: path to manifest file.
 
#

### uninstall
Uninstall the certificate.

Syntax:

`office addin-dev-cert uninstall`

#

### clean
Clean the certificates.

Syntax:

`office addin-dev-cert clean <manifest>`

`manifest`: path to manifest file.
 
#