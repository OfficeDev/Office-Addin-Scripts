# Office-Addin-dev-certs

Provides the ability to manage certificates for a development server using https://localhost.

## Command-Line Interface
* [install](#install)
* [verify](#verify)
* [uninstall](#uninstall)

#

### install
Creates an SSL certificate for "localhost" signed by a developer CA certificate and installs the developer CA certificate so that the certificates are trusted. If the certificates were installed but are no longer valid, they will be replaced with valid certificates.

Syntax:

`office addin-dev-certs install [options]`

Options:

`--machine`

Install the CA certificate for all users. You must be an Administrator.

`--days <days>`

Specifies the number of days until the CA certificate expires. Default: 30 days.
 
#

### verify
Verify the certificate.

Syntax:

`office addin-dev-certs verify`
 
#

### uninstall
Uninstall the certificate.

Syntax:

`office addin-dev-certs uninstall [options]`

Options:

`--machine`

Uninstall the CA certificate for all users. You must be an Administrator.

#
