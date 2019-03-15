# Office-Addin-dev-certs

Provides the ability to manage certificates for a development server using https://localhost.

## Command-Line Interface
* [install](#install)
* [verify](#verify)
* [uninstall](#uninstall)

#

### install
If not installed, generate an SSL certificate for localhost and a CA certificate which has issued it and install it.

Syntax:

`office addin-dev-certs install [options]`

Options:

`--ca-cert <ca-cert-path>`

Path where the CA certificate file is written. Default ./ca.crt.

`--cert <cert-path>`

Path where the SSL certificate is written. Default ./localhost.crt.

`--key <key-path>`

Path where the private key for the SSL certificate is written. Default ./localhost.key.

`--days <days>`

Specifies the validity of CA certificate in days.
 
#

### verify
Verify the certificate.

Syntax:

`office addin-dev-certs verify`
 
#

### uninstall
Uninstall the certificate.

Syntax:

`office addin-dev-certs uninstall`

#
