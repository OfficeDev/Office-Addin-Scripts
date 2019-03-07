# Office-Addin-dev-certs

Provides the ability to start and stop debugging Office Add-ins.

## Command-Line Interface
* [generate](#generate)
* [install](#install)
* [verify](#verify)
* [uninstall](#uninstall)

#


### generate
Generate an SSL certificate for localhost and a CA certificate which has issued it.

Syntax:

`office addin-dev-certs generate [options]`

Options:

`--ca-cert <ca-cert-path>`

Path where the CA certificate file is written. Default ./ca.crt.

`--cert <cert-path>`

Path where the SSL certificate is written. Default ./localhost.crt.

`--key <key-path>`

Path where the private key for the SSL certificate is written. Default ./localhost.key.
 
#

### install
Install the certificate.

Syntax:

`office addin-dev-certs install <ca-cert-path>`

`ca-cert-path`: path to CA certificate file.

 
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
