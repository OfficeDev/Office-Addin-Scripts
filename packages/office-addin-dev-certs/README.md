# Office-Addin-dev-certs

This package can be used to manage certificates for development server using https://localhost. 

## Installation

```
npm install office-addin-dev-certs
```

Upon installation a development CA certicate and localhost key and
certificate will be generated inside `<userhome>/.office-addin-dev-certs`. 
The certificate is valid for 30 days by default.

#

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

## API Usage

```js
var https = require('https')
var devCerts = require("office-addin-dev-certs");
var options =  devCerts.getHttpsServerOptions();

var server = https.createServer(options, function (req, res) {
  res.end('This is servered over HTTPS')
})

server.listen(443, function () {
  console.log('The server is running on https://localhost:443')
})
```
