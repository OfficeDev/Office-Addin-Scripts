# Office-Addin-Debugging

  This package provides the orchestration of components related to debugging Office Add-ins. When debugging is started, it will ensure that the dev-server is running, that dev settings are configured for debugging, and will register and sideload the Office Add-in. When debugging is stopped, it will unregister and shutdown components.

## Command-Line Interface
* [start](#start)
* [stop](#stop)

#

## start 
Starts debugging. 

Syntax:

`office addin-debugging start <manifest> [platform] [options]`

`manifest`: path to manifest file.

`platform`: which type of application:
* `desktop`: Office app for Windows or Mac
* `web`: Office app in the web browser

Notes:

* The dev server is needed to download the add-in from the source location specified in the manifest.

* `--packager` is needed unless `--debug-method` is `direct` and `--no-live-reload` is specified.

Options:

`--app <app`>

Specifies which Office app to use:
* `excel`
* `onenote`
* `outlook`
* `project`
* `powerpoint`
* `word`

If this is not specified, the behavior depends on the `<Hosts>` specified in the manifest. 
For a single host, it will automatically start that Office app.
For multiple hosts, it will prompt to choose the desired host.

`--debug-method <method>`

Specifies which debug method to use: 
* `direct`: debug directly using the JavaScript engine.
* `web`: debug using the JavaScript engine in a web browser or Node.
 
`--dev-server <command>`

Specifies to run the dev server using the specified command.

`--dev-server-port <port>`

Specifies the port for the dev server. If provided, the dev server is only started if not already running. 

`--document`

Specifies the document to sideload.  The document option can either be the local path to a document or a url.

` --no-debug`

Start without debugging.

` --no-live-reload`

Do not enable live-reload.

` --no-sideload`

Do not start the Office app and load the Office add-in.

` --packager <command>`

If this option is provided, the packager is started with the specified command.

` --packager-host <host>`

Host name of the packager. Default: `localhost`.

` --packager-port <port>`

Port number of the packager. Default: `8081`.

` --prod`

Specifies that debugging session is for production mode. Default is development mode.

` --source-bundle-url-host <host>`

Host name to obtain the source bundle. Default: `localhost`.

` --source-bundle-url-port <port>`

Port number to obtain the source bundle. Default: `8081`.

` --source-bundle-url-path <path>`

Path used to obtain the source bundle. 

` --source-bundle-url-extension <extension>`

Extension used to obtain the source bundle. Default: `.bundle`.

`-h, --help`

Output usage information.

#

### stop
Stops debugging.

Syntax:

`office addin-debugging stop <manifest> [options]`

`manifest`: path to manifest file.

Options:

` --prod`

Specifies that debugging session is for production mode. Default is development mode.

#

