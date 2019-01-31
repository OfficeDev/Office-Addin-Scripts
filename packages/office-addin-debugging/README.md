# Office-Addin-Debugging

Provides the ability to start and stop debugging Office Add-ins.

## Command-Line Interface
* [start](#start)
* [stop](#stop)

#

## start 
Starts debugging. 

Syntax:

`office addin-debugging start <manifest> [options]`

`manifest`: path to manifest file.

Notes:

* The dev server is needed to download the add-in from the source location specified in the manifest.

* `--packager` is needed unless `--debug-method` is `direct` and `--no-live-reload` is specified.

* `--sideload` is needed to open an Office document with the add-in. The command would typically use `office-toolbox sideload -m <manifest> -a <app>`.

Options:

`--debug-method <method>`

Specifies which debug method to use: 
* `direct`: debug directly using the JavaScript engine.
* `web`: debug using the JavaScript engine in a web browser or Node.
 
`--dev-server <command>`

Specifies to run the dev server using the specified command.

`--dev-server-port <port>`

Specifies the port for the dev server. If provided, the dev server is only started if not already running. 

` --no-debug`

Start without debugging.

` --no-live-reload`

Do not enable live-reload.

` --packager <command>`

If this option is provided, the packager is started with the specified command.

` --packager-host <host>`

Host name of the packager. Default: `localhost`.

` --packager-port <port>`

Port number of the packager. Default: `8081`.

` --sideload <command>`

Load the add-in using the specified command.

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

`--unload <command>`

Unregister the Office Add-in using the specified command. For example: `office-toolbox remove -m <manifest> -a <app>`.
 
#

