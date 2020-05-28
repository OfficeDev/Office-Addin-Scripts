# Office-Addin-Dev-Settings

Provides the ability to configure developer settings for Office Add-ins.

## Command-Line Interface
* [appcontainer](#appcontainer)
* [clear](#clear)
* [debugging](#debugging)
* [live-reload](#live-reload)
* [register](#register)
* [registered](#registered)
* [runtime-log](#runtime-log)
* [sideload](#sideload)
* [source-bundle-url](#source-bundle-url)
* [unregister](#unregister)
* [webview](#webview)

#

### appcontainer 
Display or configure settings related to the appcontainer for an Office Add-in. 

Syntax:

`office addin-dev-settings appcontainer <manifest> [options]`

`manifest`: path to manifest file.

Without options, displays the appcontainer name.

Notes:

* Without options, displays the appcontainer name and whether access to localhost is allowed.
* The appcontainer must be registered in order to allow access to loopback addresses.

Options:

`--loopback`

Allow access to loopback addresses such as `localhost`.
 
`--prevent-loopback`

Prevent access to loopback addresses such as `localhost`.

#

### clear
Clear developer settings for the Office Add-in.

Syntax:

`office addin-debugging clear <manifest>`

`manifest`: path to manifest file.

# 

### debugging 
Display or configure debugging settings for an Office Add-in. 

Syntax:

`office addin-dev-settings debugging <manifest> [options]`

`manifest`: path to manifest file. 

Without options, displays whether debugging is enabled.

Notes:

These settings do not apply when the Office Add-in runs in a web browser or WebView control.

Options:

`--disable`

Disable debugging for the Office Add-in.

`--enable`

Enable debugging for the Office Add-in.
 
`--debug-method <method>`

Specifies which debug method to use: 
* `direct`: debug directly using the JavaScript engine.
* `web`: debug using the JavaScript engine in a web browser or Node.

#

### live-reload 
Display or configure settings related to live reload for an Office Add-in. 

Syntax:

`office addin-dev-settings live-reload <manifest> [options]`

`manifest`: path to manifest file. 

Without options, displays whether live reload is enabled.

Options:

`--disable`

Disable live-reload for the Office Add-in.

`--enable`

Enable live-reload for the Office Add-in.
 
#

### register 
Registers an Office Add-in for development. 

Syntax:

`office addin-dev-settings register <manifest> [options]`

`manifest`: path to manifest file. 

#

### registered 
Displays the Office Add-ins registered for development. 

Syntax:

`office addin-dev-settings registered [options]`

#

### runtime-log 
 Use the command to enable or disable writing any Office Add-in runtime events to a log file. Without options, it displays whether runtime logging is enabled.

Notes:

The setting is not specific to a particular Office Add-in. It applies to the runtime and will show information for all Office Add-ins. 

Syntax:

`office addin-dev-settings runtime-log [options]`

Without options, displays whether runtime logging is enabled and the log file path (if enabled).

Options:

`--disable`

Disable runtime logging.

`--enable [path]`

Enable runtime logging.

* `path`: Specify the path to the log file. If not specified, uses "OfficeAddins.log.txt" in the TEMP folder.
 
#

### sideload 
Start Office and open a document so the Office Add-in is loaded. 

Syntax:

`office addin-dev-settings sideload <manifest> [options]`

`manifest`: path to manifest file. 

Note:

If the add-in supports more than one Office app, the command will prompt to choose the app unless the `--app` parameter is provided.  

Options:

`-a`
`--app`

Specify the Office application to load.

#

### source-bundle-url 
Configure the url used to obtain the source bundle from the packager for an Office Add-in.

The url is composed as:

http://`HOST`:`PORT`/`PATH` `EXTENSION`

* `HOST`: host name; default is `localhost`
* `PORT`: port number; default is `8081` 
* `PATH`: path
* `EXTENSION`: extension (including period); default is `.bundle`

Syntax:

`office addin-dev-settings source-bundle-url [options]`

Without options, displays the current source-bundle-url settings.

Options:

`-h <name>`<br>
`--host <name>`

Specify the host name or "" to use the default.

`-p <number>`<br>
`--port <number>`

Specify the port number (0 to 65535) or "" to use the default.

`--path <path>`

Specify the path or "" to use the default.

`-e <string>`<br>
`--extension <string>`

Specify the extension (which should start with a period) or "" to use the default.
 
#

### unregister 
Unregisters an Office Add-in for development. 

Syntax:

`office addin-dev-settings register <manifest> [options]`

`manifest`: path to manifest file. 

#

### webview 
Switches the webview runtime in Office for testing and development scenarios.

> Office must be running on the Beta channel to switch runtimes

Syntax:

`office addin-dev-settings webview <manifest> <runtime>`

`manifest`: path to manifest file. 

`runtime`: Office runtime to load (currently accepts 'ie', 'edge', or 'default'). 

#
