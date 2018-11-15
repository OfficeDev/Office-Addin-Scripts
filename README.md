
# <img src="./assets/Office-Dev-cp.png" width="83" height="68" /> Office Add-in Scripts
<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
When developing on the web, developers have their own workflow and tooling as part of their inner dev loop. Enforcing too many requirements creates confusion and forces developers to follow specific rules that may slow down the development process.
<br /><br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Office Addin Scripts are a set of scripts and tools that provide functionality that can be composed into repeatable, reliable and productive workflows to remove any impedance on dev inner loops.

# Packages
- **packages/office-addin-debugging:**
  - The office-addin-debugging package provides a set of scripts that assist in bootstrapping the debug workflow. This inclides booting a JS runtime and enabling the [Chrome DevTools Protocol (CDP)](https://chromedevtools.github.io/devtools-protocol/) in the runtime to allow for the connection of a debugger.
    - These scripts are not opinionated one which debugger/IDE is being used, and allows for Chrome, FireFox, [VSCode](https://code.visualstudio.com/), or [Visual Studio](https://github.com/Microsoft/nodejstools/wiki/Debugging) to connect to the process for debugging.
<br />

- **packages/office-addin-dev-settings:**
  - The office-addin-dev-settings provides scripting and a CLI to enable Dev mode in Office Hosts and surfacing dev mode settings that can be customized for each session. 
    - These scripts are not opinionated one which debugger/IDE is being used, and allows for Chrome, FireFox, [VSCode](https://code.visualstudio.com/), or [Visual Studio](https://github.com/Microsoft/nodejstools/wiki/Debugging) to connect to the process for debugging.
    - Some Example Questions this package can assist:
        - Enable debugging?
         - What port should [(CDP)](https://chromedevtools.github.io/devtools-protocol/) be enabled on?
            - Example: chrome-devtools://devtools/bundled/inspector.html?experiments=true&ws=localhost:40000
         - What is the location of an Addin's manifest for sideloading?
         - Should logging/tracing be enabled for your add-in?
         - etc...
<br />

- **packages/office-addin-manifest:**
  - The office-addin-manifest pacakge is an advanced manifest parser and SDK. The tool understands the manifest schema and provides tooling to update nodes and attributes in the manifest programmatically. 
<br />

- **packages/office-addin-node-debugger:**
  - JavaScript runtimes (JSRs), [that may be embedded](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime), execute JavaScript as the BL, can sometimes prevent the ability to remotely debug the JavaScript running. This package assists in the ability to execute JavaScript on a Node Proxy service in order to allow for other debuggers/IDEs to attach to it locally and debug.

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
