
# Office-Addin-Scripts

These packages provide functionality which can be used to perform tasks related to Office Add-ins. The packages export functions which can be imported and used in Node scripts. Many of the packages also provide a command-line interface (CLI), allowing them to be used directly from a Command Prompt / Terminal window.

The [Yo Office](https://github.com/OfficeDev/generator-office) templates provide a starting point for developing an Office Add-in. These scripts are used in the templates to provide for basic developer tasks such as debugging. 

Developers may have other workflows with different requirements and tooling. Our goal is for the these packages to serve as building blocks which can be adapted as needed. We encourage feedback and contributions from the community.

The [Excel Custom Functions](https://github.com/OfficeDev/Excel-Custom-Functions) project provides an example of how these packages may be used.


## In this repository

* [custom-functions-metadata](packages/custom-functions-metadata/README.md)

  This package allows metadata for custom functions to be generated automatically from JSDoc tags and the function parameter types.

* [custom-functions-metadata-plugin](packages/custom-functions-metadata-plugin/README.md)

  A WebPack plugin which generates the metadata for custom functions.

* [office-addin-cli](packages/office-addin-cli/README.md)

  A command-line interface for Office Add-ins.

* [office-addin-debugging](packages/office-addin-debugging/README.md)

  This package provides the orchestration of components related to debugging Office Add-ins. When debugging is started, it will ensure that the dev-server is running, that dev settings are configured for debugging, and will register and sideload the Office Add-in. When debugging is stopped, it will unregister and shutdown components.
  
* [office-addin-dev-certs](packages/office-addin-dev-certs/README.md)

  This package can be used to manage certificates for development server using https://localhost. 

* [office-addin-dev-settings](packages/office-addin-dev-settings/README.md)

  This package can be used to configure developer settings for an Office Add-in.

* [office-addin-lint](packages/office-addin-lint/README.md)

  This package can be used to ensure code quality with lint rules and standardize code formatting.

* [office-addin-manifest](packages/office-addin-manifest/README.md)

  This package provides the ability to parse, display, and modify the manifest file for Office Add-ins.

* [office-addin-node-debugger](packages/office-addin-node-debugger/README.md)

  This package allows a Node instance to serve as a proxy for debugging a JavaScript runtime hosted by an Office application.

* [office-addin-sso](packages/office-addin-sso/README.md)

  This package provides the ability to register an application in Azure Active Directory and infrastructure for implementing single sign-on (SSO) taskpane add-ins.

* [office-addin-test-helpers](packages/office-addin-test-helpers/README.md)

  This package provides tools that make validating your Office Add-in easier. You can use it with the office-addin-test-server package and the Mocha test framework (or another testing framework of your choice).

* [office-addin-test-server](packages/office-addin-test-server/README.md)

  This package provides a framework for testing Office task pane add-ins by allowing add-ins to send results to a test server. The results can then be consumed and used by tests to validate that the add-in is working as expected.

* [office-addin-usage-data](packages/office-addin-usage-data/README.md)

  This package allows for sending usage data event and exception data to the selected telemetry infrastructure (e.g. ApplicationInsights)

## Requirements

* [Node.js](https://nodejs.org) 

## Getting started

In a command prompt, run:
* `npm install`

This should also be done when after pulling additional changes or switching branches.

## Build

To build all packages, at the root directory, run:
* `npm run build`

To build a single package, in the directory for the package, run:
* `npm run build`

## Test

To run tests for all packages, at the root directory, run:
* `npm run test`

To run tests for a single package, in the directory for the package, run:
* `npm run test`

## Editing

Use `VS Code` to edit, build, test, and debug by opening the package folder in VS Code.

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions provided by the bot. You will only need to do this once across all repos using our CLA.

## Feedback

* Ask a question on [Stack Overflow](https://stackoverflow.com/questions/tagged/office-addin-scripts).

* [File an Issue](https://github.com/OfficeDev/Office-Addin-Scripts/issues).

## Reporting Security Issues

Security issues and bugs should be reported privately, via email, to the Microsoft Security
Response Center (MSRC) at [secure@microsoft.com](mailto:secure@microsoft.com). You should
receive a response within 24 hours. If for some reason you do not, please follow up via
email to ensure we received your original message. Further information, including the
[MSRC PGP](https://technet.microsoft.com/en-us/security/dn606155) key, can be found in
the [Security TechCenter](https://technet.microsoft.com/en-us/security/default).

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## License

Code licensed under the [MIT License](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/LICENSE).

## Data usage

The Office Add-in CLI tools collect anonymized usage data and send it to Microsoft. This allows us to understand how the tools are used and how to improve them.

For more details on what we collect and how to turn it off, see our [Data usage notice](usage-data.md)
