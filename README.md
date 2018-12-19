
# Office-Addin-Scripts

These packages provide functionality which can be used to automate tasks related to Office Add-ins.

These packages have been developed initially for debugging Office Add-ins which use the JavaScript runtime directly. You can see how these are used in the [Excel Custom Functions](https://github.com/OfficeDev/Excel-Custom-Functions) project.

Developers may have their own workflow and tooling. These packages provide basic building blocks which can be adapted as needed. We encourage contributions and feedback.

## In this repository

* [office-addin-debugging](packages/office-addin-debugging/README.md)
* [office-addin-dev-settings](packages/office-addin-dev-settings/README.md)
* [office-addin-manifest](packages/office-addin-manifest/README.md)
* [office-addin-node-debugger](packages/office-addin-node-debugger/README.md)

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

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## License

Code licensed under the [MIT License](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/LICENSE).