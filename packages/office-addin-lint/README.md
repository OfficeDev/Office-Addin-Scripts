# Office-Addin-Lint

Ensure code quality with lint rules and consistent code formatting.

## Command-Line Interface
* [check](#check)
* [fix](#fix)
* [prettier](#prettier)

#

### check 
Check the source code for problems.

Syntax:

`office addin-lint check [options]`

Options:

`--files <files>`

Specify the files to check. Default: `src/**/*.{ts,tsx,js,jsx}`
 
#

### fix 
Apply fixes for problems found in the source code.

Syntax:

`office addin-lint fix [options]`

Options:

`--files <files>`

Specify the files to fix. Default: `src/**/*.{ts,tsx,js,jsx}`
 
#

### prettier 
Make the source code prettier.

Syntax:

`office addin-lint prettier [options]`

Options:

`--files <files>`

Specify the files to fix. Default: `src/**/*.{ts,tsx,js,jsx}`
 
#

## Package installation
Steps to follow when installing the package and configuring for use.

Install the following packages:

`npm install -D office-addin-lint`

`npm install -D office-addin-prettier-config`

#

Add the following to package.json:

    Under scripts add the following to enable the 3 actions:

        "lint": "office-addin-lint check",
        "lint:fix": "office-addin-lint fix",
        "prettier": "office-addin-lint prettier"`

    At top level add the following to enable the prettier config:

        "prettier": "office-addin-prettier-config"
#
