# Custom Functions Metadata Specification

## Overview

When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide the extra information about the custom function. 

## Tags

### @customfunction

Syntax: @customfunction _id_ _name_

Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.

This tag is required to auto generated the metadata for the custom function.

#### id 

The id is used as the invariant identifier for the custom function stored in the document. It should not change.

* If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.
* The id must be unique for all custom functions.
* The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).

#### name

Provides the display name for the custom function. 

* If name is not provided, the id is also used as the name.
* Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).
* Must start with a letter.
* Maximum length is 128 characters.

### @helpurl

Syntax: @helpurl _url_

The provided _url_ is displayed in Excel.

### @param

Syntax: @param {_type_} [_name_] _description_

Provides the type and description for the parameter named _name_. 

If `{type}` is omitted, the TypeScript type info will be used. If there is no type info, the type will be `any`.

[] around the name denote it is an optional parameter.

See the [Types](##types) for more information.


### @returns

Syntax: @returns {_type_} description

Provides the type and description for the return value.

If `{type}` is omitted, the TypeScript type info will be used. If there is no type info, the type will be `any`.

### description

Provides the description of the custom functions. It is determined from the comment section above the function itself.

### @streaming cancelable

Used to denote custom function is streaming and cancelable. In typescript if the last parameter of the function is of type CustomFunctions.StreamingHandler, then the function will be marked streaming and cancelable.

### @volatile

Used to denote the custom function is volatile.


## Types

Support for the following types: string, number, boolean, and any.
All other typescript types will be treated as any.

## Invocation Context


Command to run tool:
Npm run generate-json [inputFile] [output]

## Example:

/**
 * This function adds 2 or 3 numbers together
 * @CustomFunction
 * @param {number} first - the first number
 * @param {number} second - the second number
 * @param {number} [third] - the third optional number
 * @helpUrl https://dev.office.com
 * @volatile
 * @streaming cancelable
 * @return {number}
  */

function add(first: number, second: number, third?: number): number
{

    return first + second + third;

}

Metadata generated:

{
    "functions":[
        
        {
            "description": "This function adds 2 or 3 numbers together",
            "helpUrl": "https://dev.office.com",
            "id": "add",
            "name": "ADD",
            "options": {
                "cancelable": true,
                "stream": true,
                "volatile": true
            },
            "parameters": [
                {
                    "description": "the first number",
                    "name": "first",
                    "optional": false,
                    "type": "number"
                },
                {
                    "description": "the second number",
                    "name": "second",
                    "optional": false,
                    "type": "number"
                },
                {
                    "description": "the third optional number",
                    "name": "third",
                    "optional": true,
                    "type": "number"
                }
            ],
            "result": {
                "type": "number"
            }
        }
    ]
}
