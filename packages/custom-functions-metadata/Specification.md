# Custom Functions Metadata Specification

## Overview

When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide the extra information about the custom function. 

## Tags

### @customfunction

Syntax: @customfunction _id_ _name_

Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.

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

Syntax: @param {_type_} name _description_

Provides the type and description for the parameter named _name_. 

If `{type}` is omitted, the TypeScript type info will be used. If there is no type info, the type will be `any`.

See the [Types](##types) for more information.


### @returns

Syntax: @returns {_type_} description

Provides the type and description for the return value.

If `{type}` is omitted, the TypeScript type info will be used. If there is no type info, the type will be `any`.


## Types


## Invocation Context






@CustomFunction required

Id: function name
Name: function name

/**
* This is the description
* @CustomFunction
*/
Description: Comments section of function

helpUrl: @helpUrl in comments section

result:{
	type: number, string, boolean, or any 
	dimensionality: scalar or matrix
}
@return {type} used for result type in javascript. 

Only adding dimensionality if matrix

Parameters: {
	Description: @param {type} parameterName – Description of parameter
	Name: populated from the function signature
	Optional: In typescript this is denoted by parameterName? Or in javascript use @param {type} [parameterName] 
	Type: @param {type} if not determined by function signature
    Dimensionality: scalar or matrix
}

Options: {
	Cancelable: @streaming cancelable 
	Stream: @streaming or last parameter of function is of type CustomFunctions.StreamingHandler<type>
	Volatile: @volatile
}

Only writing options section if one of the parameters is true.

Command to run tool:
Npm run generate-json <inputFile> <output>

Example:
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
function add(first: number, second: number, third?: number): number {
    return first + second + third;
}

Metadata generated:
{
    "functions": [
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
