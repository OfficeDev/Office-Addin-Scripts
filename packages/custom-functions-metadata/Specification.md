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

Syntax: @param {_type_} _name_ _description_

To denote a custom function parameter as optional:
* In JavaScript, put square brackets around _name_. For example: `@param {string} [text] Optional text` associated with `function f(text="default")`

* In TypeScript, do one of the following:
1. Use an optional parameter. For example: `function f(text?: string)`
2. Give the parameter a default value. For example: `function f(text: string = "abc")`

For detailed description of the @param see: [JSDoc](http://usejsdoc.org/tags-param.html)

#### {type}

* For JavaScript, `{type}` provides the type info for the paramater, or will be `any` if not provided.
* For TypeScript, `{type}` should be omitted as the type info will come from the Typescript parameter type.
* See the [Types](##types) for more information.

#### name

Specifies which parameter the @param tag applies to.

#### description

Provides the description which appears in Excel for the function parameter.

### @returns

Syntax: @returns {_type_}

Provides the type for the return value.

If `{type}` is omitted, the TypeScript type info will be used. If there is no type info, the type will be `any`.

### description

Provides the description of the custom functions. It is determined from the comment section above the function itself.

### @streaming

Used to indicate that a custom function is a streaming function. 

To denote the custom function is cancelable use @streaming cancelable.

The last parameter should be of type `CustomFunctions.StreamingHandler<ResultType>`.
The function should return `void`.

Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.   

### @volatile

Used to denote the custom function is volatile.

### @requiresAddress

Used to denote the custom function requires an address for the invocation.

## Types

### Value types

A single value may be represented using one of the following types: `any`, `boolean`, `number`, `string`.
Using `boolean`, `number`, or `string` will allow Excel to convert the value to the desired type before calling the function. 

### Matrix type

Use a two-dimensional array type to have the parameter or return value be a matrix of values. For example, the type `number[][]` indicates a matrix of numbers. `string[][]` indicates a matrix of strings.

## Invocation Context

### Javascript

To mark the custom function as streaming in javascript use the follow syntax in the comment section:
 
 @param {`CustomFunctions.StreamingHandler<ResultType>`} handler


## Usage

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
