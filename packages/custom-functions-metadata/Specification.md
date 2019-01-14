# Custom Functions Metadata Specification

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