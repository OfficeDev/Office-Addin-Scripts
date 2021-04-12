import { getOptions } from "loader-utils";
import { IAssociate } from "custom-functions-metadata";

module.exports = function addFunctionAssociations(this: any, source: string): string {
    const input = getOptions(this).input as string;
    if (!global.generateResults || !global.generateResults[input]) {
        return source;
    }
    global.generateResults[input].associate.forEach((item: IAssociate) => {
        source += `\nCustomFunctions.associate("${item.id}", ${item.functionName});`;
    });
    return source;
}
