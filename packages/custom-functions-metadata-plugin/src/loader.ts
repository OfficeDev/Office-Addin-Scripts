import { IAssociate } from "custom-functions-metadata";
import CustomFunctionsMetadataPlugin from "./customfunctionsplugin";

// Add associate calls to the functions file's source (one file)
function addFunctionAssociations(this: any, source: string): string {
  const input = this.getOptions().input as string;
  const file = this.getOptions().file as string;
  if (!CustomFunctionsMetadataPlugin.generateResults[input]) {
    return source;
  }
  const associations = CustomFunctionsMetadataPlugin.generateResults[input].associate.filter(
    (item: IAssociate) => item.sourceFileName === file
  );
  associations.forEach((item: IAssociate) => {
    source += `\nCustomFunctions.associate("${item.id}", ${item.functionName});`;
  });
  return source;
}

export = addFunctionAssociations;
