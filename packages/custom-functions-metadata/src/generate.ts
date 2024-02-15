import { IFunction, IGenerateResult, IParseTreeResult, parseTree } from "./parseTree";
import { existsSync, readFileSync } from "fs";

/**
 * Generate the metadata of the custom functions
 * @param inputFile - File that contains the custom functions
 * @param outputFileName - Name of the file to create (i.e functions.json)
 */
export async function generateCustomFunctionsMetadata(
    input: string | string[],
    wantConsoleOutput: boolean = false
  ): Promise<IGenerateResult> {
    const inputFiles: string[] = Array.isArray(input) ? input : [input];
    const functions: IFunction[] = [];
    const generateResults: IGenerateResult = {
      metadataJson: "",
      associate: [],
      errors: [],
    };
  
    if (input && inputFiles.length > 0) {
      inputFiles.forEach((inputFile) => {
        inputFile = inputFile.trim();
        if (!inputFile) {
          // ignore empty strings
        } else if (!existsSync(inputFile)) {
          throw new Error(`File not found: ${inputFile}`);
        } else {
          const sourceCode = readFileSync(inputFile, "utf-8");
          const parseTreeResult: IParseTreeResult = parseTree(sourceCode, inputFile);
          parseTreeResult.extras.forEach((extra) => extra.errors.forEach((err) => generateResults.errors.push(err)));
  
          if (generateResults.errors.length > 0) {
            if (wantConsoleOutput) {
              console.error("Errors in file: " + inputFile);
              generateResults.errors.forEach((err) => console.error(err));
            }
          } else {
            functions.push(...parseTreeResult.functions);
            generateResults.associate.push(...parseTreeResult.associate);
          }
        }
      });
  
      if (functions.length > 0) {
        const metadata = {
          allowCustomDataForDataTypeAny: true,
          functions: functions,
        }
        generateResults.metadataJson = JSON.stringify(metadata, null, 4);
      }
    }
  
    return generateResults;
  }
  
  