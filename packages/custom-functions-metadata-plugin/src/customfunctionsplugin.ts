import * as metadata from "custom-functions-metadata";
import * as fs from "fs";
import * as path from "path";
import * as webpack from "webpack";

const pluginName = "CustomFunctionsMetadataPlugin";

type Options = {input: string, output: string};

class CustomFunctionsMetadataPlugin {
    private options: Options;

    constructor(options: Options) {
        this.options = options;
    }

    public apply(compiler: webpack.Compiler) {
        const outputPath: string = (compiler.options && compiler.options.output) ? compiler.options.output.path || "" : "";
        const outputFilePath = path.resolve(outputPath, this.options.output);
        const inputFilePath = path.resolve(this.options.input);
        let errors: string[];
        let associate: metadata.IAssociate[];

        compiler.hooks.compile.tap(pluginName, async () => {
            try {
                fs.mkdirSync(outputPath);
            } catch (err) {
                if (err.code !== "EEXIST") {
                    throw err;
                }
            }
            const generateResult = await metadata.generate(inputFilePath, outputFilePath, true);
            errors = generateResult.errors;
            associate = generateResult.associate;
        });

        compiler.hooks.emit.tap(pluginName, (compilation) => {
            if (errors.length > 0) {
                compilation.errors.push("Generating metadata file:" + outputFilePath);
                errors.forEach((err: string) => compilation.errors.push(inputFilePath + " " + err));
            } else {
                const stats = fs.statSync(outputFilePath);
                const content = fs.readFileSync(outputFilePath);
                compilation.assets[this.options.output] = {
                    source() { return content; },
                    size() { return stats.size; },
                };
            }
        });

        let functionsUpdated: boolean = false;
        compiler.hooks.compilation.tap(pluginName, (compilation, params) => {
            compilation.moduleTemplates.javascript.hooks.render.tap(pluginName, (source, module) => {
                if (!functionsUpdated && module._source && module._source._name.endsWith(inputFilePath)) {
                    associate.forEach((item) => {
                        module._source._value += `\nCustomFunctions.associate("${item.id}", ${item.functionName});`;
                      });
                    functionsUpdated = true;
                }
            });
        });

    }

}

module.exports = CustomFunctionsMetadataPlugin;
