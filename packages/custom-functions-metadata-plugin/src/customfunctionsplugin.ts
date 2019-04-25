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
        let errors: string[];

        compiler.hooks.compile.tap(pluginName, async () => {
            try {
                fs.mkdirSync(outputPath);
            } catch (err) {
                if (err.code !== "EEXIST") {
                    throw err;
                }
            }
            const generateResult = await metadata.generate(this.options.input, outputFilePath, true);
            errors = generateResult.errors;
        });

        compiler.hooks.emit.tap(pluginName, (compilation) => {
            if (errors.length > 0) {
                compilation.errors.push("Generating metadata file:" + outputFilePath);
                errors.forEach((err: string) => compilation.errors.push(this.options.input + " " + err));
            } else {
                const stats = fs.statSync(outputFilePath);
                const content = fs.readFileSync(outputFilePath);
                compilation.assets[this.options.output] = {
                    source() { return content; },
                    size() { return stats.size; },
                };
            }
        });

        compiler.hooks.compilation.tap(pluginName, (compilation, params) => {
            compilation.moduleTemplates.javascript.hooks.render.tap(pluginName, () => {
                // do something here to emit the calls
                // compilation.moduleTemplates.javascript.hooks.content.call()
                let check: string;
                compilation.modules.forEach((m) => {
                    check = m._source._name;
                    if (check && check.endsWith("functions.ts") ) {
                        console.log(m._source._value);
                        m._source._value += 'CustomFunctions.associate("LOG2", logMessage);';
                    }
                });
            });
        });

    }

}

module.exports = CustomFunctionsMetadataPlugin;
