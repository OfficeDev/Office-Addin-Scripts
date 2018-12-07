import * as metadata from "custom-functions-metadata";
import * as fs from "fs";
import * as path from "path";
import * as webpack from "webpack";

const pluginName = "CustomFunctionsMetadataPlugin";

type Options = {input: string, output: string};

class CustomFunctionsPlugin {
    private options: Options;

    constructor(options: Options) {
        this.options = options;
    }

    public apply(compiler: webpack.Compiler) {

        const outputPath: string = (compiler.options && compiler.options.output) ? compiler.options.output.path || "" : "";
        const outputFilePath = path.resolve(outputPath, this.options.output);

        compiler.hooks.compile.tap(pluginName, () => {
            try {
                fs.mkdirSync(outputPath);
            } catch (err) {
                if (err.code !== "EEXIST") {
                    throw err;
                }
            }

            metadata.generate(this.options.input, outputFilePath);
        });
    }
}

module.exports = CustomFunctionsPlugin;
