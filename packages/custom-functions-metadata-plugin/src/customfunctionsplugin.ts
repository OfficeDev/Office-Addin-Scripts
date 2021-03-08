// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as metadata from "custom-functions-metadata";
import * as fs from "fs";
import * as path from "path";
import { Compiler, sources, WebpackError } from "webpack";

const pluginName = "CustomFunctionsMetadataPlugin";

type Options = {input: string, output: string};

class CustomFunctionsMetadataPlugin {
    private options: Options;

    constructor(options: Options) {
        this.options = options;
    }

    public apply(compiler: Compiler) {
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
                compilation.errors.push(new WebpackError("Generating metadata file:" + outputFilePath));
                errors.forEach((err: string) => compilation.errors.push(new WebpackError(inputFilePath + " " + err)));
            } else {
                const content = fs.readFileSync(outputFilePath);
                compilation.emitAsset(outputFilePath, new sources.RawSource(content));
            }
            
        });

        compiler.hooks.compilation.tap(pluginName, (compilation, params) => {
            compilation.moduleTemplates.javascript.hooks.render.tap(pluginName, (source : any, module : any) => {
                if (module._source && module._source._name.endsWith(inputFilePath)) {
                    associate.forEach((item) => {
                        module._source._value += `\nCustomFunctions.associate("${item.id}", ${item.functionName});`;
                    });
                }
            });
        });
    }
}

module.exports = CustomFunctionsMetadataPlugin;
