// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { generateCustomFunctionsMetadata, IGenerateResult } from "custom-functions-metadata";
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
        let generateResult: IGenerateResult;

        compiler.hooks.compile.tap(pluginName, async () => {
            generateResult = await generateCustomFunctionsMetadata(inputFilePath, true);
        });

        compiler.hooks.emit.tap(pluginName, (compilation) => {
            if (generateResult.errors.length > 0) {
                generateResult.errors.forEach((err: string) => compilation.errors.push(new WebpackError(inputFilePath + " " + err)));
            } else {
                compilation.assets[this.options.output] = new sources.RawSource(generateResult.metadataJson);
            }
        });

        compiler.hooks.compilation.tap(pluginName, (compilation, params) => {
            compilation.moduleTemplates.javascript.hooks.render.tap(pluginName, (source : any, module : any) => {
                if (module._source && module._source._name.endsWith(inputFilePath)) {
                    generateResult.associate.forEach((item: { id: string; functionName: string; }) => {
                        module._source._value += `\nCustomFunctions.associate("${item.id}", ${item.functionName});`;
                    });
                }
            });
        });
    }
}

module.exports = CustomFunctionsMetadataPlugin;
