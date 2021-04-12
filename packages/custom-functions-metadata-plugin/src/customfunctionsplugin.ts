// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { generateCustomFunctionsMetadata, IGenerateResult } from "custom-functions-metadata";
import * as fs from "fs";
import * as path from "path";
import { Compiler, sources, WebpackError, NormalModule } from "webpack";

const pluginName = "CustomFunctionsMetadataPlugin";

type Options = {input: string, output: string};

declare global {
    namespace NodeJS {
        interface Global {
            generateResults: Record<string, IGenerateResult> | undefined;
        }
    }
}

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

        compiler.hooks.beforeCompile.tapPromise(pluginName, async () => {
            generateResult = await generateCustomFunctionsMetadata(inputFilePath, true);
        });

        compiler.hooks.compilation.tap(pluginName, (compilation) => {
            if (generateResult.errors.length > 0) {
                generateResult.errors.forEach((err: string) => compilation.errors.push(new WebpackError(inputFilePath + " " + err)));
            } else {
                compilation.assets[this.options.output] = new sources.RawSource(generateResult.metadataJson);
                if (!global.generateResults) {
                    global.generateResults = {};
                }
                global.generateResults[this.options.input] = generateResult;
            }

            NormalModule.getCompilationHooks(compilation).beforeLoaders.tap(pluginName, (_, module) => {
                if (module.userRequest.endsWith(inputFilePath)) {
                    module.loaders.push({
                        loader: require.resolve("./loader.js"),
                        options: {
                            input: this.options.input,
                        },
                        ident: null,
                        type: null,
                    })
                }
            });
        });
    }
}

module.exports = CustomFunctionsMetadataPlugin;
