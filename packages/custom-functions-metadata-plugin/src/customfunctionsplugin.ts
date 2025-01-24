// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { generateCustomFunctionsMetadata, IGenerateResult } from "custom-functions-metadata";
import path from "path";
import { Compiler, Compilation, sources, WebpackError, NormalModule } from "webpack";

/* global require */

const pluginName = "CustomFunctionsMetadataPlugin";

type Options = { input: string; output: string };

class CustomFunctionsMetadataPlugin {
  private options: Options;

  constructor(options: Options) {
    this.options = options;
  }

  public static generateResults: Record<string, IGenerateResult> = {};

  public apply(compiler: Compiler) {
    let input: string[] = Array.isArray(this.options.input)
      ? this.options.input
      : [this.options.input];
    let generateResult: IGenerateResult;
    input = input.map((file) => path.resolve(file));

    compiler.hooks.beforeCompile.tapPromise(pluginName, async () => {
      generateResult = await generateCustomFunctionsMetadata(input, true);
    });

    compiler.hooks.compilation.tap(pluginName, (compilation: Compilation) => {
      if (generateResult.errors.length > 0) {
        generateResult.errors.forEach((err: string) =>
          compilation.errors.push(new WebpackError(input + " " + err))
        );
      } else {
        compilation.assets[this.options.output] = new sources.RawSource(
          generateResult.metadataJson
        );
        CustomFunctionsMetadataPlugin.generateResults[this.options.input] = generateResult;
      }

      // trigger the loader to add code to the functions files
      NormalModule.getCompilationHooks(compilation).beforeLoaders.tap(pluginName, (_, module) => {
        const found = input.find((item) => module.userRequest.endsWith(item));
        if (found) {
          module.loaders.push({
            loader: require.resolve("./loader.js"),
            options: {
              input: this.options.input,
              file: found,
            },
            ident: null,
            type: null,
          });
        }
      });
    });
  }
}

export = CustomFunctionsMetadataPlugin;
