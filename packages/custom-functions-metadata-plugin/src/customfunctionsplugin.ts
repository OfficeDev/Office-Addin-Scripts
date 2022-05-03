// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import {
  generateCustomFunctionsMetadata,
  IGenerateResult,
  MetadataOptions
} from "custom-functions-metadata";
import * as path from "path";
import { Compiler, sources, WebpackError, NormalModule } from "webpack";

/* global require */

const pluginName = "CustomFunctionsMetadataPlugin";

type Options = { input: string; output: string; options?: MetadataOptions };

class CustomFunctionsMetadataPlugin {
  private options: Options;

  constructor(options: Options) {
    this.options = options;
  }

  public static generateResults: Record<string, IGenerateResult> = {};

  public apply(compiler: Compiler) {
    const inputFilePath = path.resolve(this.options.input);
    let generateResult: IGenerateResult;

    compiler.hooks.beforeCompile.tapPromise(pluginName, async () => {
      generateResult = await generateCustomFunctionsMetadata(
        inputFilePath,
        true,
        this.options.options
      );
    });

    compiler.hooks.compilation.tap(pluginName, (compilation) => {
      if (generateResult.errors.length > 0) {
        generateResult.errors.forEach((err: string) =>
          compilation.errors.push(new WebpackError(inputFilePath + " " + err))
        );
      } else {
        compilation.assets[this.options.output] = new sources.RawSource(
          generateResult.metadataJson
        );
        CustomFunctionsMetadataPlugin.generateResults[this.options.input] =
          generateResult;
      }

      NormalModule.getCompilationHooks(compilation).beforeLoaders.tap(
        pluginName,
        (_, module) => {
          if (module.userRequest.endsWith(inputFilePath)) {
            module.loaders.push({
              loader: require.resolve("./loader.js"),
              options: {
                input: this.options.input,
              },
              ident: null,
              type: null,
            });
          }
        }
      );
    });
  }
}

export = CustomFunctionsMetadataPlugin;
