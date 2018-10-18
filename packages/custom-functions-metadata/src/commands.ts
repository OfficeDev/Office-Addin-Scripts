
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as customjson from "./custom-functions-metadata";

export async function generate(inputfile: string, outputfile: string) {
  try {
      console.log("Begin json generation")
      customjson.generate(inputfile,outputfile);
  }
  catch (err){
    console.log('Error: ${err}');
  }
}
