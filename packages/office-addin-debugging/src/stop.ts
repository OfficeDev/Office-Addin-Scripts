// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { clearDevSettings, unregisterAddIn } from "office-addin-dev-settings";
import { readManifestFile } from "office-addin-manifest";
import * as debugInfo from "./debugInfo";
import { stopProcess } from "./process";
import { usageDataObject } from "./defaults";
import { ExpectedError } from "office-addin-usage-data";

/* global process, console */

export async function stopDebugging(manifestPath: string) {
  try {
    console.log("Debugging is being stopped...");

    const isWindowsPlatform = process.platform === "win32";
    const manifestInfo = await readManifestFile(manifestPath);

    if (!manifestInfo.id) {
      throw new ExpectedError("Manifest does not contain the id for the Office Add-in.");
    }

    // clear dev settings
    if (isWindowsPlatform) {
      await clearDevSettings(manifestInfo.id);
    }

    // unregister
    try {
      await unregisterAddIn(manifestPath);
    } catch (err) {
      console.log(`Unable to unregister the Office Add-in. ${err}`);
    }

    const processId = debugInfo.readDevServerProcessId();
    if (processId) {
      stopProcess(processId);
      console.log(`Stopped dev server. Process id: ${processId}`);
      debugInfo.clearDevServerProcessId();
    }

    console.log("Debugging has been stopped.");
    usageDataObject.reportSuccess("stopDebugging()");
  } catch (err) {
    usageDataObject.reportException("stopDebugging()", err);
    throw err;
  }
}
