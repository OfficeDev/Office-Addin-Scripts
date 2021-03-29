// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import readPackageJson, { ScriptsObject } from "read-package-json-fast";

let scripts: ScriptsObject | undefined;

export function clearCachedScripts() {
    scripts = undefined;
}

async function readPackageJsonScripts(jsonPath: string) {
    const packageJson = await readPackageJson(`${jsonPath}`);
    return packageJson.scripts || {};
}

/**
 * Gets a script command from package json during a npm execution
 * @param scriptName The name of the script
 */
export async function getPackageJsonScript(scriptName: string): Promise <string | undefined> {
    if (!scriptName) {
        return undefined;
    }

    // NPM v7 Support
    if (process.env.npm_package_json) {
        if (!scripts) {
            scripts = await readPackageJsonScripts(process.env.npm_package_json);
        }
        return scripts[scriptName];
    }

    // NPM v6 Support
    return process.env[`npm_package_scripts_${scriptName.replace(/-/g, "_")}`];
}
