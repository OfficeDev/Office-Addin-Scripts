// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

const readJson = require('read-package-json-fast');

let jsonScripts : any = undefined;

async function parseJsonPackageToJsonScripts(jsonPath: string) {
    return (await readJson(`${jsonPath}`)).scripts;
}

/**
 * Gets a script command from package json during a npm execution
 * @param scriptName The name of the script
 */
export async function getPackageJsonScript(scriptName: string): Promise <string | undefined> {
    if (!scriptName) {
        return undefined;
    }

    // NPMv7 Support
    if (process.env.npm_package_json) {
        if (!jsonScripts) {
            jsonScripts = await parseJsonPackageToJsonScripts(process.env.npm_package_json);
        }
        return jsonScripts[scriptName];
    }

    // NPMv6 Support
    scriptName.replace(/-/g, "_"); // Replacing all hyphens with underscores
    return process.env["npm_package_scripts_" + scriptName];
}
