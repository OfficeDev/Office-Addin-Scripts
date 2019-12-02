// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides the writing and retrieval of application data
*/
import * as chalk from 'chalk';
import * as defaults from './defaults';
import { execSync } from "child_process";
import * as fs from 'fs';
import * as os from 'os';
import { modifyManifestFile } from 'office-addin-manifest';

export function addSecretToCredentialStore(ssoAppName: string, secret: string, isTest: boolean = false): void {
    try {
        switch (process.platform) {
            case "win32":
                console.log(`Adding application secret for ${ssoAppName} to Windows Credential Store`);
                const addSecretToWindowsStoreCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.addSecretCommandPath}" "${ssoAppName}" "${os.userInfo().username}" "${secret}"`;
                execSync(addSecretToWindowsStoreCommand, { stdio: "pipe" });
                break;
            case "darwin":
                console.log(`Adding application secret for ${ssoAppName} to Mac OS Keychain. You will need to provide an admin password to update the Keychain`);
                // Check first to see if the secret already exists i the keychain. If it does, delete it and recreate it
                const existingSecret = getSecretFromCredentialStore(ssoAppName, isTest);
                if (existingSecret !== '') {
                    const updatSecretInMacStoreCommand = `${isTest ? '' : 'sudo'} security add-generic-password -a ${os.userInfo().username} -U -s "${ssoAppName}" -w "${secret}"`;
                    execSync(updatSecretInMacStoreCommand, { stdio: "pipe" });
                } else {
                    const addSecretToMacStoreCommand = `${isTest ? '' : 'sudo'} security add-generic-password -a ${os.userInfo().username} -s "${ssoAppName}" -w "${secret}"`;
                    execSync(addSecretToMacStoreCommand, { stdio: "pipe" });
                }
                break;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }
    } catch (err) {
        throw new Error(`Unable to add secret for ${ssoAppName} to Windows Credential Store. \n${err}`);
    }
}

export function getSecretFromCredentialStore(ssoAppName: string, isTest: boolean = false): string {
    try {
        switch (process.platform) {
            case "win32":
                console.log(`Getting application secret for ${ssoAppName} from Windows Credential Store`);
                const getSecretFromWindowsStoreCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.getSecretCommandPath}" "${ssoAppName}" "${os.userInfo().username}"`;
                return execSync(getSecretFromWindowsStoreCommand, { stdio: "pipe" }).toString();
            case "darwin":
                console.log(`Getting application secret for ${ssoAppName} from Mac OS Keychain`);
                const getSecretFromMacStoreCommand = `${isTest ? "" : "sudo"} security find-generic-password -a ${os.userInfo().username} -s ${ssoAppName} -w`;
                return execSync(getSecretFromMacStoreCommand, { stdio: "pipe" }).toString();;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }

    } catch (err) {
        return '';
    }
}

function updateEnvFile(applicationId: string, port: string, envFilePath: string = defaults.envDataFilePath): void {
    console.log(`Updating ${envFilePath} with application ID and port`);
    try {
        // Update .ENV file
        if (fs.existsSync(envFilePath)) {
            const appDataContent = fs.readFileSync(envFilePath, 'utf8');

            // Check to see if the fallbackauthdialog file has already been updated and return if it has.
            if (!appDataContent.includes('{CLIENT_ID}') || !appDataContent.includes('{PORT}')) {
                console.log(chalk.yellow(`${envFilePath} has already been updated. You will need to update the CLIENT_ID and PORT settings manually`));
                return;
            }

            const updatedAppDataContent = appDataContent.replace('{CLIENT_ID}', applicationId).replace('{PORT}', port);
            fs.writeFileSync(envFilePath, updatedAppDataContent);
        } else {
            throw new Error(`${envFilePath} does not exist`)
        }
    } catch (err) {
        throw new Error(`Unable to write SSO application data to ${envFilePath}. \n${err}`);
    }
}

function updateFallBackAuthDialogFile(applicationId: string, port: string, fallbackAuthDialogPath: string = defaults.fallbackAuthDialogTypescriptFilePath, isTest: boolean = false): void {
    let isTypecript: boolean = false;
    try {
        // Update fallbackAuthDialog file
        let srcFileContent = '';
        if (fs.existsSync(defaults.fallbackAuthDialogTypescriptFilePath)) {
            srcFileContent = fs.readFileSync(defaults.fallbackAuthDialogTypescriptFilePath, 'utf8');
            fallbackAuthDialogPath = defaults.fallbackAuthDialogTypescriptFilePath;
            isTypecript = true;
        } else if (fs.existsSync(defaults.fallbackAuthDialogJavascriptFilePath)) {
            srcFileContent = fs.readFileSync(defaults.fallbackAuthDialogJavascriptFilePath, 'utf8');
            fallbackAuthDialogPath = defaults.fallbackAuthDialogJavascriptFilePath;
        } else if (isTest) {
            srcFileContent = fs.readFileSync(defaults.testFallbackAuthDialogFilePath, 'utf8');
        }
        else {
            if (isTest) {
                throw new Error(`${defaults.testFallbackAuthDialogFilePath} does not exist`);
            } else {
                throw new Error(`${isTypecript ? defaults.fallbackAuthDialogTypescriptFilePath : defaults.fallbackAuthDialogJavascriptFilePath} does not exist`)
            }

        }

        // Check to see if the fallbackauthdialog file has already been updated and return if it has.
        if (!srcFileContent.includes('{PORT}')) {
            console.log(chalk.yellow(`${fallbackAuthDialogPath} has already been updated. You will need to update the 'clientId' and 'redirectUrl' port settings manually`));
            return;
        }

        // Update fallbackauthdialog file
        console.log(`Updating ${fallbackAuthDialogPath} with application ID`);
        const updatedSrcFileContent = srcFileContent.replace('{application GUID here}', applicationId).replace('{PORT}', port);
        fs.writeFileSync(fallbackAuthDialogPath, updatedSrcFileContent);
    } catch (err) {
        if (isTest) {
            throw new Error(`Unable to write SSO application data to ${defaults.testFallbackAuthDialogFilePath}. \n${err}`);
        } else {
            throw new Error(`Unable to write SSO application data to ${isTypecript ? defaults.fallbackAuthDialogTypescriptFilePath : defaults.fallbackAuthDialogJavascriptFilePath}. \n${err}`);
        }

    }
}

async function updateProjectManifest(applicationId: string, port: string, manifestPath: string): Promise<void> {
    console.log(`Updating ${manifestPath} with applicationId and port`);
    try {
        if (fs.existsSync(manifestPath)) {
            // Update manifest with application guid and unique manifest id
            const manifestContent: string = await fs.readFileSync(manifestPath, 'utf8');

            // Check to see if the manifest has already been updated and return if it has
            if (!manifestContent.includes('{PORT}')) {
                console.log(chalk.yellow(`${manifestPath} has already been updated. You will need to update the 'WebApplicationInfo' and port settings manually`));
                return;
            }

            // Update manifest file
            const re: RegExp = new RegExp('{application GUID here}', 'g');
            const rePort = new RegExp('{PORT}', 'g');
            const updatedManifestContent: string = manifestContent.replace(re, applicationId).replace(rePort, port);
            await fs.writeFileSync(manifestPath, updatedManifestContent);
            await modifyManifestFile(manifestPath, 'random');
        } else {
            throw new Error(`${manifestPath} does not exist`)
        }
    } catch (err) {
        throw new Error(`Unable to update ${manifestPath}. \n${err}`);
    }
}

export async function writeApplicationData(applicationId: string,
    port: string,
    manifestPath: string = defaults.manifestFilePath,
    envFilePath: string = defaults.envDataFilePath,
    fallbackAuthDialogPath = defaults.fallbackAuthDialogTypescriptFilePath,
    isTest: boolean = false) {
    updateEnvFile(applicationId, port, envFilePath);
    updateFallBackAuthDialogFile(applicationId, port, fallbackAuthDialogPath, isTest);
    await updateProjectManifest(applicationId, port, manifestPath)
}
