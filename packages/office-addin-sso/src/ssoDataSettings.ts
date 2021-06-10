// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides the writing and retrieval of application data
*/
import * as defaults from "./defaults";
import { execSync } from "child_process";
import * as fs from "fs";
import * as os from "os";
import { modifyManifestFile } from "office-addin-manifest";
import { ExpectedError } from "office-addin-usage-data";

/* global process */

export function addSecretToCredentialStore(ssoAppName: string, secret: string, isTest: boolean = false): void {
  try {
    switch (process.platform) {
      case "win32": {
        const addSecretToWindowsStoreCommand = `powershell -ExecutionPolicy Bypass -File "${
          defaults.addSecretCommandPath
        }" "${ssoAppName}" "${os.userInfo().username}" "${secret}"`;
        execSync(addSecretToWindowsStoreCommand, { stdio: "pipe" });
        break;
      }
      case "darwin": {
        // Check first to see if the secret already exists i the keychain. If it does, delete it and recreate it
        const existingSecret = getSecretFromCredentialStore(ssoAppName, isTest);
        if (existingSecret !== "") {
          const updatSecretInMacStoreCommand = `${isTest ? "" : "sudo"} security add-generic-password -a ${
            os.userInfo().username
          } -U -s "${ssoAppName}" -w "${secret}"`;
          execSync(updatSecretInMacStoreCommand, { stdio: "pipe" });
        } else {
          const addSecretToMacStoreCommand = `${isTest ? "" : "sudo"} security add-generic-password -a ${
            os.userInfo().username
          } -s "${ssoAppName}" -w "${secret}"`;
          execSync(addSecretToMacStoreCommand, { stdio: "pipe" });
        }
        break;
      }
      default: {
        const errorMessage: string = `Platform not supported: ${process.platform}`;
        throw new ExpectedError(errorMessage);
      }
    }
  } catch (err) {
    const errorMessage: string = `Unable to add secret to Credential Store: \n${err}`;
    throw new Error(errorMessage);
  }
}

export function getSecretFromCredentialStore(ssoAppName: string, isTest: boolean = false): string {
  try {
    switch (process.platform) {
      case "win32": {
        const getSecretFromWindowsStoreCommand = `powershell -ExecutionPolicy Bypass -File "${
          defaults.getSecretCommandPath
        }" "${ssoAppName}" "${os.userInfo().username}"`;
        return execSync(getSecretFromWindowsStoreCommand, {
          stdio: "pipe",
        }).toString();
      }
      case "darwin": {
        const getSecretFromMacStoreCommand = `${isTest ? "" : "sudo"} security find-generic-password -a ${
          os.userInfo().username
        } -s ${ssoAppName} -w`;
        return execSync(getSecretFromMacStoreCommand, {
          stdio: "pipe",
        }).toString();
      }
      default: {
        const errorMessage: string = `Platform not supported: ${process.platform}`;
        throw new ExpectedError(errorMessage);
      }
    }
  } catch {
    return "";
  }
}

function updateEnvFile(applicationId: string, port: string, envFilePath: string = defaults.envDataFilePath): boolean {
  try {
    // Update .ENV file
    if (fs.existsSync(envFilePath)) {
      const appDataContent = fs.readFileSync(envFilePath, "utf8");
      // Check to see if the fallbackauthdialog file has already been updated and return if it has.
      if (!appDataContent.includes("{CLIENT_ID}") || !appDataContent.includes("{PORT}")) {
        return false;
      }

      const updatedAppDataContent = appDataContent.replace("{CLIENT_ID}", applicationId).replace("{PORT}", port);
      fs.writeFileSync(envFilePath, updatedAppDataContent);
      return true;
    } else {
      throw new ExpectedError(`${envFilePath} does not exist`);
    }
  } catch (err) {
    const errorMessage: string = `Unable to write SSO application data to .env file: \n${err}`;
    throw new Error(errorMessage);
  }
}

function updateFallBackAuthDialogFile(
  applicationId: string,
  port: string,
  fallbackAuthDialogPath: string = defaults.fallbackAuthDialogTypescriptFilePath,
  isTest: boolean = false
): boolean {
  let isTypecript: boolean = false;
  try {
    // Update fallbackAuthDialog file
    let srcFileContent = "";
    if (fs.existsSync(defaults.fallbackAuthDialogTypescriptFilePath)) {
      srcFileContent = fs.readFileSync(defaults.fallbackAuthDialogTypescriptFilePath, "utf8");
      fallbackAuthDialogPath = defaults.fallbackAuthDialogTypescriptFilePath;
      isTypecript = true;
    } else if (fs.existsSync(defaults.fallbackAuthDialogJavascriptFilePath)) {
      srcFileContent = fs.readFileSync(defaults.fallbackAuthDialogJavascriptFilePath, "utf8");
      fallbackAuthDialogPath = defaults.fallbackAuthDialogJavascriptFilePath;
    } else if (isTest) {
      srcFileContent = fs.readFileSync(defaults.testFallbackAuthDialogFilePath, "utf8");
    } else {
      if (isTest) {
        throw new ExpectedError(`${defaults.testFallbackAuthDialogFilePath} does not exist`);
      } else {
        const errorMessage: string = `${
          isTypecript ? defaults.fallbackAuthDialogTypescriptFilePath : defaults.fallbackAuthDialogJavascriptFilePath
        } does not exist`;
        throw new Error(errorMessage);
      }
    }

    // Check to see if the fallbackauthdialog file has already been updated and return if it has.
    if (!srcFileContent.includes("{PORT}")) {
      return false;
    }

    // Update fallbackauthdialog file
    const updatedSrcFileContent = srcFileContent
      .replace("{application GUID here}", applicationId)
      .replace("{PORT}", port);
    fs.writeFileSync(fallbackAuthDialogPath, updatedSrcFileContent);
    return true;
  } catch (err) {
    if (isTest) {
      throw new Error(`Unable to write SSO application data to ${defaults.testFallbackAuthDialogFilePath}. \n${err}`);
    } else {
      const errorMessage: string = `Unable to write SSO application data to ${
        isTypecript ? defaults.fallbackAuthDialogTypescriptFilePath : defaults.fallbackAuthDialogJavascriptFilePath
      }. \n${err}`;
      throw new Error(errorMessage);
    }
  }
}

async function updateProjectManifest(applicationId: string, port: string, manifestPath: string): Promise<boolean> {
  try {
    if (fs.existsSync(manifestPath)) {
      // Update manifest with application guid and unique manifest id
      const manifestContent: string = await fs.readFileSync(manifestPath, "utf8");

      // Check to see if the manifest has already been updated and return if it has
      if (!manifestContent.includes("{PORT}")) {
        return false;
      }

      // Update manifest file
      const re: RegExp = new RegExp("{application GUID here}", "g");
      const rePort = new RegExp("{PORT}", "g");
      const updatedManifestContent: string = manifestContent.replace(re, applicationId).replace(rePort, port);
      await fs.writeFileSync(manifestPath, updatedManifestContent);
      await modifyManifestFile(manifestPath, "random");
      return true;
    } else {
      const errorMessage: string = "Manifest does not exist at specified location";
      throw new ExpectedError(errorMessage);
    }
  } catch (err) {
    const errorMessage: string = `Unable to update manifest: \n${err}`;
    throw new Error(errorMessage);
  }
}

export async function writeApplicationData(
  applicationId: string,
  port: string,
  manifestPath: string = defaults.manifestFilePath,
  envFilePath: string = defaults.envDataFilePath,
  fallbackAuthDialogPath = defaults.fallbackAuthDialogTypescriptFilePath,
  isTest: boolean = false
): Promise<boolean> {
  const envFileUpdated: boolean = updateEnvFile(applicationId, port, envFilePath);
  const fallbackAuthDialogFileUpdated: boolean = updateFallBackAuthDialogFile(
    applicationId,
    port,
    fallbackAuthDialogPath,
    isTest
  );
  const manifestUpdated: boolean = await updateProjectManifest(applicationId, port, manifestPath);
  return envFileUpdated && fallbackAuthDialogFileUpdated && manifestUpdated;
}
