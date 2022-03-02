// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { createReadStream } from "fs";
import { ManifestUtil, TeamsAppManifest } from "@microsoft/teams-manifest";
import fetch from "node-fetch";
import { OfficeAddinManifest } from "./manifestOperations";

export class ManifestValidationDetails {
  public adminInstallOnly?: boolean;
  public capabilities?: object;
  public defaultLocale?: string;
  public description?: string;
  public displayName?: string;
  public hosts?: string[];
  public iconUrl?: string;
  public localizedDescriptions?: object;
  public localizedIconUrls?: object;
  public localizedRootSourceUrls?: object;
  public productId?: string;
  public providerName?: string;
  public requirements?: string;
  public rootSourceUrl?: string;
  public subtype?: string;
  public supportedLanguages?: string[];
  public supportedProducts?: ManifestValidationProduct[];
  public type?: string;
  public version?: string;
}

export class ManifestValidationIssue {
  public code?: string;
  public column?: number;
  public line?: number;
  public title?: string;
  public content?: string;
  public helpUrl?: string;
}

export class ManifestValidationProduct {
  public code?: string;
  public title?: string;
  public version?: string;
}

export class ManifestValidationReport {
  public status?: string;
  public errors?: ManifestValidationIssue[];
  public warnings?: ManifestValidationIssue[];
  public notes?: ManifestValidationIssue[];
  public addInDetails?: ManifestValidationDetails;
}

export class ManifestValidation {
  public isValid: boolean;
  public report?: ManifestValidationReport;
  public status?: number;
  public jsonErrors?: string[];

  constructor() {
    this.isValid = false;
  }
}

export async function validateManifest(
  manifestPath: string,
  verifyProduction: boolean = false
): Promise<ManifestValidation> {
  const validation: ManifestValidation = new ManifestValidation();
  const clientId: string = verifyProduction ? "Default" : "devx";

  // read the manifest file to ensure the file path is valid
  await OfficeAddinManifest.readManifestFile(manifestPath);

  if (manifestPath.endsWith(".json")) {
    const manifest: TeamsAppManifest = await ManifestUtil.loadFromPath(manifestPath);
    const validationResult: string[] = await ManifestUtil.validateManifest(manifest);

    if (validationResult.length !== 0) {
      // There are errors
      validation.isValid = false;
      validation.jsonErrors = validationResult;
    } else {
      validation.isValid = true;
    }
  } else {
    const stream = await createReadStream(manifestPath);
    let response;

    try {
      response = await fetch(`https://validationgateway.omex.office.net/package/api/check?clientId=${clientId}`, {
        body: stream,
        headers: {
          "Content-Type": "application/xml",
        },
        method: "POST",
      });
    } catch (err) {
      throw new Error(`Unable to contact the manifest validation service.\n${err}`);
    }

    const text = await response.text();
    const json = JSON.parse(text.trim());

    if (json) {
      validation.report = json;
      validation.status = response.status;
    }

    if (validation.report) {
      const result = validation.report.status;

      if (result) {
        switch (result.toLowerCase()) {
          case "accepted":
            validation.isValid = true;
            break;
        }
      }
    }
  }
  return validation;
}
