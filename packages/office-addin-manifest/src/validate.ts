// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { createReadStream } from "fs";
import { ManifestUtil, TeamsAppManifest } from "@microsoft/teams-manifest";
import fetch from "node-fetch";
import { OfficeAddinManifest } from "./manifestOperations";
import { usageDataObject } from "./defaults";

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
  public statusText?: string;

  constructor() {
    this.isValid = false;
  }
}

export async function validateManifest(
  manifestPath: string,
  verifyProduction: boolean = false
): Promise<ManifestValidation> {
  try {
    const validation: ManifestValidation = new ManifestValidation();

    // read the manifest file to ensure the file path is valid
    await OfficeAddinManifest.readManifestFile(manifestPath);

    if (manifestPath.endsWith(".json")) {
      const manifest: TeamsAppManifest = await ManifestUtil.loadFromPath(manifestPath);
      const validationResult: string[] = await ManifestUtil.validateManifest(manifest);
      if (validationResult.length !== 0) {
        // There are errors
        validation.isValid = false;
        validation.report = new ManifestValidationReport();
        validation.report.errors = [];
        validationResult.forEach((error: string) => {
          let issue: ManifestValidationIssue = new ManifestValidationIssue();
          issue.content = error;
          issue.title = "Error";

          validation.report?.errors?.push(issue);
        });
      } else {
        validation.isValid = true;
      }
    } else {
      const stream = await createReadStream(manifestPath);
      const clientId: string = verifyProduction ? "Default" : "devx";
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

      validation.status = response.status;
      validation.statusText = response.statusText;

      const text = await response.text();

      try {
        const json = JSON.parse(text.trim());
        if (json) {
          validation.report = json;
        }
      } catch {} // eslint-disable-line no-empty

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
  } catch (err: any) {
    usageDataObject.reportException("validateManifest()", err);
    throw err;
  }
}
