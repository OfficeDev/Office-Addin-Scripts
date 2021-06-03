// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { createReadStream } from "fs";
import fetch from "node-fetch";
import { readManifestFile } from "./manifestInfo";
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

  constructor() {
    this.isValid = false;
  }
}

export async function validateManifest(
  manifestPath: string
): Promise<ManifestValidation> {
  try {
    const validation: ManifestValidation = new ManifestValidation();

    // read the manifest file to ensure the file path is valid
    await readManifestFile(manifestPath);

    const stream = await createReadStream(manifestPath);
    let response;

    try {
      response = await fetch(
        "https://validationgateway.omex.office.net/package/api/check?gates=DisableIconDimensionValidation",
        {
          body: stream,
          headers: {
            "Content-Type": "application/xml",
          },
          method: "POST",
        }
      );
    } catch (err) {
      throw new Error(
        `Unable to contact the manifest validation service.\n${err}`
      );
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
    usageDataObject.reportSuccess("validateManifest()");

    return validation;
  } catch (err) {
    usageDataObject.reportException("validateManifest()", err);
    throw err;
  }
}
