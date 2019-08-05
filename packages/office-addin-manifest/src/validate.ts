// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { createReadStream } from "fs";
import fetch from "node-fetch";
import { readManifestFile } from "./manifestInfo";

export class ManifestValidationDetails {
    public capabilities?: string[];
    public capabilitiesCodes?: string[];
    public defaultLocale?: string;
    public defaultSourceLocations?: string[];
    public description?: string;
    public displayName?: string;
    public iconUrl?: string;
    public localizedDescriptions?: object;
    public localizedIconUrls?: object;
    public productId?: string;
    public providerName?: string;
    public supportedProducts?: ManifestValidationProduct[];
    public version?: string;
}

export class ManifestValidationIssue {
    public code?: string;
    public column?: number;
    public line?: number;
    public title?: string;
    public detail?: string;
    public link?: string;
}

export class ManifestValidationProduct {
    public productCode?: string;
    public title?: string;
    public version?: string;
}

export class ManifestValidationReport {
    public result?: string;
    public errors?: ManifestValidationIssue[];
    public warnings?: ManifestValidationIssue[];
    public suggestions?: ManifestValidationIssue[];
    public infos?: ManifestValidationIssue[];
}

export class ManifestValidation {
    public isValid: boolean;
    public report?: ManifestValidationReport;
    public details?: ManifestValidationDetails;
    public status?: number;

    constructor() {
        this.isValid = false;
    }
}

export async function validateManifest(manifestPath: string): Promise<ManifestValidation> {
    const validation: ManifestValidation = new ManifestValidation();

    // read the manifest file
    // const manifest = await readManifestFile(manifestPath);
    const stream = await createReadStream(manifestPath);
    let response;

    try {
        response = await fetch("https://verificationservice.osi.office.net/ova/addincheckingagent.svc/api/addincheck",
            {
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
    const report: ManifestValidationReport = (json && json.checkReport) ? json.checkReport.validationReport : undefined;
    const details: ManifestValidationDetails = (json && json.checkReport) ? json.checkReport.details : undefined;

    if (report) {
        validation.report = report;
        validation.details = details;

        if (report.result) {
            switch (report.result.toLowerCase()) {
                case "passed":
                    validation.isValid = true;
                    break;
            }
        }
    } else {
        throw new Error("The manifest validation service did not return the expected response.");
    }

    return validation;
}
