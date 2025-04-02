// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

export class DefaultSettings {
  public sourceLocation?: string;
  public requestedWidth?: string;
  public requestedHeight?: string;
}

export enum ManifestType {
  JSON = "json",
  XML = "xml",
}

export class ManifestInfo {
  public id?: string;
  public allowSnapshot?: string;
  public alternateId?: string;
  public appDomains?: string[];
  public defaultLocale?: string;
  public description?: string;
  public displayName?: string;
  public highResolutionIconUrl?: string;
  public hosts?: string[];
  public iconUrl?: string;
  public officeAppType?: string;
  public permissions?: string;
  public providerName?: string;
  public supportUrl?: string;
  public version?: string;
  public manifestType?: ManifestType;

  public defaultSettings?: DefaultSettings;
}
