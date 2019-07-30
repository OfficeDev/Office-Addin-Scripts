// copyright (c) Microsoft Corporation. All rights reserved.
// licensed under the MIT license.

/**
 * The Office apps which can host Office Add-ins.
 */
export enum OfficeApp {
  // the string values should be lowercase
  Excel = "excel",
  OneNote = "onenote",
  Outlook = "outlook",
  PowerPoint = "powerpoint",
  Project = "project",
  Word = "word",
  // when adding new entries, update the toOfficeApp() function
  // since there isn't an automatic reverse mapping from string to enum values
}

// initialized once since this list won't change
const officeApps: OfficeApp[] = Object.keys(OfficeApp).map<OfficeApp>(key => parseOfficeApp(key));

/**
 * Get the Office app for the manifest Host name
 * @param host Host name
 */
export function getOfficeAppForManifestHost(host: string): OfficeApp | undefined {
  switch (host.toLowerCase()) {
    case "document":
      return OfficeApp.Word;
    case "mailbox":
      return OfficeApp.Outlook;
    case "notebook":
      return OfficeApp.OneNote;
    case "presentation":
      return OfficeApp.PowerPoint;
    case "project":
      return OfficeApp.Project;
    case "workbook":
      return OfficeApp.Excel;
    default:
      return undefined;
  }
}

/**
 * Gets the Office application name suitable for display.
 * @param app Office app
 */
export function getOfficeAppName(app: OfficeApp): string {
  switch (app) {
    case OfficeApp.Excel:
      return "Excel";
    case OfficeApp.OneNote:
      return "OneNote";
    case OfficeApp.Outlook:
      return "Outlook";
    case OfficeApp.PowerPoint:
      return "PowerPoint";
    case OfficeApp.Project:
      return "Project";
    case OfficeApp.Word:
      return "Word";
    default:
      throw new Error(`Unable to provide name for Office app "${app}"`);
  }
}

/**
 * Gets the Office application names suitable for display.
 * @param apps Office apps
 */
export function getOfficeAppNames(apps: OfficeApp[]): string[] {
  return apps.map(app => getOfficeAppName(app));
}

/**
 * Returns the Office apps that support Office add-ins.
 */
export function getOfficeApps(): OfficeApp[] {
  return officeApps;
}

/**
 * Get the Office apps for the manifest Host names.
 * @param hosts Host names specified in the manifest.
 */
export function getOfficeAppsForManifestHosts(hosts?: string[]): OfficeApp[] {
  const apps: OfficeApp[] = [];

  if (hosts) {
    hosts.forEach(host => {
      const app = getOfficeAppForManifestHost(host);

      if (app) {
        apps.push(app);
      }
    });
  }

  return apps;
}

/**
 * Converts the string to the OfficeApp enum value.
 * @param value string
 * @throws Error if the value is not a valid Office app.
 */
export function parseOfficeApp(value: string): OfficeApp {
  const officeApp = toOfficeApp(value);

  if (!officeApp) {
    throw new Error(`"${value}" is not a valid Office app.`);
  }

  return officeApp;
}

/**
 * Converts the strings to the OfficeApp enum values.
 * @param input "all" for all Office apps, or a comma-separated list of one or more Office apps.
 * @throws Error if a value is not a valid Office app.
 */
export function parseOfficeApps(input: string): OfficeApp[] {
  if (input.trim().toLowerCase() === "all") {
    return getOfficeApps();
  } else {
    return input.split(",").map<OfficeApp>(appString => parseOfficeApp(appString));
  }
}

/**
 * Returns the OfficeApp enum for the value, or undefined if not valid.
 * @param value Office app string
 */
export function toOfficeApp(value: string): OfficeApp | undefined {
  switch (value ? value.trim().toLowerCase() : value) {
    case OfficeApp.Excel:
      return OfficeApp.Excel;
    case OfficeApp.OneNote:
      return OfficeApp.OneNote;
    case OfficeApp.Outlook:
      return OfficeApp.Outlook;
    case OfficeApp.PowerPoint:
      return OfficeApp.PowerPoint;
    case OfficeApp.Project:
      return OfficeApp.Project;
    case OfficeApp.Word:
      return OfficeApp.Word;
    default:
      return undefined;
  }
}
