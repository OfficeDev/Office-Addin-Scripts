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
  Project = "project",
  PowerPoint = "powerpoint",
  Word = "word"
  // when adding new entries, update the toOfficeApp() function
  // since there isn't an automatic reverse mapping from string to enum values
}

// initialized once since this list won't change
const officeApps: OfficeApp[] = Object.keys(OfficeApp).map<OfficeApp>(key => parseOfficeApp(key));

/**
 * Returns the OfficeApp specified by the value.
 * @param value string
 * @throws Error if the value is not a valid Office app.
 */
export function parseOfficeApp(value: string): OfficeApp {
  const officeApp = toOfficeApp(value);

  if (!officeApp) {
    throw new Error(`${value} is not a valid Office app.`);
  }

  return officeApp;
}

/**
 * Returns the Office apps specified by the value.
 * @param input "all" for all Office apps, or a comma-separated list of one or more Office apps.
 * @throws Error if a value is not a valid Office app.
 */
export function parseOfficeApps(input: string): OfficeApp[] {
  if (input === "all") {
    return getOfficeApps();
  } else {
    return input.split(",").map<OfficeApp>(appString => parseOfficeApp(appString));
  }
}

/**
 * Returns the OfficeApp for the value, or undefined if not valid.
 * @param value OfficeApp string
 */
export function toOfficeApp(value: string): OfficeApp | undefined {
  switch (value.toLowerCase()) {
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

/**
 * Returns the supported Office apps
 */
export function getOfficeApps(): OfficeApp[] {
    return officeApps;
}

