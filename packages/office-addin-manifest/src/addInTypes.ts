// copyright (c) Microsoft Corporation. All rights reserved.
// licensed under the MIT license.

import { ExpectedError } from "office-addin-usage-data";

/**
 * The types of Office add-ins.
 */
export enum AddInType {
  // the string values should be lowercase
  Content = "content",
  Mail = "mail",
  TaskPane = "taskpane",
  // when adding new entries, update the other functions
}

// initialized once since this list won't change
const addInTypes: AddInType[] = Object.keys(AddInType).map<AddInType>((key) => parseAddInType(key));

/**
 * Get the Office app for the manifest Host name
 * @param host Host name
 */
export function getAddInTypeForManifestOfficeAppType(officeAppType: string): AddInType | undefined {
  switch (officeAppType ? officeAppType.trim().toLowerCase() : officeAppType) {
    case "contentapp":
      return AddInType.Content;
    case "mailapp":
      return AddInType.Mail;
    case "taskpaneapp":
      return AddInType.TaskPane;
    default:
      return undefined;
  }
}

/**
 * Returns the Office add-in types.
 */
export function getAddInTypes(): AddInType[] {
  return addInTypes;
}

/**
 * Converts the string to the AddInType enum value.
 * @param value string
 * @throws Error if the value is not a valid Office add-in type.
 */
export function parseAddInType(value: string): AddInType {
  const addInType = toAddInType(value);

  if (!addInType) {
    throw new ExpectedError(`${value} is not a valid Office add-in type.`);
  }

  return addInType;
}

/**
 * Converts the strings to the AddInType enum values.
 * @param input "all" for all Office add-in types, or a comma-separated list of one or more Office apps.
 * @throws Error if a value is not a valid Office app.
 */
export function parseAddInTypes(input: string): AddInType[] {
  if (input.trim().toLowerCase() === "all") {
    return getAddInTypes();
  } else {
    return input.split(",").map<AddInType>((appString) => parseAddInType(appString));
  }
}

/**
 * Returns the AddInType enum for the value, or undefined if not valid.
 * @param value Office add-in type string
 */
export function toAddInType(value: string): AddInType | undefined {
  switch (value.trim().toLowerCase()) {
    case AddInType.Content:
      return AddInType.Content;
    case AddInType.Mail:
      return AddInType.Mail;
    case AddInType.TaskPane:
      return AddInType.TaskPane;
    default:
      return undefined;
  }
}
