// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/* eslint-disable @typescript-eslint/no-unused-vars */

import { UsageDataLevel } from "./usageData";

/**
 * @deprecated Usage data reporting has been removed - always returns false
 */
export function needToPromptForUsageData(_groupName: string): boolean {
  return false;
}

/**
 * @deprecated Usage data reporting has been removed - no-op
 */
export function modifyUsageDataJsonData(_groupName: string, _property: any, _value: any): void {}

/**
 * @deprecated Usage data reporting has been removed - returns empty string
 */
export function readDeviceID(): string {
  return "";
}

/**
 * @deprecated Usage data reporting has been removed - returns undefined
 */
export function readUsageDataJsonData(): any {
  return undefined;
}

/**
 * @deprecated Usage data reporting has been removed - always returns off
 */
export function readUsageDataLevel(_groupName: string): UsageDataLevel {
  return UsageDataLevel.off;
}

/**
 * @deprecated Usage data reporting has been removed - returns undefined
 */
export function readUsageDataObjectProperty(_groupName: string, _propertyName: string): any {
  return undefined;
}

/**
 * @deprecated Usage data reporting has been removed - no-op
 */
export function writeUsageDataJsonData(_groupName: string, _level: UsageDataLevel): void {}

/**
 * @deprecated Usage data reporting has been removed - always returns false
 */
export function groupNameExists(_groupName: string): boolean {
  return false;
}

/**
 * @deprecated Usage data reporting has been removed - returns undefined
 */
export function readUsageDataSettings(_groupName?: string): object | undefined {
  return undefined;
}
