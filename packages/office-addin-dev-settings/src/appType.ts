import { ExpectedError } from "office-addin-usage-data";

/**
 * The type of Office application.
 */
export enum AppType {
  /* eslint-disable no-unused-vars */
  /**
   * Office application for Windows or Mac
   */
  Desktop = "desktop",

  /**
   * Office application for the web browser
   */
  Web = "web",
}

/**
 * Parse the input text and get the associated AppType
 * @param text app-type/platform text
 * @returns AppType or undefined.
 */
export function parseAppType(text: string | undefined): AppType | undefined {
  switch (text ? text.toLowerCase() : undefined) {
    case "desktop":
    case "macos":
    case "win32":
    case "ios":
    case "android":
      return AppType.Desktop;
    case "web":
      return AppType.Web;
    case undefined:
      return undefined;
    default:
      throw new ExpectedError(
        `Please select a valid app type instead of '${text}'.`
      );
  }
}
