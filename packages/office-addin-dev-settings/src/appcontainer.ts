import * as childProcess from "child_process";
import { URL } from "whatwg-url";

/**
 * Adds a loopback exemption for the appcontainer.
 * @param name Appcontainer name
 * @throws Error if platform does not support appcontainers.
 */
export function addLoopbackExemptionForAppcontainer(name: string): Promise<void> {
  if (!isAppcontainerSupported()) {
      throw new Error(`Platform not supported: ${process.platform}.`);
  }

  return new Promise((resolve, reject) => {
      const command = `CheckNetIsolation.exe LoopbackExempt -a -n=${name}`;

      childProcess.exec(command, (error, stdout) => {
      if (error) {
          reject(error);
      } else {
          resolve();
      }
      });
  });
}

/**
 * Returns whether appcontainer is supported on the current platform.
 * @returns True if platform supports using appcontainer; false otherwise.
 */
export function isAppcontainerSupported() {
  switch (process.platform) {
    case "win32":
      return true;
    default:
      return false;
  }
}

/**
 * Adds a loopback exemption for the appcontainer.
 * @param name Appcontainer name
 * @throws Error if platform does not support appcontainers.
 */
export function isLoopbackExemptionForAppcontainer(name: string): Promise<boolean> {
  if (!isAppcontainerSupported()) {
      throw new Error(`Platform not supported: ${process.platform}.`);
  }

  return new Promise((resolve, reject) => {
      const command = `CheckNetIsolation.exe LoopbackExempt -s`;

      childProcess.exec(command, (error, stdout) => {
      if (error) {
          reject(error);
      } else {
        const expr = new RegExp(`Name: ${name}`, "i");
        const found: boolean = expr.test(stdout);
        resolve(found);
      }
    });
  });
}

/**
 * Returns the name of the appcontainer used to run an Office Add-in.
 * @param sourceLocation Source location of the Office Add-in.
 * @param isFromStore True if installed from the Store; false otherwise.
 */
export function getAppcontainerName(sourceLocation: string, isFromStore = false): string {
  const url: URL = new URL(sourceLocation);
  const origin: string = url.origin;
  const addinType: number = isFromStore ? 0 : 1; // 0 if from Office Add-in store, 1 otherwise.
  const guid: string = "04ACA5EC-D79A-43EA-AB47-E50E47DD96FC";

  // Appcontainer name format is "{addinType}_{origin}{guid}".
  // Replace characters ":" and "/" with "_".
  const name = `${addinType}_${origin}${guid}`.replace(/[://]/g, "_");

  return name;
}

/**
 * Removes a loopback exemption for the appcontainer.
 * @param name Appcontainer name
 * @throws Error if platform doesn't support appcontainers.
 */
export function removeLoopbackExemptionForAppcontainer(name: string): Promise<void> {
  if (!isAppcontainerSupported()) {
      throw new Error(`Platform not supported: ${process.platform}.`);
  }

  return new Promise((resolve, reject) => {
      const command = `CheckNetIsolation.exe LoopbackExempt -d -n=${name}`;

      childProcess.exec(command, (error, stdout) => {
      if (error) {
          reject(error);
      } else {
          resolve();
      }
      });
  });
}
