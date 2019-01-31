import * as childProcess from "child_process";
import * as crypto from "crypto";
import * as net from "net";

/**
 * Determines whether a port is in use.
 * @param port port number (0 - 65535)
 * @returns true if port is in use; false otherwise.
 */
export function isPortInUse(port: number): Promise<boolean> {
  validatePort(port);

  return new Promise((resolve) => {
    const server = net
      .createServer()
      .once("error", () => {
        resolve(true);
      })
      .once("listening", () => {
        server.close();
        resolve(false);
      })
      .listen(port);
  });
}

/**
 * Parse the port from a string which ends with colon and a number.
 * @param text string to parse
 * @example "127.0.0.1:3000" returns 3000
 * @example "[::1]:1900" returns 1900
 * @example "Local Address" returns undefined
 */
function parsePort(text: string): number | undefined {
    const result = text.match(/:(\d+)$/);

    return result ? parseInt(result[1], 10) : undefined;
  }

/**
 * Return the process ids using the port.
 * @param port port number (0 - 65535)
 * @returns Promise to array containing process ids, or empty if none.
 */
export function getProcessIdsForPort(port: number): Promise<number[]> {
  validatePort(port);

  return new Promise((resolve, reject) => {
    const isWin32 = process.platform === "win32";
    const command = isWin32 ? `netstat -ano` : `lsof -n -i:${port}`;

    childProcess.exec(command, (error, stdout) => {
      if (error) {
        if (error.code === 1) {
          // no processes are using the port
          resolve([]);
        } else {
          reject(error);
        }
      } else {
        const processIds = new Set();
        const lines = stdout.trim().split("\n");
        if (isWin32) {
          lines.forEach((line) => {
            const [protocol, localAddress, foreignAddress, status, processId] = line.split(" ").filter((text) => text);
            if (processId !== undefined) {
              const localAddressPort = parsePort(localAddress);
              if (localAddressPort === port) {
                processIds.add(parseInt(processId, 10));
              }
            }
          });
        } else {
          lines.forEach((line) => {
            const [process, processId, user, fd, type, device, size, node, name] = line.split(" ").filter((text) => text);
            if ((processId !== undefined) && (processId !== "PID")) {
              processIds.add(parseInt(processId, 10));
            }
          });
        }

        resolve(Array.from(processIds));
      }
    });
  });
}

/**
 * Returns a random port number which is not in use.
 * @returns Promise to number from 0 to 65535
 */
export async function randomPortNotInUse(): Promise<number> {
  let port: number;
  do {
    port = randomPortNumber();
  } while (await isPortInUse(port));

  return port;
}

/**
 * Returns a random number between 0 and 65535
 */
function randomPortNumber(): number {
   return crypto.randomBytes(2).readUInt16LE(0, false);
}

/**
 * Throw an error if the port is not a valid number.
 * @param port port number
 * @throws Error if port is not a number from 0 to 65535.
 */
function validatePort(port: number): void {
  if ((typeof(port) !== "number") || (port < 0) || (port > 65535)) {
    throw new Error("Port should be a number from 0 to 65535.");
  }
}
