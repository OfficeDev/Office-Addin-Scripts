import * as childProcess from "child_process";

/* global console */

export async function publish(zipPath: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const command = `npx @microsoft/teamsfx-cli@0.11.0-rc.0 provision manifest --file-path ${zipPath}`;

    childProcess.exec(command, (error, stdout) => {
      if (error) {
        reject(stdout);
      } else {
        console.log("Successfully registered package with https://dev.teams.microsoft.com/apps");
        resolve();
      }
    });
  });
}
