import * as chalk from "chalk";
import * as commander from "commander";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { groupName } from "./defaults";
import * as defaults from "./defaults";
import { telemetryLevel } from "./officeAddinTelemetry";
import * as jsonData from "./telemetryJsonData";
const telemetryJsonFilePath: string = path.join(os.homedir(), "/officeAddinTelemetry.json");

export function turnTelemetryGroupOff(): void {
   modifyTelemetryConfigSetting(telemetryLevel.basic /* disable */);
}

export function turnTelemetryGroupOn(): void {
   modifyTelemetryConfigSetting(telemetryLevel.verbose /* enable */);
}

export function listTelemetryGroups(command: commander.Command): void {
   if (fs.existsSync(defaults.telemetryJsonFilePath)) {
      const telemetryJsonData = jsonData.readTelemetryJsonData();
      console.log(chalk.default.blue(`\nTelemetry groups and enabled settings listed in ${defaults.telemetryJsonFilePath}:\n`));
      for (const key of Object.keys(telemetryJsonData.telemetryInstances)) {
         console.log(`${key}:\n`);
         for (const value of Object.keys(telemetryJsonData.telemetryInstances[key])) {
            console.log(`  ${value}:${telemetryJsonData.telemetryInstances[key][value]}\n`);
         }
      }
   } else {
      console.log(chalk.default.red(`No telemetry configuration file found at ${defaults.telemetryJsonFilePath}`));
   }
}

function modifyTelemetryConfigSetting(level: telemetryLevel) {
   try {
      if (fs.existsSync(telemetryJsonFilePath)) {
         if (jsonData.groupNameExists(groupName)) {
            if (jsonData.readTelemetryLevel(groupName) === level) {
               console.log(chalk.default.yellow(`\nTelemetry is already set to ${level} for telemetry group: ${chalk.default.blue(groupName)}\n`));
            } else {
               jsonData.modifyTelemetryJsonData(groupName, "telemetryLevel", level);
               console.log(chalk.default.green(`\nTelemetry has been set to ${level} for ${chalk.default.blue(groupName)}\n`));
            }
         } else {
            console.log(chalk.default.yellow(`\nTelemetry group name ${chalk.default.blue(groupName)} not found in ${telemetryJsonFilePath}\n`));
         }
      } else {
         console.log(chalk.default.red(`\nNo telemetry configuration file found at ${telemetryJsonFilePath}\n`));
      }
   } catch (err) {
      throw new Error(`Error occurred while trying to stop telemetry for ${groupName}:\n${err}`);
   }
}
