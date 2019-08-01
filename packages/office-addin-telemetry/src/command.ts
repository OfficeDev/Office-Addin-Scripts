import * as chalk from "chalk";
import * as commander from "commander";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { groupNameExists, readTelemetryJsonData, writeTelemetryJsonData } from "./telemetryJsonData";
const telemetryJsonFilePath: string = path.join(os.homedir(), "/officeAddinTelemetry.json");

export function stopTelemetryGroup(telemetryGroupName: string, command: commander.Command): void {
   const commandOption = getCommandOptionString(command.filepath);
   const telemetryConfigFilePath = commandOption ? commandOption : telemetryJsonFilePath;
   modifyTelemetryConfigSetting(telemetryGroupName, false /* disable */, telemetryConfigFilePath);
}

export function startTelemetryGroup(telemetryGroupName: string, command: commander.Command): void {
   const commandOption = getCommandOptionString(command.filepath);
   const telemetryConfigFilePath = commandOption ? commandOption : telemetryJsonFilePath;
   modifyTelemetryConfigSetting(telemetryGroupName, true /* enable */, telemetryConfigFilePath);
}

export function listTelemetryGroups(command: commander.Command): void {
   const commandOption = getCommandOptionString(command.filepath);
   const telemetryConfigFilePath = commandOption ? commandOption : telemetryJsonFilePath;
   if (fs.existsSync(telemetryConfigFilePath)) {
      const telemetryJsonData = readTelemetryJsonData(telemetryConfigFilePath);
      console.log(chalk.default.blue(`\nTelemetry groups and enabled settings listed in ${telemetryConfigFilePath}:\n`));
      for (const key of Object.keys(telemetryJsonData.telemetryInstances)) {
         console.log(`${key}:\n`);
         for (const value of Object.keys(telemetryJsonData.telemetryInstances[key])) {
         console.log(`  ${value}:${telemetryJsonData.telemetryInstances[key][value]}\n`);
         }
      }
   } else {
      console.log(chalk.default.red(`No telemetry configuration file found at ${telemetryConfigFilePath}`));
   }
}

function modifyTelemetryConfigSetting(telemetryGroupName: string, enable: boolean, telemetryConfigFilePath: string) {
   try {
      if (fs.existsSync(telemetryConfigFilePath)) {
         const telemetryJsonData = readTelemetryJsonData(telemetryJsonFilePath);
         if (groupNameExists(telemetryJsonData, telemetryGroupName)) {
            if (telemetryJsonData.telemetryInstances[telemetryGroupName].telemetryEnabled === enable) {
               console.log(chalk.default.yellow(`\nTelemetry is already set to ${enable} for telemetry group: ${chalk.default.blue(telemetryGroupName)}\n`));
            } else {
               telemetryJsonData.telemetryInstances[telemetryGroupName].telemetryEnabled = enable;
               writeTelemetryJsonData(telemetryJsonData, telemetryConfigFilePath);
               console.log(chalk.default.green(`\nTelemetry has been set to ${enable} for ${chalk.default.blue(telemetryGroupName)}\n`));
            }
         } else {
            console.log(chalk.default.yellow(`\nTelemetry group name ${chalk.default.blue(telemetryGroupName)} not found in ${telemetryConfigFilePath}\n`));
         }
      } else {
         console.log(chalk.default.red(`\nNo telemetry configuration file found at ${telemetryConfigFilePath}\n`));
      }
   } catch (err) {
      throw new Error(`Error occurred while trying to stop telemetry for ${telemetryGroupName}:\n${err}`);
   }
}

function getCommandOptionString(option: string | boolean, defaultValue?: string): string | undefined {
   // For a command option defined with an optional value, e.g. "--option [value]",
   // when the option is provided with a value, it will be of type "string", return the specified value;
   // when the option is provided without a value, it will be of type "boolean", return undefined.
   return (typeof (option) === "boolean") ? defaultValue : option;
}
