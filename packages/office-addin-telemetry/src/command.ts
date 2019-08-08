import * as chalk from "chalk";
import * as commander from "commander";
import * as fs from "fs";
import * as defaults from "./defaults";
import { telemetryLevel } from "./officeAddinTelemetry";
import * as jsonData from "./telemetryJsonData";

export function listTelemetrySettings(): void {
   const telemetrySettings = jsonData.readTelemetrySettings(defaults.groupName);
   if (telemetrySettings) {
      console.log(chalk.default.blue(`\nTelemetry settings for ${defaults.groupName}:\n`));
      for (const value of Object.keys(telemetrySettings)) {
         console.log(`  ${value}: ${telemetrySettings[value]}\n`);
      }
   } else {
      console.log(chalk.default.red(`No telemetry settings for ${defaults.groupName}`));
   }
}

export function turnTelemetryOff(): void {
   setTelemetryLevel(telemetryLevel.basic);
}

export function turnTelemetryOn(): void {
   setTelemetryLevel(telemetryLevel.verbose);
}

function setTelemetryLevel(level: telemetryLevel) {
   try {
      jsonData.modifyTelemetryJsonData(defaults.groupName, "telemetryLevel", level);
      console.log(chalk.default.green(`Telemetry level is now set to ${level}`));
   } catch (err) {
      throw new Error(`Error occurred while trying to change telemetry level`);
   }
}
