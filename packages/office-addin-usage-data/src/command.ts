import * as chalk from "chalk";
import * as defaults from "./defaults";
import { UsageDataLevel } from "./officeAddinUsage-Data";
import * as jsonData from "./usageDataJsonData";

export function listUsageDataSettings(): void {
   const usageDataSettings = jsonData.readUsageDataSettings(defaults.groupName);
   if (usageDataSettings) {
      console.log(chalk.default.blue(`\nUsage-Data settings for ${defaults.groupName}:\n`));
      for (const value of Object.keys(usageDataSettings)) {
         console.log(`  ${value}: ${usageDataSettings[value]}\n`);
      }
   } else {
      console.log(chalk.default.red(`No usage data settings for ${defaults.groupName}`));
   }
}

export function turnUsageDataOff(): void {
   setUsageDataLevel(UsageDataLevel.off);
}

export function turnUsageDataOn(): void {
   setUsageDataLevel(UsageDataLevel.on);
}
function setUsageDataLevel(UsageData: UsageDataLevel) {
   try {
      jsonData.modifyUsageDataJsonData(defaults.groupName, "usageDataLevel", UsageData);
      console.log(chalk.default.green(`Usage-Data is now ${UsageData}`));
   } catch (err) {
      throw new Error(`Error occurred while trying to change UsageData level`);
   }
}
