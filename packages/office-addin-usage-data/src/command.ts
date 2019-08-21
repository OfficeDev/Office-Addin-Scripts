import * as defaults from "./defaults";
import { UsageDataLevel } from "./usageData";
import * as jsonData from "./usageDataSettings";

export function listUsageDataSettings(): void {
   const usageDataSettings = jsonData.readUsageDataSettings(defaults.groupName);

   if (usageDataSettings) {
      for (const value of Object.keys(usageDataSettings)) {
         console.log(`${value}: ${usageDataSettings[value]}\n`);
      }
   } else {
      console.log(`No usage data settings.`);
   }
}

export function turnUsageDataOff(): void {
   setUsageDataLevel(UsageDataLevel.off);
}

export function turnUsageDataOn(): void {
   setUsageDataLevel(UsageDataLevel.on);
}

function setUsageDataLevel(level: UsageDataLevel) {
   try {
      jsonData.modifyUsageDataJsonData(defaults.groupName, "usageDataLevel", level);

      switch (level) {
         case UsageDataLevel.off:
            console.log("Usage data has been turned off.");
            break;
         case UsageDataLevel.on:
            console.log("Usage data has been turned on.");
            break;
      }
   } catch (err) {
      throw new Error(`Unable to set the usage data level.\n${err}`);
   }
}
