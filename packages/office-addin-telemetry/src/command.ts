import * as commander from "commander";
export async function help(command: commander.Command) {
   console.log("This package allows for sending telemetry event and exception data to the selected telemetry infrastructure (e.g. ApplicationInsights).");
}
