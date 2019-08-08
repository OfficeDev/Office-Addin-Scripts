import * as commander from "commander";
import { logErrorMessage } from "office-addin-cli";
import * as os from "os";
import * as path from "path";
import * as commands from "./command";

const telemetryJsonFilePath: string = path.join(os.homedir(), "/officeAddinTelemetry.json");

commander.name("office-addin-telemetry");
commander.version(process.env.npm_package_version || "(version not available)");

commander
    .command(`list`)
    .description(`Display the current telemetry settings.`)
    .action(commands.listTelemetrySettings);

commander
    .command(`off`)
    .description(`Sets the telemetry level to Basic.`)
    .action(commands.turnTelemetryOff);

commander
    .command(`on`)
    .description(`Sets the telemetry level to Verbose.`)
    .action(commands.turnTelemetryOn);

// if the command is not known, display an error
commander.on("command:*", function() {
    logErrorMessage(`The command syntax is not valid.\n`);
    process.exitCode = 1;
    commander.help();
});

if (process.argv.length > 2) {
    commander.parse(process.argv);
} else {
    commander.help();
}
