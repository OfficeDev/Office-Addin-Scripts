import * as commander from "commander";
import { logErrorMessage } from "office-addin-cli";
import * as os from "os";
import * as path from "path";
import * as commands from "./command";

const telemetryJsonFilePath: string = path.join(os.homedir(), "/officeAddinTelemetry.json");

commander.name("office-addin-telemetry");
commander.version(process.env.npm_package_version || "(version not available)");

commander
    .command(`on`)
    .description(`Sets telemetry level to verbose for the specified group name`)
    .action(commands.turnTelemetryGroupOn);

commander
    .command(`off`)
    .description(`Sets telemetry level to basic for the specified group name`)
    .action(commands.turnTelemetryGroupOff);

commander
    .command(`list`)
    .description(`List all telemetry groups in the specified telemetry config file.`)
    .action(commands.listTelemetryGroups);

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
