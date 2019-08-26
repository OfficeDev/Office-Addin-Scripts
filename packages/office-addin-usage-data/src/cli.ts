import * as commander from "commander";
import { logErrorMessage } from "office-addin-cli";
import * as commands from "./command";

commander.name("office-addin-usage-data");
commander.version(process.env.npm_package_version || "(version not available)");

commander
    .command(`list`)
    .description(`Display the current usage data settings.`)
    .action(commands.listUsageDataSettings);

commander
    .command(`off`)
    .description(`Turn off reporting of anonymous usage data.`)
    .action(commands.turnUsageDataOff);

commander
    .command(`on`)
    .description(`Turn on reporting of anonymous usage data.`)
    .action(commands.turnUsageDataOn);

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
