import * as commander from "commander";
import * as commands from "./commands";
import { defaultPort } from "./testServer";

commander.name("office-addin-test-server");
commander.version(process.env.npm_package_version || "(version not available)");

commander
    .command("start")
    .option(`-p --port [port number]", "Port number must be between 0 - 65535. Default: ${defaultPort}`)
    .action(commands.start);

if (process.argv.length > 2) {
    commander.parse(process.argv);
} else {
    commander.help();
}
