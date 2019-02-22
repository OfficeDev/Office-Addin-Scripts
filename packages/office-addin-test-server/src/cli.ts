import * as commander from "commander";
import * as commands from "./commands";

commander
    .command("start")
    .option("-p --port [port number]", "Port number must be between 0 - 65535. If no port specified, port defaults to 8080")
    .action(commands.start);

commander
    .command("stop")
    .action(commands.stop);

commander.parse(process.argv);