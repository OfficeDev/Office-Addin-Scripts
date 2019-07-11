import * as commander from "commander";

commander.name("office-addin-telemetry");
commander.version(process.env.npm_package_version || "(version not available)");

commander
    .version('0.0.1')
    .command("start")
    .command("stop")
    .command('telemetryStatus')
    //.option(`-p --port [port number]", "Port number must be between 0 - 65535. Default: ${defaultPort}`)
    //.action(commands.start);

    commander.on("command:*", function() {
        console.error(`The command syntax is not valid.\n`);
        process.exitCode = 1;
        commander.help();
    });

    if (process.argv.length > 2) {
        commander.parse(process.argv);
    } else {
        commander.help();
    }