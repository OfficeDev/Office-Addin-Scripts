#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-cli";
import * as commands from "./commands";

commander.name("office-addin-dev-settings");
commander.version(process.env.npm_package_version || "(version not available)");

commander
    .command("appcontainer <manifest-path>")
    .description("Display or configure the appcontainer used to run the Office Add-in.")
    .option("--loopback", `Allow access to loopback addresses such as "localhost".`)
    .option("--prevent-loopback", `Prevent access to loopback addresses such as "localhost".`)
    .option("-y,--yes", "Provide approval without any prompts.")
    .action(commands.appcontainer);

commander
    .command("clear [manifest-path]")
    .description("Clear all dev settings for the Office Add-in.")
    .action(commands.clear);

commander
    .command("debugging <manifest-path>")
    .option("--enable", `Enable debugging for the add-in.`)
    .option("--disable", "Disable debugging for the add-in.")
    .option("--debug-method <method>", "Specify the debug method: 'direct' or 'proxy'.")
    .option("--open-dev-tools", "Open the web browser dev tools (if supported).")
    .description("Configure debugging for the Office Add-in.")
    .action(commands.debugging);

commander
    .command("live-reload <manifest-path>")
    .option("--enable", `Enable live-reload for the add-in.`)
    .option("--disable", "Disable live-reload for the add-in")
    .description("Configure live-reload for the Office Add-in.")
    .action(commands.liveReload);

commander
    .command("register <manifest-path>")
    .description("Register the Office Add-in for development.")
    .action(commands.register);

commander
    .command("registered")
    .description("Show the Office Add-ins registered for development.")
    .action(commands.registered);

commander
    .command("runtime-log")
    .option("--enable [path]", `Enable the runtime log.`)
    .option("--disable", "Disable the runtime log.")
    .description("Configure the runtime log for all Office Add-ins.")
    .action(commands.runtimeLogging);

commander
    .command("sideload <manifest-path> [app-type]")
    .description("Launch Office with the Office Add-in loaded.")
    .option("-a,--app <app>", `The Office app to launch. ("Excel", "Outlook", "PowerPoint", or "Word")`)
    .option("-d,--document <document>", `The location of the document to be sideloaded - this can be an absolute file path or url`)
    .action(commands.sideload)
    .on("--help", () => {
        console.log("\n[app-type] specifies the type of Office app::\n");
        console.log("\t'desktop': Office app for Windows or Mac (default),");
        console.log("\t'web': Office running in the web browser");
    });

commander
    .command("source-bundle-url <manifest-path>")
    .description("Specify values for components of the source bundle url.")
    .option("-h,--host <host>", `The host name to use, or "" to use the default ('localhost').`)
    .option("-p,--port <port>", `The port number to use, or "" to use the default (8081).`)
    .option("--path <path>", `The path to use, or "" to use the default.`)
    .option("-e,--extension <extension>", `The extension to use, or "" to use the default (".bundle").`)
    .action(commands.sourceBundleUrl);

commander
    .command("unregister <manifest-path>")
    .description("Unregister the Office Add-in for development.")
    .action(commands.unregister);

commander
    .command("webview <manifest-path> [web-view-type]")
    .description("Specify the type of web view to use when debugging. Windows only.")
    .action(commands.webView)
    .on("--help", () => {
        console.log("\nFor [web-view-type], choose one of the following values:\n");
        console.log("\t'edge' or 'edge-chromium' for Microsoft Edge (Chromium)");
        console.log("\t'edge-legacy' for the legacy Microsoft Edge (EdgeHTML)");
        console.log("\t'ie' for Internet Explorer 11");
        console.log("\t'default' to remove any preference");
        console.log("\nOmit [web-view-type] to see the current setting.");
    });

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
