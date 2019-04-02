#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";

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
.option("--disable", "Disable debugging for the add-in")
.option("--debug-method <method>", "Specify the debug method: 'direct' or 'proxy'.")
.description("Configure debugging for the Office Add-in.")
.action(commands.debugging);

commander
.command("live-reload <manifest-path>")
.option("--enable", `Enable live-reload for the add-in.`)
.option("--disable", "Disable live-reload for the add-in")
.description("Configure live-reload for the Office Add-in.")
.action(commands.liveReload);

commander
.command("runtime-log")
.option("--enable [path]", `Enable the runtime log.`)
.option("--disable", "Disable the runtime log.")
.description("Configure the runtime log for all Office Add-ins.")
.action(commands.runtimeLogging);

commander
.command("source-bundle-url <manifest-path>")
.description("Specify values for components of the source bundle url.")
.option("-h,--host <host>", `The host name to use, or "" to use the default ('localhost').`)
.option("-p,--port <port>", `The port number to use, or "" to use the default (8081).`)
.option("--path <path>", `The path to use, or "" to use the default.`)
.option("-e,--extension <extension>", `The extension to use, or "" to use the default (".bundle").`)
.action(commands.sourceBundleUrl);

commander.parse(process.argv);
