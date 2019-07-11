// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
import * as commander from "commander";
import { OfficeAddinTelemetry } from "./officeAddinTelemetry";
var addInTelemetry: any;
export async function help(command: commander.Command) {
   console.log("This is a telemetry package able to intergrate easily into desired programs creating telemetry for a desired data infrastructure");
}