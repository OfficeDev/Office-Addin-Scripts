// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
import * as fs from 'fs';//used
import * as chalk from 'chalk';//used
import * as commnder from "commander";
import { OfficeAddinTelemetry } from "./officeAddinTelemetry";
/*var addInTelemetry: any;
export async function start(command: commnder.Command) {
        var readlineSync = require('readline-sync');
        var key = readlineSync.question('What is the instrumentation key?');
        addInTelemetry = new OfficeAddinTelemetry(key);
        
}export async function stop(command: commnder.Command) {
    if(addInTelemetry !== {}){
    addInTelemetry.setTelemetryOff();
    }
}
export async function telemetryStatus(command: commnder.Command) {
    addInTelemetry.isTelemetryOn();
}*/