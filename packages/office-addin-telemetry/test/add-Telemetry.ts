import * as assert from 'assert';
import * as fs from "fs";
import { OfficeAddinTelemetry } from "../src/officeAddinTelemetry";
import * as path from 'path';
import {  describe, before, it } from 'mocha';
const addInTelemetry = new OfficeAddinTelemetry("de0d9e7c-1f46-4552-bc21-4e43e489a015");
    describe('addTelemetry', () => {
            it('should add object to telemetry', () => {
                var test1 ={};
                var compare ={"Tester2": true }
                addInTelemetry.addTelemetry(test1, "Tester2", true);
                assert(JSON.stringify(test1) === JSON.stringify(compare));
            });
          });
