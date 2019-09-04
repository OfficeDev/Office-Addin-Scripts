// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as childProcess from "child_process";
import * as commander from "commander";
import * as inquirer from "inquirer";
import * as mocha from "mocha";
import * as sinon from "sinon";
import * as appcontainer from "../../src/appcontainer";
import * as commands from "../../src/commands";
const addinId = "9982ab78-55fb-472d-b969-b52ed294e173";

if (process.platform === "win32") { // only windows is supported
  describe("Appcontainer edgewebview tests", async function() {
    const appcontainerName = "edgewebview";
    let sandbox: sinon.SinonSandbox;

    beforeEach(function() {
      sandbox = sinon.createSandbox();
    });
    afterEach(function() {
      sandbox.restore();
    });
    it("loopback already enabled", async function() {
      const command: commander.Command = new commander.Command();
      command.loopback = true;
      command.yes = true;
      const appcontaineId = await appcontainer.getAppcontainerNameFromManifestPath(appcontainerName);
      await appcontainer.addLoopbackExemptionForAppcontainer(appcontaineId);
      const addLoopbackExemptionForAppcontainer = sandbox.spy(appcontainer, "addLoopbackExemptionForAppcontainer");
      await commands.appcontainer(appcontainerName, command);
      assert.strictEqual(addLoopbackExemptionForAppcontainer.callCount, 0);
      await appcontainer.removeLoopbackExemptionForAppcontainer("Microsoft.win32webviewhost_cw5n1h2txyewy");
    });
    it("loopback not enabled, user doesn't gives consent", async function() {
      const command: commander.Command = new commander.Command();
      command.loopback = true;
      sandbox.stub(inquirer, "prompt").resolves({didUserConfirm: false});
      const exec = sandbox.spy(childProcess, "exec");
      await commands.appcontainer(appcontainerName, command);
      assert.strictEqual(exec.callCount, 1); // because one query to check if loopback status
    });
    it("loopback not enabled, user gives consent", async function() {
      const command: commander.Command = new commander.Command();
      command.loopback = true;
      const appcontaineId = await appcontainer.getAppcontainerNameFromManifestPath(appcontainerName);
      await appcontainer.removeLoopbackExemptionForAppcontainer(appcontaineId);
      sandbox.stub(inquirer, "prompt").resolves({didUserConfirm: true});
      const exec = sandbox.spy(childProcess, "exec");
      await commands.appcontainer(appcontainerName, command);
      assert.strictEqual(exec.callCount, 2);
    });
  });
}
