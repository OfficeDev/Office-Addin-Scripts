import * as assert from "assert";
import * as commander from "commander";
import * as inquirer from "inquirer";
import * as mocha from "mocha";
import * as sinon from "sinon";
import * as appcontainer from "../../src/appcontainer";
import * as commands from "../../src/commands";
const addinId = "9982ab78-55fb-472d-b969-b52ed294e173";

describe("Appcontainer edgewebview tests", async function() {
  const appcontainerName = "edgewebview";
  let sandbox: sinon.SinonSandbox;
  const command: commander.Command = new commander.Command();
  command.loopback = true;

  beforeEach(function() {
    sandbox = sinon.createSandbox();
  });
  afterEach(function() {
    sandbox.restore();
  });
  it("loopback already enabled", async function() {
    command.loopback = true;
    const appcontaineId = await commands.getAppcontainerName(appcontainerName);
    await appcontainer.addLoopbackExemptionForAppcontainer(appcontaineId);
    const addLoopbackExemptionForAppcontainer = sandbox.spy(appcontainer, "addLoopbackExemptionForAppcontainer");
    await commands.appcontainer(appcontainerName, command);
    assert.strictEqual(addLoopbackExemptionForAppcontainer.callCount, 0);
    await appcontainer.removeLoopbackExemptionForAppcontainer("Microsoft.win32webviewhost_cw5n1h2txyewy");
  });
  it("loopback not enabled, user doesn't gives consent", async function() {
    sandbox.stub(inquirer, "prompt").resolves({didUserConfirm: false});
    const addLoopbackExemptionForAppcontainer = sandbox.spy(appcontainer, "addLoopbackExemptionForAppcontainer");
    await commands.appcontainer(appcontainerName, command);
    assert.strictEqual(addLoopbackExemptionForAppcontainer.callCount, 0);
  });
  it("loopback not enabled, user gives consent", async function() {
    const appcontaineId = await commands.getAppcontainerName(appcontainerName);
    await appcontainer.removeLoopbackExemptionForAppcontainer(appcontaineId);
    sandbox.stub(inquirer, "prompt").resolves({didUserConfirm: true});
    const addLoopbackExemptionForAppcontainer = sandbox.spy(appcontainer, "addLoopbackExemptionForAppcontainer");
    await commands.appcontainer(appcontainerName, command);
    assert.strictEqual(addLoopbackExemptionForAppcontainer.callCount, 1);
  });
});
