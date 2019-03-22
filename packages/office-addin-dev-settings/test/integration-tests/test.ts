import * as assert from "assert";
import * as commander from "commander";
import * as inquirer from "inquirer";
import * as mocha from "mocha";
import * as sinon from "sinon";
import * as appcontainer from "../src/appcontainer";
import * as commands from "../src/commands"
import * as devSettings from "../src/dev-settings";
const addinId = "9982ab78-55fb-472d-b969-b52ed294e173";


describe("Appcontainer", async function() {
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
    const isLoopbackExemptionForAppcontainer = sinon.fake.returns(true);
    const addLoopbackExemptionForAppcontainer = sinon.fake();
    sandbox.stub(appcontainer, "isLoopbackExemptionForAppcontainer").callsFake(isLoopbackExemptionForAppcontainer);
    sandbox.stub(appcontainer, "addLoopbackExemptionForAppcontainer").callsFake(addLoopbackExemptionForAppcontainer);
    await commands.appcontainer("EdgeWebView", command);
    assert.strictEqual(isLoopbackExemptionForAppcontainer.calledWith("Microsoft.win32webviewhost_cw5n1h2txyewy"), true);
    assert.strictEqual(addLoopbackExemptionForAppcontainer.callCount, 0);
  });
  it("loopback not enabled, user doesn't gives consent", async function() {
    const command: commander.Command = new commander.Command();
    command.loopback = true;
    const isLoopbackExemptionForAppcontainer = sinon.fake.returns(false);
    const addLoopbackExemptionForAppcontainer = sinon.fake();
    sandbox.stub(appcontainer, "isLoopbackExemptionForAppcontainer").callsFake(isLoopbackExemptionForAppcontainer);
    sandbox.stub(appcontainer, "addLoopbackExemptionForAppcontainer").callsFake(addLoopbackExemptionForAppcontainer);
    sandbox.stub(inquirer, "prompt").resolves({didUserConfirm: false});
    await commands.appcontainer("EdgeWebView", command);
    assert.strictEqual(isLoopbackExemptionForAppcontainer.calledWith("Microsoft.win32webviewhost_cw5n1h2txyewy"), true);
    assert.strictEqual(addLoopbackExemptionForAppcontainer.callCount, 0);
  });
  it("loopback not enabled, user gives consent", async function() {
    const command: commander.Command = new commander.Command();
    command.loopback = true;
    const isLoopbackExemptionForAppcontainer = sinon.fake.returns(false);
    const addLoopbackExemptionForAppcontainer = sinon.fake();
    sandbox.stub(appcontainer, "isLoopbackExemptionForAppcontainer").callsFake(isLoopbackExemptionForAppcontainer);
    sandbox.stub(appcontainer, "addLoopbackExemptionForAppcontainer").callsFake(addLoopbackExemptionForAppcontainer);
    sandbox.stub(inquirer, "prompt").resolves({didUserConfirm: true});
    await commands.appcontainer("EdgeWebView", command);
    assert.strictEqual(isLoopbackExemptionForAppcontainer.calledWith("Microsoft.win32webviewhost_cw5n1h2txyewy"), true);
    assert.strictEqual(addLoopbackExemptionForAppcontainer.callCount, 1);
  });
});