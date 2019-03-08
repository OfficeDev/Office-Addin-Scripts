import * as assert from "assert";
import * as childProcess from "child_process";
import * as mocha from "mocha";
import * as path from "path";
import {generateCertificates} from "../src/generate";
import {installCaCertificate} from "../src/install";
import {uninstallCaCertificate} from "../src/uninstall";
import {verifyCaCertificate} from "../src/verify";
import * as verify from "../src/verify";

describe("office-addin-dev-certs", function() {
    const sinon = require("sinon");
    const mkcert = require("mkcert");
    const testCertificateDir = "";
    let sandbox = sinon.createSandbox();
    describe("generate-tests", function() {
        beforeEach(function() {
            sandbox = sinon.createSandbox();
        });
        afterEach(function() {
            sandbox.restore();
        });
        it("certificate already installed case", async function() {
            const verifyCertificate = sandbox.fake.returns(true);
            const createCA = sandbox.fake();
            sandbox.stub(mkcert, "createCA").callsFake(createCA);
            sandbox.stub(verify, "verifyCaCertificate").callsFake(verifyCertificate);
            await generateCertificates(path.join(testCertificateDir, "ca.crt"), path.join(testCertificateDir, "localhost.crt"), path.join(testCertificateDir, "localhost.key"), false);
            assert.strictEqual(verifyCertificate.callCount, 1);
            assert.strictEqual(createCA.callCount, 0);
        });
        it("certificate not installed, createCA fails case", async function() {
            const verifyCertificate = sandbox.fake.returns(false);
            const createCert = sandbox.fake();
            const cert = {cert: "cert", key: "key"};
            sandbox.stub(verify, "verifyCaCertificate").callsFake(verifyCertificate);
            sandbox.stub(mkcert, "createCA").rejects(cert);
            sandbox.stub(mkcert, "createCert").callsFake(createCert);
            await generateCertificates(path.join(testCertificateDir, "ca.crt"), path.join(testCertificateDir, "localhost.crt"), path.join(testCertificateDir, "localhost.key"), false);
            assert.strictEqual(verifyCertificate.callCount, 1);
            assert.strictEqual(createCert.callCount, 0);
        });
        it("certificate not installed case", async function() {
            const verifyCertificate = sandbox.fake.returns(false);
            const createCert = sandbox.fake();
            const cert = {cert: "cert", key: "key"};
            sandbox.stub(verify, "verifyCaCertificate").callsFake(verifyCertificate);
            sandbox.stub(mkcert, "createCA").resolves(cert);
            sandbox.stub(mkcert, "createCert").callsFake(createCert);
            await generateCertificates(path.join(testCertificateDir, "ca.crt"), path.join(testCertificateDir, "localhost.crt"), path.join(testCertificateDir, "localhost.key"), false);
            assert.strictEqual(verifyCertificate.callCount, 1);
            assert.strictEqual(createCert.callCount, 1);
        });
    });
    describe("install-tests", function() {
        beforeEach(function() {
            sandbox = sinon.createSandbox();
        });
        afterEach(function() {
            sandbox.restore();
        });
        it("execSync fail case", async function() {
            const error = {stderr : "test error"};
            sandbox.stub(childProcess, "execSync").throws(error);
            try {
                await installCaCertificate(path.join(testCertificateDir, "ca.crt"));
            } catch (err) {
                assert.strictEqual(err.message, "test error");
            }
        });
        it("install success case", async function() {
            const execSync = sandbox.fake();
            sandbox.stub(childProcess, "execSync").callsFake(execSync);
            try {
                await installCaCertificate(path.join(testCertificateDir, "ca.crt"));
                assert.strictEqual(execSync.callCount, 1);
            } catch (err) {
                // not expecting any exception
                assert.strictEqual(0, 1);
            }
        });
    });
    describe("uninstall-tests", function() {
        beforeEach(function() {
            sandbox = sinon.createSandbox();
        });
        afterEach(function() {
            sandbox.restore();
        });
        it("execSync fail case", async function() {
            const error = {stderr : "test error"};
            sandbox.stub(childProcess, "execSync").throws(error);
            try {
                await uninstallCaCertificate();
            } catch (err) {
                assert.strictEqual(err.message, "test error");
            }
        });
        it("install success case", async function() {
            const execSync = sandbox.fake();
            sandbox.stub(childProcess, "execSync").callsFake(execSync);
            try {
                await uninstallCaCertificate();
                assert.strictEqual(execSync.callCount, 1);
            } catch (err) {
                // not expecting any exception
                assert.strictEqual(0, 1);
            }
        });
    });
    describe("verify-tests", function() {
        beforeEach(function() {
            sandbox = sinon.createSandbox();
        });
        afterEach(function() {
            sandbox.restore();
        });
        it("execSync fail case", async function() {
            const error = {stderr : "test error"};
            sandbox.stub(childProcess, "execSync").throws(error);
            try {
                await verifyCaCertificate();
            } catch (err) {
                assert.strictEqual(err.message, "test error");
            }
        });
        it("certificate not found in trusted store case", async function() {
            const execSync = sandbox.fake.returns("");
            sandbox.stub(childProcess, "execSync").callsFake(execSync);
            try {
                const ret = await verifyCaCertificate();
                assert.strictEqual(execSync.callCount, 1);
                assert.strictEqual(ret, false);
            } catch (err) {
                // not expecting any exception
                assert.strictEqual(0, 1);
            }
        });
        it("certificate found in trusted store case", async function() {
            const execSync = sandbox.fake.returns("Certificate details");
            sandbox.stub(childProcess, "execSync").callsFake(execSync);
            try {
                const ret = await verifyCaCertificate();
                assert.strictEqual(execSync.callCount, 1);
                assert.strictEqual(ret, true);
            } catch (err) {
                // not expecting any exception
                assert.strictEqual(0, 1);
            }
        });
    });
});
