import * as assert from "assert";
import * as childProcess from "child_process";
import * as fsExtra from "fs-extra";
import * as mocha from "mocha";
import * as path from "path";
import {generateCertificates} from "../src/generate";
import {installCaCertificate} from "../src/install";
import * as uninstall from "../src/uninstall";
import * as verify from "../src/verify";

describe("office-addin-dev-certs", function() {
    const sinon = require("sinon");
    const mkcert = require("mkcert");
    const fs = require("fs");
    let sandbox = sinon.createSandbox();
    const testCertificateDir = "certs";
    const testCaCertificatePath = path.join(testCertificateDir, "ca.crt");
    const testCertificatePath = path.join(testCertificateDir, "localhost.crt");
    const testKeyPath = path.join(testCertificateDir, "localhost.key");
    const cert = {cert: "cert", key: "key"};
    describe("generate-tests", function() {
        beforeEach(function() {
            sandbox = sinon.createSandbox();
        });
        afterEach(function() {
            sandbox.restore();
        });
        it("certificate not installed, ensureDir fails case", async function() {
            const createCert = sandbox.fake();
            const error = "test error";
            sandbox.stub(fsExtra, "ensureDirSync").throws(error);
            try {
                await generateCertificates(testCaCertificatePath, testCertificatePath, testKeyPath, 30);
                // expecting exception
                assert.strictEqual(0, 1);
            } catch (err) {
                assert.strictEqual(err.toString().includes("Unable to create directory"), true);
            }
            assert.strictEqual(createCert.callCount, 0);
        });
        it("certificate not installed, createCA fails case", async function() {
            const createCert = sandbox.fake();
            const error = "test error";
            sandbox.stub(mkcert, "createCA").rejects(error);
            sandbox.stub(mkcert, "createCert").callsFake(createCert);
            try {
                await generateCertificates(testCaCertificatePath, testCertificatePath, testKeyPath, 30);
                // expecting exception
                assert.strictEqual(0, 1);
            } catch (err) {
                assert.strictEqual(err.toString().includes("Unable to generate CA certificate"), true);
            }
            assert.strictEqual(createCert.callCount, 0);
        });
        it("certificate not installed, createCert fails case", async function() {
            const createCert = sandbox.fake();
            const error = "test error";
            sandbox.stub(mkcert, "createCA").resolves(cert);
            sandbox.stub(mkcert, "createCert").rejects(error);
            try {
                await generateCertificates(testCaCertificatePath, testCertificatePath, testKeyPath, 30);
                // expecting exception
                assert.strictEqual(0, 1);
            } catch (err) {
                assert.strictEqual(err.toString().includes("Unable to generate localhost certificate"), true);
            }
            assert.strictEqual(createCert.callCount, 0);
        });
        it("certificate not installed, fs write sync fails case", async function() {
            const createCert = sandbox.fake();
            const error = "test error";
            sandbox.stub(mkcert, "createCA").resolves(cert);
            sandbox.stub(mkcert, "createCert").resolves(cert);
            sandbox.stub(fs, "writeSync").throws(error);

            try {
                await generateCertificates(testCaCertificatePath, testCertificatePath, testKeyPath, 30);
                // expecting exception
                assert.strictEqual(0, 1);
            } catch (err) {
                assert.strictEqual(err.toString().includes("Unable to write generated certificates"), true);
            }
            assert.strictEqual(createCert.callCount, 0);
        });
        it("certificate not installed case", async function() {
            const createCert = sandbox.fake();
            const writeSync = sandbox.fake();
            sandbox.stub(mkcert, "createCA").resolves(cert);
            sandbox.stub(mkcert, "createCert").resolves(cert);
            sandbox.stub(fs, "writeSync").callsFake(writeSync);
            await generateCertificates(testCaCertificatePath, testCertificatePath, testKeyPath, 30);
            assert.strictEqual(writeSync.callCount, 3);
        });
    });
    describe("install-tests", function() {
        const uninstallCaCertificate = sandbox.fake();
        beforeEach(function() {
            sandbox = sinon.createSandbox();
        });
        afterEach(function() {
            sandbox.restore();
        });
        it("execSync fail case", async function() {
            const error = {stderr : "test error"};
            sandbox.stub(uninstall, "uninstallCaCertificate").callsFake(uninstallCaCertificate);
            sandbox.stub(childProcess, "execSync").throws(error);
            try {
                await installCaCertificate(testCaCertificatePath);
            } catch (err) {
                assert.strictEqual(err.message, "Unable to install the CA certificate. test error");
            }
        });
        it("install success case", async function() {
            const execSync = sandbox.fake();
            sandbox.stub(uninstall, "uninstallCaCertificate").callsFake(uninstallCaCertificate);
            sandbox.stub(childProcess, "execSync").callsFake(execSync);
            try {
                await installCaCertificate(testCaCertificatePath);
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
                await uninstall.uninstallCaCertificate();
            } catch (err) {
                assert.strictEqual(err.message, "Unable to uninstall the CA certificate.\ntest error");
            }
        });
        it("install success case", async function() {
            const execSync = sandbox.fake();
            sandbox.stub(childProcess, "execSync").callsFake(execSync);
            try {
                await uninstall.uninstallCaCertificate();
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
                await verify.verifyCaCertificate();
            } catch (err) {
                assert.strictEqual(err.message, "test error");
            }
        });
        it("certificate not found in trusted store case", async function() {
            const execSync = sandbox.fake.returns("");
            sandbox.stub(childProcess, "execSync").callsFake(execSync);
            try {
                const ret = await verify.verifyCaCertificate();
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
                const ret = await verify.verifyCaCertificate();
                assert.strictEqual(execSync.callCount, 1);
                assert.strictEqual(ret, true);
            } catch (err) {
                // not expecting any exception
                assert.strictEqual(0, 1);
            }
        });
    });
});
