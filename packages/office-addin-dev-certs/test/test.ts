// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as childProcess from "child_process";
import * as fsExtra from "fs-extra";
import * as mocha from "mocha";
import * as path from "path";
import * as defaults from "../src/defaults";
import * as generate from "../src/generate";
import { getHttpsServerOptions } from "../src/httpsServerOptions";
import * as install from "../src/install";
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
    const lockWaitTime = 100;

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
                await generate.generateCertificates(testCaCertificatePath, testCertificatePath, testKeyPath, 30);
                // expecting exception
                assert.strictEqual(0, 1);
            } catch (err) {
                assert.strictEqual(err.toString().includes("Unable to create the directory"), true);
            }
            assert.strictEqual(createCert.callCount, 0);
        });
        it("certificate not installed, createCA fails case", async function() {
            const createCert = sandbox.fake();
            const error = "test error";
            sandbox.stub(mkcert, "createCA").rejects(error);
            sandbox.stub(mkcert, "createCert").callsFake(createCert);
            try {
                await generate.generateCertificates(testCaCertificatePath, testCertificatePath, testKeyPath, 30);
                // expecting exception
                assert.strictEqual(0, 1);
            } catch (err) {
                assert.strictEqual(err.toString().includes("Unable to generate the CA certificate"), true);
            }
            assert.strictEqual(createCert.callCount, 0);
        });
        it("certificate not installed, createCert fails case", async function() {
            const createCert = sandbox.fake();
            const error = "test error";
            sandbox.stub(mkcert, "createCA").resolves(cert);
            sandbox.stub(mkcert, "createCert").rejects(error);
            try {
                await generate.generateCertificates(testCaCertificatePath, testCertificatePath, testKeyPath, 30);
                // expecting exception
                assert.strictEqual(0, 1);
            } catch (err) {
                assert.strictEqual(err.toString().includes("Unable to generate the localhost certificate"), true);
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
                await generate.generateCertificates(testCaCertificatePath, testCertificatePath, testKeyPath, 30);
                // expecting exception
                assert.strictEqual(0, 1);
            } catch (err) {
                assert.strictEqual(err.toString().includes("Unable to write generated certificates"), true);
            }
            assert.strictEqual(createCert.callCount, 0);
        });
        it("certificate not installed case", async function() {
            const writeSync = sandbox.fake();
            sandbox.stub(mkcert, "createCA").resolves(cert);
            sandbox.stub(mkcert, "createCert").resolves(cert);
            sandbox.stub(fs, "writeSync").callsFake(writeSync);
            sandbox.stub(fs, "existsSync").returns(false);
            await generate.generateCertificates(testCaCertificatePath, testCertificatePath, testKeyPath, 30);
            assert.strictEqual(writeSync.callCount, 3);
            fsExtra.removeSync(testCertificateDir);
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
                await install.installCaCertificate(testCaCertificatePath);
            } catch (err) {
                assert.strictEqual(err.message, "Unable to install the CA certificate. test error");
            }
        });
        it("certificate already installed case", async function() {
            const execSync = sandbox.fake();
            sandbox.stub(childProcess, "execSync").callsFake(execSync);
            sandbox.stub(verify, "isCaCertificateInstalled").returns(true);
            await install.installCaCertificate(testCaCertificatePath);
            assert.strictEqual(execSync.callCount, 0);
        });
        it("install success case", async function() {
            const execSync = sandbox.fake();
            sandbox.stub(childProcess, "execSync").callsFake(execSync);
            sandbox.stub(verify, "isCaCertificateInstalled").returns(false);
            try {
                await install.installCaCertificate(testCaCertificatePath);
                assert.strictEqual(execSync.callCount, 1);
            } catch (err) {
                // not expecting any exception
                assert.strictEqual(0, 1);
            }
        });
        if (process.platform === "win32") {
            it("with --machine option", async function() {
                const execSync = sandbox.fake();
                const machine = true;
                sandbox.stub(childProcess, "execSync").callsFake(execSync);
                sandbox.stub(verify, "isCaCertificateInstalled").returns(false);
                await install.installCaCertificate(testCaCertificatePath, machine);
                assert.strictEqual(execSync.callCount, 1);
                assert.strictEqual(execSync.calledWith(`powershell -ExecutionPolicy Bypass scripts\\install.ps1 LocalMachine '${testCaCertificatePath}'`), true);
            });
            it("without --machine option", async function() {
                const execSync = sandbox.fake();
                const machine = false;
                sandbox.stub(childProcess, "execSync").callsFake(execSync);
                sandbox.stub(verify, "isCaCertificateInstalled").returns(false);
                await install.installCaCertificate(testCaCertificatePath, machine);
                assert.strictEqual(execSync.callCount, 1);
                assert.strictEqual(execSync.calledWith(`powershell -ExecutionPolicy Bypass scripts\\install.ps1 CurrentUser '${testCaCertificatePath}'`), true);
            });
        }
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
            sandbox.stub(verify, "isCaCertificateInstalled").returns(true);
            try {
                await uninstall.uninstallCaCertificate();
                assert.strictEqual(0, 1);
            } catch (err) {
                assert.strictEqual(err.message, "Unable to uninstall the CA certificate.\ntest error");
            }
        });
        it("uninstall success case", async function() {
            const execSync = sandbox.fake();
            sandbox.stub(childProcess, "execSync").callsFake(execSync);
            sandbox.stub(verify, "isCaCertificateInstalled").returns(true);
            try {
                await uninstall.uninstallCaCertificate();
                assert.strictEqual(execSync.callCount, 1);
            } catch (err) {
                // not expecting any exception
                assert.strictEqual(0, 1);
            }
        });
        if (process.platform === "win32") {
            it("with --machine option", async function() {
                const isCaCertificateInstalled = sandbox.fake.returns(true);
                const execSync = sandbox.fake();
                const machine = true;
                sandbox.stub(childProcess, "execSync").callsFake(execSync);
                sandbox.stub(verify, "isCaCertificateInstalled").callsFake(isCaCertificateInstalled);
                await uninstall.uninstallCaCertificate(machine);
                assert.strictEqual(execSync.callCount, 1);
                assert.strictEqual(execSync.calledWith(`powershell -ExecutionPolicy Bypass scripts\\uninstall.ps1 LocalMachine '${defaults.caCertificatePath}' '${defaults.certificateName}'`), true);
            });
            it("without --machine option", async function() {
                const isCaCertificateInstalled = sandbox.fake.returns(true);
                const execSync = sandbox.fake();
                const machine = false;
                sandbox.stub(childProcess, "execSync").callsFake(execSync);
                sandbox.stub(verify, "isCaCertificateInstalled").callsFake(isCaCertificateInstalled);
                await uninstall.uninstallCaCertificate(machine);
                assert.strictEqual(execSync.callCount, 1);
                assert.strictEqual(execSync.calledWith(`powershell -ExecutionPolicy Bypass scripts\\uninstall.ps1 CurrentUser '${defaults.caCertificatePath}' '${defaults.certificateName}'`), true);
            });
        }
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
                await verify.isCaCertificateInstalled();
            } catch (err) {
                assert.strictEqual(err.message, "test error");
            }
        });
        it("certificate not found in trusted store case", async function() {
            let execSync;
            if (process.platform === "darwin") {
                execSync = sandbox.fake.throws("test error");
            } else {
                execSync = sandbox.fake.returns("");
            }
            sandbox.stub(childProcess, "execSync").callsFake(execSync);
            try {
                const ret = await verify.isCaCertificateInstalled();
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
                const ret = await verify.isCaCertificateInstalled();
                assert.strictEqual(execSync.callCount, 1);
                assert.strictEqual(ret, true);
            } catch (err) {
                // not expecting any exception
                assert.strictEqual(0, 1);
            }
        });
    });
    describe("getHttpsServerOptions-tests", function() {
        beforeEach(function() {
            sandbox = sinon.createSandbox();
        });
        afterEach(function() {
            sandbox.restore();
        });
        it("already installed certificate case", async function() {
            const ensureCertificatesAreInstalled = sandbox.fake();
            sandbox.stub(install, "ensureCertificatesAreInstalled").callsFake(ensureCertificatesAreInstalled);
            sandbox.stub(fs, "readFileSync").returns("test");
            const serverOptions = await getHttpsServerOptions();
            assert.strictEqual(serverOptions.cert, "test");
            assert.strictEqual(serverOptions.key, "test");
        });
        it("already installed certificate case, readsync fail case", async function() {
            const ensureCertificatesAreInstalled = sandbox.fake();
            sandbox.stub(install, "ensureCertificatesAreInstalled").callsFake(ensureCertificatesAreInstalled);
            sandbox.stub(fs, "readFileSync").throws("test error");
            try {
                await getHttpsServerOptions();
                // expecting exception
                assert.strictEqual(0, 1);
            } catch (err) {

                assert.strictEqual(err.toString().includes("Unable to read the certificate file."), true);
            }
        });
    });
    describe("deleteCertificateFiles-tests", function() {
        const certificateDirectory: string = path.resolve("./certs");
        const testFile = "test.txt";
        const testFilePath = path.join(certificateDirectory, testFile);
        const localhostCertificatePath = path.join(certificateDirectory, defaults.localhostCertificateFileName);
        const localhostKeyPath = path.join(certificateDirectory, defaults.localhostKeyFileName);
        const caCertificatePath = path.join(certificateDirectory, defaults.caCertificateFileName);
        beforeEach(function() {
            fsExtra.ensureDirSync(certificateDirectory);
            fsExtra.outputFileSync(localhostCertificatePath, "test");
            fsExtra.outputFileSync(localhostKeyPath, "test");
            fsExtra.outputFileSync(caCertificatePath, "test");
        });
        afterEach(function() {
            fsExtra.removeSync(certificateDirectory);
        });
        it("extrafile in certificate folder case", async function() {
            fsExtra.outputFileSync(testFilePath, "test");
            await uninstall.deleteCertificateFiles(certificateDirectory);
            assert.strictEqual(fsExtra.existsSync(certificateDirectory), true);
            assert.strictEqual(fsExtra.existsSync(testFilePath), true);
            assert.strictEqual(fsExtra.existsSync(localhostCertificatePath), false);
            assert.strictEqual(fsExtra.existsSync(localhostKeyPath), false);
            assert.strictEqual(fsExtra.existsSync(caCertificatePath), false);
        });
        it("clean certificate folder case", async function() {
            await uninstall.deleteCertificateFiles(certificateDirectory);
            assert.strictEqual(fsExtra.existsSync(certificateDirectory), false);
        });
    });
});
