import * as fs from "fs";
import * as defaults from "./defaults";
import { ensureCertificatesAreInstalled } from "./install";

interface IHttpsServerOptions {
    cert: Buffer;
    key: Buffer;
}

export async function gethttpsServerOptions(): Promise<IHttpsServerOptions> {
    await ensureCertificatesAreInstalled();

    const httpsServerOptions = {} as IHttpsServerOptions;
    try {
        httpsServerOptions.cert = fs.readFileSync(defaults.localhostCertificatePath);
    } catch (err) {
        throw new Error(`Unable to read the certificate file.\n${err}`);
    }

    try {
        httpsServerOptions.key = fs.readFileSync(defaults.localhostKeyPath);
    } catch (err) {
        throw new Error(`Unable to read the certificate key.\n${err}`);
    }

    return httpsServerOptions;
}
