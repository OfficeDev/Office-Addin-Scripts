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
        httpsServerOptions.key = fs.readFileSync(defaults.localhostKeyPath);
    } catch (err) {
        throw new Error(`Error occured while reading certificate files.\n${err}`);
    }

    return httpsServerOptions;
}
