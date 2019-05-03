import * as lockfile from "lockfile";
import * as defaults from "./defaults";

export async function acquireLock(file: string, wait: number = defaults.maxWaitTime): Promise<void> {
    try {
        await new Promise(function(resolve, reject) {
            lockfile.lock(file, {wait}, (err) => { err ? reject(err) : resolve(); });
        });
    } catch (err) {
        throw new Error(`Another process using office-addin-dev-certs took too long to finish.\n${err}`);
    }
}

export function releaseLock(file: string) {
    lockfile.unlockSync(file);
}
