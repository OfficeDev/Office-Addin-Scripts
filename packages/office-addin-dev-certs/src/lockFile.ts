import * as lockfile from "lockfile";
import * as defaults from "./defaults";

export class LockFile{
    static aquiredLock: boolean = false;

    async acquireLock(file: string, wait: number = defaults.maxWaitTime): Promise<void> {
        if(LockFile.aquiredLock) return;

        try {
            await new Promise(function(resolve, reject) {
                lockfile.lock(file, {wait}, (err) => { err ? reject(err) : resolve(); });
            });
            LockFile.aquiredLock = true;
        } catch (err) {
            throw new Error(`Another process using office-addin-dev-certs took too long to finish.\n${err}`);
        }
    }
    
    releaseLock(file: string) {
        lockfile.unlockSync(file);
        LockFile.aquiredLock = false;
    }
}
