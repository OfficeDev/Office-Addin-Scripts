import * as lockfile from "lockfile";
import * as defaults from "./defaults";

export class LockFile{
    static aquiredLock: boolean = false;
    filePath: string;

    constructor(filePath: string = defaults.devCertsLockPath){
        this.filePath = filePath;
    }

    async acquireLock(wait: number = defaults.maxWaitTime): Promise<void> {
        // If same process tries to aquire lock again, just return
        if(LockFile.aquiredLock) return;

        try {
            await new Promise((resolve, reject) => {
                lockfile.lock(this.filePath, {wait}, (err) => { err ? reject(err) : resolve(); });
            });
            LockFile.aquiredLock = true;
    } catch (err) {
            throw new Error(`Another process using office-addin-dev-certs took too long to finish.\n${err}`);
        }
    }
    
    releaseLock() {
        lockfile.unlockSync(this.filePath);
        LockFile.aquiredLock = false;
    }
}
