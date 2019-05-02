import * as lf from "lockfile";
import * as defaults from "./defaults";


export function lock(file: string, wait: number = defaults.maxWaitTimeToAcquireLock): Promise<void> {
    return new Promise(function(resolve, reject) {
        lf.lock(file, {wait: wait}, function(err) {
            if (err) {
                reject(err);
            }
            resolve();
        });
    });
}

export function unlockSync(file: string) {
    lf.unlockSync(file);
}
