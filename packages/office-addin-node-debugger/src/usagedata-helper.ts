import * as defaults from './defaults';
import * as usageData from "office-addin-usage-data";
const usageDataObject: usageData.OfficeAddinUsageData = new usageData.OfficeAddinUsageData(defaults.usageDataOptions);

export function sendUsageDataSuccessEvent(method: string) {
    const usageDataInfo = {
        Method: [method],
        Platform: [process.platform],
        Succeeded: [true]
    }
    usageDataObject.reportEvent(defaults.usageDataProjectName, usageDataInfo);
}

export function sendUsageDataCustomEvent(usageDataInfo: Object) {
    usageDataObject.reportEvent(defaults.usageDataProjectName, usageDataInfo);
}

export function sendUsageDataException(method: string, error: string) {
    usageDataObject.reportError(defaults.usageDataProjectName, new Error(`${method} error: ${error}`));
}