import * as os from "os";
import * as path from "path";

export const telemetryJsonFilePath: string = path.join(os.homedir(), "/officeAddinTelemetry.json");
export const groupName = "Office-Addin-Telemetry";
