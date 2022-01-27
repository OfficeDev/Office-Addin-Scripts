import { ObjectData } from "./objectData";

export enum Host {
  excel = "excel",
  outlook = "outlook",
  powerpoint = "powerpoint",
  word = "word",
  other = "other",
  notFound = "notFound",
}

export function getHostType(object: ObjectData | undefined): Host {
  if (object && object["host"]) {
    let validHost = Host.other;
    Object.values(Host).forEach((host: string) => {
      if (object["host"] === host) {
        validHost = object["host"];
      }
    });
    return validHost;
  }
  return Host.notFound;
}
