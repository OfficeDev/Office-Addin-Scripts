import { ObjectData } from "./objectData";

export enum Host {
  excel = "excel",
  onenote = "onenote",
  outlook = "outlook",
  powerpoint = "powerpoint",
  project = "project",
  visio = "visio",
  word = "word",
  unknow = "unknow",
}

export function getHostType(object: ObjectData | undefined): Host {
  let validHost = Host.unknow;
  if (object && object["host"]) {
    Object.values(Host).forEach((host: string) => {
      if (object["host"] === host) {
        validHost = object["host"];
      }
    });
  }
  return validHost;
}
