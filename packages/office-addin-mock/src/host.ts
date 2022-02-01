import { OfficeApp } from "office-addin-manifest";
import { ObjectData } from "./objectData";

export function getHostType(
  object: ObjectData | undefined
): OfficeApp | undefined {
  let validHost = undefined;
  if (object && object["host"]) {
    Object.values(OfficeApp).forEach((host: string) => {
      if (object["host"] === host) {
        validHost = object["host"];
      }
    });
  }
  return validHost;
}
