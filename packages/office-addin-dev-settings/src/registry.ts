import { spawn } from "child_process";

const REG_SZ: string = "REG_SZ";
const REG_DWORD: string = "REG_DWORD";

function addValue(path: string, value: string, type: string, data: string): void {
  spawn("reg", ["add", path, "/v", value, "/t", type, "/d", data, "/f"], { stdio: "inherit" });
}

export function addBooleanValue(path: string, value: string, data: boolean): void {
  addValue(path, value, "REG_DWORD", data ? "1" : "0");
}

export function addNumberValue(path: string, value: string, data: number): void {
  addValue(path, value, REG_DWORD, data.toString());
}

export function addStringValue(path: string, value: string, data: string): void {
  addValue(path, value, REG_SZ, data);
}

export function deleteValue(path: string, value: string): void {
  spawn("reg", ["delete", path, "/v", value, "/f"], { stdio: "inherit" });
}

export function deleteKey(path: string): void {
  spawn("reg", ["delete", path, "/f"], { stdio: "inherit" });
}
