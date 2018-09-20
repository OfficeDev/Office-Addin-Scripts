import * as winreg from "winreg";

export class RegistryTypes {
  public static readonly REG_BINARY: string = winreg.REG_BINARY;
  public static readonly REG_DWORD: string = winreg.REG_DWORD;
  public static readonly REG_EXPAND_SZ: string = winreg.REG_EXPAND_SZ;
  public static readonly REG_MULTI_SZ: string = winreg.REG_MULTI_SZ;
  public static readonly REG_NONE: string = winreg.REG_NONE;
  public static readonly REG_QWORD: string = winreg.REG_QWORD;
  public static readonly REG_SZ: string = winreg.REG_SZ;
}

export class RegistryValue {
  public key: string;
  public name: string;
  public type: string;
  public data: string;

  constructor(key: string, name: string, type: string, data: string) {
    this.key = key;
    this.name = name;
    this.type = type;
    this.data = data;
  }
}

async function addValue(path: string, value: string, type: string, data: string): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    const onError = (err: any) => {
      if (err) {
        reject(new Error(`Unable to set registry value "${value}" to "${data}" (${type}) for key "${path}".\n${err}`));
      } else {
        resolve();
      }
    };

    try {
      registryKey(path).set(value, type, data, onError);
    } catch (err) {
      onError(err);
    }
  });
}

export async function addBooleanValue(path: string, value: string, data: boolean): Promise<void> {
  return addValue(path, value, winreg.REG_DWORD, data ? "1" : "0");
}

export async function addNumberValue(path: string, value: string, data: number): Promise<void> {
  return addValue(path, value, winreg.REG_DWORD, data.toString());
}

export async function addStringValue(path: string, value: string, data: string): Promise<void> {
  return addValue(path, value, winreg.REG_SZ, data);
}

export async function deleteKey(path: string): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    const onError = (err: any) => {
      if (err) {
        reject(new Error(`Unable to delete registry key "${path}".\n${err}`));
      } else {
        resolve();
      }
    };

    try {
      registryKey(path).destroy(onError);
    } catch (err) {
      onError(err);
    }
  });
}

export async function deleteValue(path: string, value: string): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    const onError = (err: any) => {
      if (err) {
        reject(new Error(`Unable to delete registry value "${value}" in key "${path}".\n${err}`));
      } else {
        resolve();
      }
    };

    try {
      registryKey(path).remove(value, onError);
    } catch (err) {
      onError(err);
    }
  });
}

export async function doesKeyExist(path: string): Promise<boolean> {
  return new Promise<boolean>((resolve, reject) => {
    const onError = (err: any, exists: boolean = false) => {
      if (err) {
        reject(new Error(`Unable to determine if registry key exists: "${path}".\n${err}`));
      } else {
        resolve(exists);
      }
    };

    try {
      registryKey(path).keyExists(onError);
    } catch (err) {
      onError(err);
    }
  });
}

export async function doesValueExist(path: string, value: string): Promise<boolean> {
  return new Promise<boolean>((resolve, reject) => {
    const onError = (err: any, exists: boolean = false) => {
      if (err) {
        reject(new Error(`Unable to determine if registry value "${value}" exists for key "${path}".\n${err}`));
      } else {
        resolve(exists);
      }
    };

    try {
      registryKey(path).valueExists(value, onError);
    } catch (err) {
      onError(err);
    }
  });
}

export async function getValue(path: string, value: string): Promise<RegistryValue | undefined> {
  return new Promise<RegistryValue>((resolve, reject) => {
    const onError = (err: any, item?: winreg.RegistryItem) => {
      if (err) {
        resolve(undefined);
      } else {
        resolve(new RegistryValue(path, item!.name, item!.type, item!.value));
      }
    };

    try {
      registryKey(path).get(value, onError);
    } catch (err) {
      onError(err);
    }
  });
}

function normalizeRegistryHive(hive: string): string {
  switch (hive) {
    case "HKEY_CURRENT_USER":
      return winreg.HKCU;
    case "HKEY_LOCAL_MACHINE":
      return winreg.HKLM;
    case "HKEY_CLASSES_ROOT":
      return winreg.HKCR;
    case "HKEY_CURRENT_CONFIG":
      return winreg.HKCC;
    case "HKEY_USERS":
      return winreg.HKU;
    default:
      return hive;
  }
}

export function registryKey(path: string) {
  if (!path) { throw new Error("Please provide a registry key path."); }

  const index = path.indexOf("\\");

  if (index <= 0) { throw new Error(`The registry key path is not valid: "${path}".`); }

  const hive = path.substring(0, index);
  const subpath = path.substring(index);

  const key = new winreg({ hive: normalizeRegistryHive(hive), key: subpath });

  return key;
}
