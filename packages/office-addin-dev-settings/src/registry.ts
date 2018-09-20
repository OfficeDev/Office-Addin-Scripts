import * as winreg from "winreg";

export class RegistryKey {
  public winreg: winreg.Registry;

  get path(): string {
    return this.winreg.path;
  }

  constructor(path: string) {
    if (!path) { throw new Error("Please provide a registry key path."); }

    const index = path.indexOf("\\");

    if (index <= 0) { throw new Error(`The registry key path is not valid: "${path}".`); }

    const hive = path.substring(0, index);
    const subpath = path.substring(index);

    this.winreg = new winreg({ hive: normalizeRegistryHive(hive), key: subpath });
  }
}

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

  public get isNumberType(): boolean {
    return isNumberType(this.type);
  }

  public get isStringType(): boolean {
    return isStringType(this.type);
  }

  constructor(key: string, name: string, type: string, data: string) {
    this.key = key;
    this.name = name;
    this.type = type;
    this.data = data;
  }
}

async function addValue(key: RegistryKey, value: string, type: string, data: string): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    const onError = (err: any) => {
      if (err) {
        reject(new Error(`Unable to set registry value "${value}" to "${data}" (${type}) for key "${key.path}".\n${err}`));
      } else {
        resolve();
      }
    };

    try {
      key.winreg.set(value, type, data, onError);
    } catch (err) {
      onError(err);
    }
  });
}

export async function addBooleanValue(key: RegistryKey, value: string, data: boolean): Promise<void> {
  return addValue(key, value, winreg.REG_DWORD, data ? "1" : "0");
}

export async function addNumberValue(key: RegistryKey, value: string, data: number): Promise<void> {
  return addValue(key, value, winreg.REG_DWORD, data.toString());
}

export async function addStringValue(key: RegistryKey, value: string, data: string): Promise<void> {
  return addValue(key, value, winreg.REG_SZ, data);
}

export async function deleteKey(key: RegistryKey): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    const onError = (err: any) => {
      if (err) {
        // it's not an error if the key does not exist
        if (err instanceof Error && err.message.match("unable to find the specified registry key")) {
          resolve();
        } else {
          reject(new Error(`Unable to delete registry key "${key.path}".\n${err}`));
        }
      } else {
        resolve();
      }
    };

    try {
      key.winreg.destroy(onError);
    } catch (err) {
      onError(err);
    }
  });
}

export async function deleteValue(key: RegistryKey, value: string): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    const onError = (err: any) => {
      if (err) {
        // it's not an error if the key or value does not exist
        if (err instanceof Error && err.message.match("unable to find the specified registry key or value")) {
          resolve();
        } else {
          reject(new Error(`Unable to delete registry value "${value}" in key "${key.path}".\n${err}`));
        }
      } else {
        resolve();
      }
    };

    try {
      key.winreg.remove(value, onError);
    } catch (err) {
      onError(err);
    }
  });
}

export async function doesKeyExist(key: RegistryKey): Promise<boolean> {
  return new Promise<boolean>((resolve, reject) => {
    const onError = (err: any, exists: boolean = false) => {
      if (err) {
        reject(new Error(`Unable to determine if registry key exists: "${key.path}".\n${err}`));
      } else {
        resolve(exists);
      }
    };

    try {
      key.winreg.keyExists(onError);
    } catch (err) {
      onError(err);
    }
  });
}

export async function doesValueExist(key: RegistryKey, value: string): Promise<boolean> {
  return new Promise<boolean>((resolve, reject) => {
    const onError = (err: any, exists: boolean = false) => {
      if (err) {
        reject(new Error(`Unable to determine if registry value "${value}" exists for key "${key.path}".\n${err}`));
      } else {
        resolve(exists);
      }
    };

    try {
      key.winreg.valueExists(value, onError);
    } catch (err) {
      onError(err);
    }
  });
}

export async function getNumberValue(key: RegistryKey, value: string): Promise<number | undefined> {
  const registryValue: RegistryValue | undefined = await getValue(key, value);

  return (registryValue && registryValue.isNumberType) ? parseInt(registryValue.data, undefined) : undefined;
}

export async function getStringValue(key: RegistryKey, value: string): Promise<string | undefined> {
  const registryValue: RegistryValue | undefined = await getValue(key, value);

  return (registryValue && registryValue.isStringType) ? registryValue.data : undefined;
}

export async function getValue(key: RegistryKey, value: string): Promise<RegistryValue | undefined> {
  return new Promise<RegistryValue>((resolve, reject) => {
    const onError = (err: any, item?: winreg.RegistryItem) => {
      if (err) {
        resolve(undefined);
      } else {
        resolve(item ? new RegistryValue(key.path, item.name, item.type, item.value) : undefined);
      }
    };

    try {
      key.winreg.get(value, onError);
    } catch (err) {
      onError(err);
    }
  });
}

export function isNumberType(registryType: string) {
  // NOTE: REG_QWORD is not included as a number type since it cannot be returned as a "number".
  return (registryType === RegistryTypes.REG_DWORD);
}

export function isStringType(registryType: string) {
  switch (registryType) {
    case RegistryTypes.REG_SZ:
      return true;
    default:
      return false;
  }
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
