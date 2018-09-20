import * as winreg from "winreg";

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
