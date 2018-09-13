import * as child from 'child_process';

const spawn = child.spawn;
const HKEY_CURRENT_USER = 'HKEY_CURRENT_USER';
const HKEY_LOCAL_MACHINE = 'HKEY_LOCAL_MACHINE';
const REG_SZ = 'REG_SZ';
const REG_DWORD = 'REG_DWORD';

export const DeveloperSettingsRegistryKey = HKEY_CURRENT_USER + '\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer';

function addValue(path: string, value: string, type: string, data: string) {
  spawn('reg', ['add', path, '/v', value, '/t', type, '/d', data, '/f'], { stdio: 'inherit' });
}

export function addBooleanValue(path: string, value: string, data: boolean) {
  addValue(path, value, 'REG_DWORD', data ? '1' : '0')
}

export function addNumberValue(path: string, value: string, data: number) {
  addValue(path, value, REG_DWORD, data.toString());
}

export function addStringValue(path: string, value: string, data: string) {
  addValue(path, value, REG_SZ, data);
}

export function deleteValue(path: string, value: string) {
  spawn('reg', ['delete', path, '/v', value, '/f'], { stdio: 'inherit' });
}

function deleteKey(path: string) {
  spawn('reg', ['delete', path, '/f'], { stdio: 'inherit' });
}

export function getDeveloperSettingsRegistryKey(addinId: string) {
  if (!addinId) throw 'addinId is required.';
  if (typeof addinId !== 'string') throw 'addinId should be a string.';
  return DeveloperSettingsRegistryKey + '\\' + addinId;
}

export function deleteDeveloperSettingsRegistryKey(addinId: string) {
  const key = getDeveloperSettingsRegistryKey(addinId);
  deleteKey(key);
}
