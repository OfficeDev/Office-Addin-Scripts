import * as defaults from "./defaults";

export const VERIFY_SUCCESS_MSG = `You have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}.\nKey: ${defaults.localhostKeyPath}.`;
export const VERIFY_FAILURE_MSG = `Use "install" for trusted access to https://localhost`;
export const INSTALL_SUCCESS_MSG = `You now have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}.\nKey: ${defaults.localhostKeyPath}.`;
export const ALREADY_INSTALL_MSG = `You already have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}.\nKey: ${defaults.localhostKeyPath}.`;
export const UNINSTALL_SUCESSS_MSG = `You no longer have trusted access to https://localhost.`;
