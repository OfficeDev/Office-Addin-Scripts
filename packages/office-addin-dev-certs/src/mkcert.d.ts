declare module "url" {
    export interface ICertificateInfo {
        cert: string;
        key: string;
    }

    export async function createCA({organization, countryCode, state, locality, validityDays}: {organization: string, countryCode: string, state: string, locality: string, validityDays: number}): ICertificateInfo;
    export async function createCert({domains, validityDays, caKey, caCert} : {domains: string, validityDays: number, caKey: string, caCert: string}): ICertificateInfo;
}