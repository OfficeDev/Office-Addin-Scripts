declare module "read-package-json-fast" {
  declare type ScriptsObject = { [key: string]: string };
  export default function readPackageJson(
    filePath: string /* eslint-disable-line @typescript-eslint/no-unused-vars */
  ): Promise<{ scripts: ScriptsObject; [key: string]: unknown }>;
}
