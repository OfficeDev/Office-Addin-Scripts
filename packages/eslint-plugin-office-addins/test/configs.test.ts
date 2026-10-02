/// <reference types="jest" />
import { ESLint } from "eslint";

// eslint-disable-next-line @typescript-eslint/no-require-imports
const plugin = require("../src/main");

const code = `export function f(): number {
  const unused = 1;
  return 2;
}
`;

describe("recommended config", () => {
  for (const filePath of ["src/file.ts", "src/file.tsx"]) {
    it(`lints ${filePath} with the TypeScript parser`, async () => {
      const eslint = new ESLint({
        overrideConfigFile: true,
        overrideConfig: plugin.configs.recommended,
      });
      const [result] = await eslint.lintText(code, { filePath });
      expect(result.messages.map((message) => message.ruleId)).toEqual([
        "@typescript-eslint/no-unused-vars",
      ]);
    });
  }
});
