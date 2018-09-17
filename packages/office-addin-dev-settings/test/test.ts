import * as assert from "assert";
import * as mocha from "mocha";
import { enableDebugging } from "../src/dev-settings";

const addinId = "9982ab78-55fb-472d-b969-b52ed294e173";

describe("DevSettings", function() {
  describe("enableDebugging", function() {
    it("should enable debugging", async function() {
      await enableDebugging(addinId);
    });
  });
});
