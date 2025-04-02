import { RuleTester } from "@typescript-eslint/rule-tester";
import rule from "../../src/rules/no-office-initialize";

const ruleTester = new RuleTester();

ruleTester.run("no-office-initialize", rule, {
  valid: [
    {
      code: `Office.onReady();`,
    },
    {
      code: `Office.onReady(function(info) {
          console.log(info);
        });`,
    },
    {
      code: `Office.onReady()
          .then(function() {
            console.log("Testing Office.onReady followed by .then");
          });`,
    },
    {
      code: `(async () => {
          await Office.onReady();
            console.log("Testing Office.onReady followed by await");
        })();`,
    },
  ],
  invalid: [
    {
      code: `Office.initialize = function () {};`,
      errors: [{ messageId: "noOfficeInitialize" }],
    },
    {
      code: `Office.initialize = function () {
          console.log("Testing Office.initialize with normal function");
        };`,
      errors: [{ messageId: "noOfficeInitialize" }],
    },
    {
      code: `Office.initialize = function (reason) {
          $(document).ready(function () {
              console.log(reason);
          });
        };`,
      errors: [{ messageId: "noOfficeInitialize" }],
    },
    {
      code: `Office.initialize = () => {
          console.log("Testing Office.initialize with arrow function");
        };`,
      errors: [{ messageId: "noOfficeInitialize" }],
    },
  ],
});
