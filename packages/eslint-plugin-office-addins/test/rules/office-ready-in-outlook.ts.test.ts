import { ESLintUtils } from '@typescript-eslint/experimental-utils'
import rule from '../../src/rules/office-ready-in-outlook';

const ruleTester = new ESLintUtils.RuleTester({
  parser: '@typescript-eslint/parser',
});

ruleTester.run('office-ready-in-outlook', rule, {
  valid: [ 
    {
      code: `
        Office.onReady();`
    },
    {
      code: `
        Office.onReady(function(info) {
          if (info.host === Office.HostType.Excel) {
              // Do Excel-specific initialization (for example, make add-in task pane's
              // appearance compatible with Excel "green").
          }
          if (info.platform === Office.PlatformType.PC) {
              // Make minor layout changes in the task pane.
          }
          console.log("Office.js is now ready");
        });`
    },
    {
      code: `
        Office.onReady()
          .then(function() {
              if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
                  console.log("Sorry, this add-in only works with newer versions of Excel.");
              }
          });`
    },
    {
      code: `
        (async () => {
          await Office.onReady();
          if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
              console.log("Sorry, this add-in only works with newer versions of Excel.");
          }
        })();`
    }
  ],
  invalid: [
    {
      code: `
        Office.initialize = function () {};`,
      errors: [{ messageId: "officeOnReadyInOutlook" }]
    },
    {
      code: `
        Office.initialize = function () {
          if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
              console.log("Sorry, this add-in only works with newer versions of Excel.");
          }
        };`,
      errors: [{ messageId: "officeOnReadyInOutlook" }]
    },
    {
      code: `
        Office.initialize = function (reason) {
          $(document).ready(function () {
              switch (reason) {
                  case 'inserted': console.log('The add-in was just inserted.');
                  case 'documentOpened': console.log('The add-in is already part of the document.');
              }
          });
        };`,
      errors: [{ messageId: "officeOnReadyInOutlook" }]
    },
    {
      code: `
        Office.initialize = () => {
          console.log("Testing arrow functions");
        };`,
      errors: [{ messageId: "officeOnReadyInOutlook" }]
    }
  ]
});