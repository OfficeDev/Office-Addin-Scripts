import { RuleTester } from '@typescript-eslint/rule-tester'
import rule from '../../src/rules/call-sync-after-load';

const ruleTester = new RuleTester({
  parser: '@typescript-eslint/parser',
});

ruleTester.run('call-sync-after-load', rule, {
  valid: [ 
    {
      code: `
        var property = worksheet.getItem("sheet");
        property.load("values");
        await context.sync();
        console.log(property.values);`
    },
    {
      code: `
        var fakeGet = worksheet.notAGetFunction("props");
        await context.sync();
        fakeGet.load("props");
        console.log(fakeGet.props);`
    },
    {
      code: `
        var fakeGet = worksheet.notAGetFunction("props");
        await context.sync();
        property.load("props");
        console.log(property.props);`
    },
    {
      code: `
        var table = worksheet.getTables();
        return context.sync().then(function () {
          table.delete();
        });`
    },
    {
      code: `
        var range = worksheet.getSelectedRange();
        range.getCell(0,0);`
    },
		{
			code: `
			  var range = worksheet.getSelectedRange();
			  range.load("font/fill/color, address");
			  await context.sync();
			  console.log(range.font.fill.color);
			  console.log(range.address);`
		},
    {
      code: `
        var range = worksheet.getSelectedRange();
        range.load("*");
        await context.sync();
        console.log(range.address);`
    },
    {
      code: `
        var range = worksheet.getSelectedRange();
        range.load();
        await context.sync();
        console.log(range.address);`
    },
    {
      code: `
        var property = worksheet.getItem("sheet");
        context.load(property, "values");
        await context.sync();
        console.log(property.values);`
    },
    {
      code: `
        var property = worksheet.getItem("sheet");
        context.load(property);
        await context.sync();
        console.log(property.values);`
    },
  ],
  invalid: [
    {
      code: `
        var property = worksheet.getItem("sheet");
        await context.sync();
        property.load("values");
        console.log(property.values);`,
      errors: [{ messageId: "callSyncAfterLoad", data: { name: "property", loadValue: "values" }}]
    },
    {
      code: `
        var property = worksheet.getItem("sheet");
        property.load("values");
        await context.sync();
        console.log(property.values);
        property.load("length");
        console.log(property.length);`,
      errors: [{ messageId: "callSyncAfterLoad", data: { name: "property", loadValue: "length" }}]
    },
    {
      code: `
        var range = worksheet.getSelectedRange();
        range.load(["font/fill/color", "address"]);
        console.log(range.font.fill.color);`,
      errors: [{ messageId: "callSyncAfterLoad", data: { name: "range", loadValue: "font/fill/color" }}]
    },
    {
      code: `
        var range = worksheet.getSelectedRange();
        range.load("*");
        console.log(range.address);`,
      errors: [{ messageId: "callSyncAfterLoad", data: { name: "range", loadValue: "address" }}]
    },
    {
      code: `
        var range = worksheet.getSelectedRange();
        range.load();
        console.log(range.address);`,
      errors: [{ messageId: "callSyncAfterLoad", data: { name: "range", loadValue: "address" }}]
    },
    {
      code: `
        var property = worksheet.getItem("sheet");
        await context.sync();
        context.load(property, "values");
        console.log(property.values);`,
      errors: [{ messageId: "callSyncAfterLoad", data: { name: "property", loadValue: "values" }}]
    },
    {
      code: `
        var property = worksheet.getItem("sheet");
        context.load(property, "values");
        console.log(property.values);`,
      errors: [{ messageId: "callSyncAfterLoad", data: { name: "property", loadValue: "values" }}]
    },
    {
      code: `
        var property = worksheet.getItem("sheet");
        await context.sync();
        context.load(property);
        console.log(property.values);`,
      errors: [{ messageId: "callSyncAfterLoad", data: { name: "property", loadValue: "values" }}]
    },
    {
      code: `
        var property = worksheet.getItem("sheet");
        context.load(property);
        console.log(property.values);`,
      errors: [{ messageId: "callSyncAfterLoad", data: { name: "property", loadValue: "values" }}]
    },
  ]
});
