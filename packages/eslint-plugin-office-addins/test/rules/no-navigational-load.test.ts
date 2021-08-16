import { ESLintUtils } from '@typescript-eslint/experimental-utils'
import rule from '../../src/rules/no-navigational-load';

const ruleTester = new ESLintUtils.RuleTester({
	parser: '@typescript-eslint/parser',
});

ruleTester.run('no-navigational-load', rule, {
	valid: [ 
		{
			code: `
                var range = worksheet.getRange("A1");
                range.load("borders/fill/size");
                console.log(range.borders.fill.size);`
		},
		{
			code: `
                var sheetName = 'Sheet1';
                var rangeAddress = 'A1:B2';
                var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);  
                myRange.load('address');
                context.sync()
                        .then(function () {
                        console.log (myRange.address);   // ok
                        });`
		},
		{
			code: `
                var selectedRange = context.workbook.getSelectedRange();
                selectedRange.load('text'); // Scalar`
		},
		{
			code: `
                var selectedRange = context.workbook.getSelectedRange();
                selectedRange.load('values');
                if(selectedRange.values === [2]){}`
		},
		{
			code: `
                var myRange = context.workbook.worksheets.notAGet();
                myRange.load('notAProperty');
                var test = myRange.notAProperty;`
    	},
		{
			code: `
                var range = context.workbook.getRange();
                range.load({borders: { fill: { color: true } } });
                if (range.borders.fill.color);`
    	},
		{
			code: `
                var range = context.workbook.getRange();
                range.borders.fill.load("color");
                console.log(range.borders.fill.color);`
    	},
	],
	invalid: [
		{
			code: `
                var property = worksheet.getItem("sheet");
                property.load('thisIsNotAProperty'); //Not a property`,
			errors: [{ messageId: "navigationalLoad", data: { loadValue: "thisIsNotAProperty" } }]
		},
		{
			code: `
                var property = worksheet.getItem("sheet");
                property.load('styles'); // Navigational`,
			errors: [{ messageId: "navigationalLoad", data: { loadValue: "styles" } }]
		},
		{
			code: `
                var range = worksheet.getRange("A1");
                range.load('fill'); // Navigational`,
			errors: [{ messageId: "navigationalLoad", data: { loadValue: "fill" } }]
		},
		{
			code: `
                var range = worksheet.getRange("A1");
                range.load("borders/fill");
                console.log(range.borders.fill); // Navigational`,
			errors: [{ messageId: "navigationalLoad", data: { loadValue: "borders/fill" } }]
		},
		{
			code: `
                var range = worksheet.getRange("A1");
                range.load({borders: { fill: true } });`,
			errors: [{ messageId: "navigationalLoad", data: { loadValue: "borders/fill" } }]
		},
		{
			code: `
                var range = context.workbook.getRange();
                range.borders.load("fill");
                console.log(range.borders.fill);`,
			errors: [{ messageId: "navigationalLoad", data: { loadValue: "fill" } }]
		},
	]
});
