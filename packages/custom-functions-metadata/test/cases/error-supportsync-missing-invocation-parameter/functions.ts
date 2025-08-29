// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Test missing invocation parameter type when @supportSync is used and has RichAPI calls
 * @param x string
 * @customfunction
 * @supportSync
 */
async function customFunctionRequestContext(x: string) {
    const context = new Excel.RequestContext();
}

/**
 * Test missing invocation parameter type when @supportSync is used and has RichAPI calls
 * @param x string
 * @customfunction
 * @supportSync
 */
async function customFunctionExcelRun(x: string) {
    await Excel.run(async (context) => {
        const range = context.workbook.tables.getItem("Table1").getRange();
        range.load("values");
        await context.sync();
    });
}
