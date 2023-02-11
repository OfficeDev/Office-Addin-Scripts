# Office-Addin-Mock

This library makes it much easier to unit test the Office JavaScript API. It works with any JavaScript unit testing framework.

For details about the library and how to use it, see [Unit testing in Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/testing/unit-testing).

## Installation

Install `office-addin-mock` in the root of an add-in project with this command:

```
npm i office-addin-mock --save-dev
```

## Examples

For basic examples that use the [Jest](https://jestjs.io) framework, see [Unit testing in Office Add-ins - Examples](https://learn.microsoft.com/office/dev/add-ins/testing/unit-testing#examples). Below are some examples using other frameworks.

### Testing with Mocha for Excel platform

```Javascript
import { OfficeMockObject } from "office-addin-mock";

function run() {
  try {
    await Excel.run(async (context) => {
      const range: Excel.Range = context.workbook.getSelectedRange();
      range.load("address");

      // Update the cell color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

const MockData = {
  context: {
    workbook: {
      range: {
        address: "G4",
        format: {
          fill: {},
        },
      },
      getSelectedRange: function () {
        return this.range;
      },
    },
  },
  run: async function(callback) {
    await callback(this.context);
  },
};

describe(`Run`, function () {
  it("Excel", async function () {
    const excelMock = new OfficeMockObject(MockData) as any;
    global.Excel = excelMock;
    await run();
    assert.strictEqual(excelMock.context.workbook.range.format.fill.color, "yellow");
  });
});
```

### Testing a function with Mocha for PowerPoint platform

```Javascript
import { OfficeMockObject } from "office-addin-mock";

async function run() {
  const options = { coercionType: Office.CoercionType.Text };

  await Office.context.document.setSelectedDataAsync(" ", options);
  await Office.context.document.setSelectedDataAsync("Hello World!", options);
}

const PowerPointMockData = {
  context: {
    document: {
      setSelectedDataAsync: function (data, options?) {
        this.data = data;
        this.options = options;
      },
    },
  },
  CoercionType: {
    Text: {},
  },
};

describe(`PowerPoint`, function () {
  it("Run", async function () {
    const officeMock = new OfficeMockObject(PowerPointMockData);
    global.Office = officeMock;

    await run();

    assert.strictEqual(officeMock.context.document.data, "Hello World!");
  });
});

```

## Reference

### OfficeMockObject class

Represents a mock Office object.

#### Constructor

The object parameter provides initial values for the mock object. (Optional)
The host parameter identifies the host of the tests. (Optional)

```
constructor(object?: Object, host?: OfficeApp | undefined); 
```

Host can be any of the following:

```Javascript
OfficeApp {
  Excel = "excel",
  OneNote = "onenote",
  Outlook = "outlook",
  PowerPoint = "powerpoint",
  Project = "project",
  Word = "word",
}
```

#### Methods

##### load

Mock implementation of the `load` method in the application-specific Office.js APIs.

- The `propertyArgument` specifies the properties that should be loaded.  

```
load(propertyArgument: string | string[] | Object): void;
```

##### sync

Mock replacement for the `sync` method in the application-specific Office.js APIs.

```
sync(): void;
```
