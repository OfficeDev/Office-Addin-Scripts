# Office-Addin-Mock

  This library provides an easier way to unit test the Office JavaScript library (hereafter, "Office-js") API.
  This library does not depend on Office, so it doesn't tests actual interactions with it.
  
  It aims to solve problems that arise when trying to mock the API`s objects:

- Office-js APIs need to be loaded by an HTML file, so they are not available before loading.
- Some test APIs may require the entire object to be mocked, which can have more than 100 properties, making mocking not feasible.
- Tests need to preserve the order of the functions `load` or `sync`, which are difficult to test because stateless test APIs do not support easily adding state variables to handle those functions.

## Installation

Install `office-addin-mock`

```
npm i office-addin-mock --save-dev
```

## Usage

The examples used here will be using [Mocha](mochajs.org/) testing framework. Any JavaScript framework should work, feel free to use others if needed.

Import `office-addin-mock` to your testing file:

```Typescript
import { OfficeMockObject } from "office-addin-mock";
```

Create an object structure to represent the mock object. Override all the properties and methods you want to use.

```Typescript
const MockData = {
  workbook: {
    range: {
      address: "C2",
    },
    getSelectedRange: function () {
      return this.range;
    },
  },
};
```

In your test code, create an `OfficeMockObject` with an argument of the object you created:

```Typescript
const contextMock = new OfficeMockObject(MockData) as any;
```

You can now use this newly created object as a mock of the original Office-js object.

## Examples

1. Testing a function that calls an Office-js API:

```Typescript
async function getSelectedRangeAddress(context: Excel.RequestContext): Promise<string> {
  const range: Excel.Range = context.workbook.getSelectedRange();

  range.load("address");
  await context.sync();

  return range.address;
}

const MockData = {
  workbook: {
    range: {
      address: "C2",
    },
    getSelectedRange: function () {
      return this.range;
    },
  },
};

describe(`getSelectedRangeAddress`, function () {
  it("Returns correct value", async function () {
    const contextMock = new OfficeMockObject(MockData) as any;
    assert.strictEqual(await getSelectedRangeAddress(contextMock), "C2");
  });
});
```

2. Testing a function that uses the global Excel variable:

```Typescript
async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range: Excel.Range = context.workbook.getSelectedRange();

      // Load the range address
      range.load("address");

      // Update the fill color
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
};

describe(`Run`, function () {
  it("Using json", async function () {
    const excelMock = new OfficeMockObject(MockData) as any;
    excelMock.addMockFunction("run", async function (callback) {
      await callback(excelMock.context);
    });
    global.Excel = excelMock;
    await run();
    assert.strictEqual(excelMock.context.workbook.range.format.fill.color, "yellow");
  });
});
```
