# Office-Addin-Mock

  This library provides an easier way to unit test the Office JavaScript library (hereafter, "Office-js") API.
  This library does not depend on Office, so it doesn't test actual interactions with it.
  
  It aims to solve problems that arise when trying to mock the API's objects:

- Office-js APIs need to be loaded by an HTML file, so they are not available before loading.
- Some test APIs may require the entire object to be mocked, which can have more than 100 properties, making mocking not feasible.
- Tests need to preserve the order of the functions `load` or `sync`, which are difficult to test because stateless test APIs do not support easily adding state variables to handle those functions.

## Installation

Install `office-addin-mock`

```
npm i office-addin-mock --save-dev
```

## Usage

The examples used here will be using [Mocha](mochajs.org/) and [Jest](https://jestjs.io/) testing frameworks. Any JavaScript framework should work, feel free to use others if needed.

Import `office-addin-mock` to your testing file:

```Javascript
import { OfficeMockObject } from "office-addin-mock";
```

Create an object structure to represent the mock object. Override all the properties and methods you want to use.

```Javascript
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

```Javascript
const contextMock = new OfficeMockObject(MockData);
```

You can now use this newly created object as a mock of the original Office-js object.

## Examples

1. Testing with Jest for Excel platform:

```Javascript
async function getSelectedRangeAddress(context) {
  const range = context.workbook.getSelectedRange();

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

test("Excel", async function () {
  const contextMock = new OfficeMockObject(MockData);
  expect(await getSelectedRangeAddress(contextMock)).toBe("C2");
});
```

2. Testing with Mocha for Excel platform:

```Javascript
function run() {
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
  it("Excel", async function () {
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

3. Testing with Jest for Word platform:

```Javascript
async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}

const WordMockData = {
  context: {
    document: {
      body: {
        paragraph: {
          font: {},
          text: "",
        },
        insertParagraph: function (paragraphText, insertLocation) {
          this.paragraph.text = paragraphText;
          this.paragraph.insertLocation = insertLocation;
          return this.paragraph;
        },
      },
    },
  },
  InsertLocation: {
    end: "End",
  },
};

test("Word", async function () {
  const wordMock = new officeAddinMock.OfficeMockObject(WordMockData);
  wordMock.addMockFunction("run", async function (callback) {
    await callback(wordMock.context);
  });
  global.Word = wordMock;

  await run();

  expect(wordMock.context.document.body.paragraph.font.color).toBe("blue");
});
```

4. Testing a function with Mocha for Powerpoint platform:

```Javascript
async function run() {
  /**
   * Insert your PowerPoint code here
   */
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
