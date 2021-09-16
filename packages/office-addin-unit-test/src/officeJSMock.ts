import { usageDataObject } from "./defaults";

export class OfficeJSMock {
  constructor(json?: OfficeObject) {
    this.properties = new Map<string, OfficeJSMock>();
    this.loaded = false;
    this.resetValue(undefined);
    if (json) {
      this.populate(json);
    }
  }

  // Adds a function to OfficeJSMock
  addMockFunction(methodName: string, functionality?: Function) {
    this[methodName] = functionality ? functionality : function () {};
  }

  // Adds a OfficeJSMock to OfficeJSMock
  addMockObject(objectName: string) {
    const officeJSMock = new OfficeJSMock();
    officeJSMock.isMockObject = true;
    this.properties.set(objectName, officeJSMock);
    this[objectName] = this.properties.get(objectName);
  }

  // load of Office.js API
  load(property: string) {
    this.loadMultipleProperties(property);
  }

  // Adds a property of any type to OfficeJSMock
  setMock(propertyName: string, value: unknown) {
    if (!this.properties.has(propertyName)) {
      const officeJSMock = new OfficeJSMock();
      officeJSMock.isMockObject = false;
      this.properties.set(propertyName, officeJSMock);
    }
    this.properties.get(propertyName)?.resetValue(value);
    this[propertyName] = this.properties.get(propertyName)?.value;
  }

  // sync of Office.js API
  sync() {
    this.properties.forEach((property: OfficeJSMock, key: string) => {
      property.sync();
      this.assignValue(key);
    });
    if (this.loaded) {
      this.value = this.valueBeforeLoaded;
    }
  }

  private assignValue(propertyName: string) {
    if (this.properties.get(propertyName)?.isMockObject) {
      this[propertyName] = this.properties.get(propertyName);
    } else {
      this[propertyName] = this.properties.get(propertyName)?.value;
    }
  }

  private loadCalled() {
    this.loaded = true;
    this.value = `Error, context.sync() was not called`;
  }

  private loadMultipleProperties(property: string) {
    property
      .replace(/\s/g, "")
      .split(",")
      .forEach((individualProperties: string) => {
        this.loadNavigational(individualProperties);
      });
  }

  private loadNavigational(propertyName: string) {
    const properties: Array<string> = propertyName.split("/");
    let navigationalOfficeJSMock: OfficeJSMock = this;

    // Iterating through navigational properties
    for (let i = 0; i < properties.length - 1; i++) {
      const property = properties[i];

      const checkingUndefined =
        navigationalOfficeJSMock.properties.get(property);
      if (checkingUndefined) {
        navigationalOfficeJSMock = checkingUndefined;
      } else {
        throw new Error(
          `Navigational property ${property} needs to be present in object model before load is called.`
        );
      }
    }
    const scalarProperty: string = properties[properties.length - 1];
    navigationalOfficeJSMock.loadScalar(scalarProperty);
  }

  private loadScalar(propertyName: string) {
    if (this.properties.has(propertyName)) {
      this.properties.get(propertyName)?.loadCalled();
      this.assignValue(propertyName);

      this.properties
        .get(propertyName)
        ?.properties.forEach((property: OfficeJSMock) => {
          property.loadCalled();
        });
    } else {
      throw new Error(
        `Property ${propertyName} needs to be present in object model before load is called.`
      );
    }
  }

  private populate(json: OfficeObject) {
    try {
      Object.keys(json).forEach((property: string) => {
        if (typeof json[property] === "object") {
          this.addMockObject(property);
          this[property].populate(json[property]);
          this[property].setName(property);
        } else if (typeof json[property] === "function") {
          this.addMockFunction(property, json[property]);
        } else {
          this.setMock(property, json[property]);
        }
      });
      usageDataObject.reportSuccess("populate()");
    } catch (err: any) {
      usageDataObject.reportException("populate()", err);
    }
  }

  private resetValue(value: unknown) {
    this.value = `Error, property was not loaded`;
    this.valueBeforeLoaded = value;
  }

  private properties: Map<string, OfficeJSMock>;
  private loaded: boolean;
  private value: unknown;
  private valueBeforeLoaded: unknown;
  private isMockObject: boolean | undefined;
  /* eslint-disable-next-line */
  [key: string]: any;
}

class OfficeObject {
  /* eslint-disable-next-line */
  [key: string]: any;
}
