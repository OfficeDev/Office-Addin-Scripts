import { usageDataObject } from "./defaults";

export class OfficeJSMock {
  constructor(object?: OfficeObject) {
    this.properties = new Map<string, OfficeJSMock>();
    this.loaded = false;
    this.resetValue(undefined);
    if (object) {
      this.populate(object);
    }
  }

  // Adds a function to OfficeJSMock
  addMockFunction(methodName: string, functionality?: Function) {
    this[methodName] = functionality ? functionality : function () {};
  }

  // Adds an object to OfficeJSMock
  addMock(objectName: string) {
    const officeJSMock = new OfficeJSMock();
    officeJSMock.isMockObject = true;
    this.properties.set(objectName, officeJSMock);
    this[objectName] = this.properties.get(objectName);
  }

  // Mock replacement of the load of Office.js API
  load(propertyArgument: string) {
    propertyArgument
    .replace(/\s/g, "")
    .split(",")
    .forEach((completeProperties: string) => {
      this.loadNavigational(completeProperties);
    });
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

  // Mock replacement for the sync of Office.js API
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
    if (!this.loaded) {
      this.loaded = true;
      this.value = `Error, context.sync() was not called`;
    }
  }

  private loadNavigational(completePropertyName: string) {
    const properties: Array<string> = completePropertyName.split("/");
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

  private loadScalar(scalarPropertyName: string) {
    if (this.properties.has(scalarPropertyName)) {
      this.properties.get(scalarPropertyName)?.loadCalled();
      this.assignValue(scalarPropertyName);

      this.properties
        .get(scalarPropertyName)
        ?.properties.forEach((property: OfficeJSMock) => {
          property.loadCalled();
        });
    } else {
      throw new Error(
        `Property ${scalarPropertyName} needs to be present in object model before load is called.`
      );
    }
  }

  private populate(object: OfficeObject) {
    try {
      Object.keys(object).forEach((property: string) => {
        if (typeof object[property] === "object") {
          this.addMock(property);
          this[property].populate(object[property]);
          this[property].setName(property);
        } else if (typeof object[property] === "function") {
          this.addMockFunction(property, object[property]);
        } else {
          this.setMock(property, object[property]);
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
    this.loaded = false;
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
