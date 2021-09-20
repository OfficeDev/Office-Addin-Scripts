import { usageDataObject } from "./defaults";

/**
 * Creates an office-js mockable object
 * @param object Object structure to provide initial values for the mock object (Optional)
 */
export class OfficeMockObject {
  constructor(object?: OfficeObject) {
    this.properties = new Map<string, OfficeMockObject>();
    this.loaded = false;
    this.resetValue(undefined);
    if (object) {
      this.populate(object);
    }
  }

  /**
   * Adds a function to OfficeMockObject. The function can be accessed by simply calling `this.methodName`
   * @param methodName Method name of the function to be added
   * @param functionality Function to be added to the object. A blank function will be added if no argument is provided (Optional)
   */
  addMockFunction(methodName: string, functionality?: Function) {
    this[methodName] = functionality ? functionality : function () {};
  }

  /**
   * Adds a mock object to OfficeMockObject. The object can be accessed by simply calling `this.objectName`
   * @param objectName Object name of the object to be added
   */
  addMock(objectName: string) {
    const officeMockObject = new OfficeMockObject();
    officeMockObject.isMockObject = true;
    this.properties.set(objectName, officeMockObject);
    this[objectName] = this.properties.get(objectName);
  }

  /**
   * Mock replacement of the load of Office.js API
   * @param propertyArgument Argument of the load call. Will load any properties in the argument
   */
  load(propertyArgument: string | string[]) {
    if (typeof propertyArgument === "string") {
      this.loadMultipleProperties(propertyArgument);
    } else {
      propertyArgument.forEach((property: string) => {
        this.loadMultipleProperties(property);
      });
    }
  }

  /**
   * Adds a property of any type to OfficeMockObject
   * @param propertyName Property name to the property to be added
   * @param value Value this added property will have
   */
  setMock(propertyName: string, value: unknown) {
    if (!this.properties.has(propertyName)) {
      const officeMockObject = new OfficeMockObject();
      officeMockObject.isMockObject = false;
      this.properties.set(propertyName, officeMockObject);
    }
    this.properties.get(propertyName)?.resetValue(value);
    this[propertyName] = this.properties.get(propertyName)?.value;
  }

  /**
   * Mock replacement for the sync of Office.js API
   */
  sync() {
    this.properties.forEach((property: OfficeMockObject, key: string) => {
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

  private loadMultipleProperties(properties: string) {
    properties
      .replace(/\s/g, "")
      .split(",")
      .forEach((completeProperties: string) => {
        this.loadNavigational(completeProperties);
      });
  }

  private loadNavigational(completePropertyName: string) {
    const properties: Array<string> = completePropertyName.split("/");
    let navigationalOfficeMockObject: OfficeMockObject = this;

    // Iterating through navigational properties
    for (let i = 0; i < properties.length - 1; i++) {
      const property = properties[i];

      const checkingUndefined =
        navigationalOfficeMockObject.properties.get(property);
      if (checkingUndefined) {
        navigationalOfficeMockObject = checkingUndefined;
      } else {
        throw new Error(
          `Navigational property ${property} needs to be present in object model before load is called.`
        );
      }
    }
    const scalarProperty: string = properties[properties.length - 1];
    navigationalOfficeMockObject.loadScalar(scalarProperty);
  }

  private loadScalar(scalarPropertyName: string) {
    if (this.properties.has(scalarPropertyName)) {
      this.properties.get(scalarPropertyName)?.loadCalled();
      this.assignValue(scalarPropertyName);

      this.properties
        .get(scalarPropertyName)
        ?.properties.forEach((property: OfficeMockObject) => {
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

  private properties: Map<string, OfficeMockObject>;
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
