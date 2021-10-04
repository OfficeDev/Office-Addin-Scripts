import { usageDataObject } from "./defaults";

/**
 * Creates an office-js mockable object
 * @param object Object structure to provide initial values for the mock object (Optional)
 */
export class OfficeMockObject {
  constructor(object?: ObjectData) {
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
   * @param methodBody Function to be added to the object. A blank function will be added if no argument is provided (Optional)
   */
  addMockFunction(methodName: string, methodBody?: Function) {
    this[methodName] = methodBody ? methodBody : function () {};
  }

  /**
   * Adds a mock object to OfficeMockObject. The object can be accessed by simply calling `this.objectName`
   * @param objectName Object name of the object to be added
   */
  addMock(objectName: string) {
    const officeMockObject = new OfficeMockObject();
    officeMockObject.isObject = true;
    this.properties.set(objectName, officeMockObject);
    this[objectName] = this.properties.get(objectName);
  }

  /**
   * Mock replacement of the load method in the Office.js API
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
      officeMockObject.isObject = false;
      this.properties.set(propertyName, officeMockObject);
    }
    this.properties.get(propertyName)?.resetValue(value);
    this[propertyName] = this.properties.get(propertyName)?.value;
  }

  /**
   * Mock replacement for the sync method in the Office.js API
   */
  sync() {
    this.properties.forEach((property: OfficeMockObject, key: string) => {
      property.sync();
      this.makePropertyCallable(key);
    });
    if (this.loaded) {
      this.value = this.valueBeforeLoaded;
    }
  }

  private makePropertyCallable(propertyName: string) {
    if (this.properties.get(propertyName)?.isObject) {
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

      const retrievedProperty:
        | OfficeMockObject
        | undefined = navigationalOfficeMockObject.properties.get(property);
      if (retrievedProperty) {
        navigationalOfficeMockObject = retrievedProperty;
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
      this.makePropertyCallable(scalarPropertyName);

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

  private populate(objectData: ObjectData) {
    try {
      Object.keys(objectData).forEach((property: string) => {
        if (typeof objectData[property] === "object") {
          this.addMock(property);
          this[property].populate(objectData[property]);
        } else if (typeof objectData[property] === "function") {
          this.addMockFunction(property, objectData[property]);
        } else {
          this.setMock(property, objectData[property]);
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
  private isObject: boolean | undefined;
  /* eslint-disable-next-line */
  [key: string]: any;
}

// It represents the Object to be used when populating Office JS with data.
class ObjectData {
  /* eslint-disable-next-line */
  [key: string]: any;
}
