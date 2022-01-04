import { isValidError, PossibleErrors } from "./possibleErrors";

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
   * @deprecated Add the function to the JSON in the constructor instead
   */
  addMockFunction(methodName: string, methodBody?: Function) {
    this[methodName] = methodBody ? methodBody : function () {};
  }

  /**
   * addMock(name) will add a property named “name”, with a new OfficeMockObject as its value, to the object
   * @param objectName Object name of the object to be added
   * @deprecated Add the object to the JSON in the constructor instead
   */
  addMock(objectName: string) {
    if (this[objectName] !== undefined) {
      throw new Error("Mock object already exists");
    }

    const officeMockObject = new OfficeMockObject();
    officeMockObject.isObject = true;
    this.properties.set(objectName, officeMockObject);
    this[objectName] = this.properties.get(objectName);
  }

  /**
   * Mock replacement of the load method in the Office.js API
   * @param propertyArgument Argument of the load call. Will load any properties in the argument
   */
  load(propertyArgument: string | string[] | ObjectData) {
    let properties: string[] = [];

    if (typeof propertyArgument === "string") {
      properties = Array(propertyArgument);
    } else if (Array.isArray(propertyArgument)) {
      properties = propertyArgument;
    } else {
      properties = this.parseObjectPropertyIntoArray(propertyArgument);
    }

    properties.forEach((property: string) => {
      this.loadMultipleProperties(property);
    });
  }

  /**
   * Adds a property of any type to OfficeMockObject
   * @param propertyName Property name to the property to be added
   * @param value Value this added property will have
   * @deprecated Add the property to the JSON in the constructor instead
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
  async sync() {
    this.properties.forEach(async (property: OfficeMockObject, key: string) => {
      await property.sync();
      this.updatePropertyCall(key);
    });
    if (this.loaded) {
      this.value = this.valueBeforeLoaded;
    }
  }

  private loadAllProperties() {
    this.properties.forEach((property, propertyName: string) => {
      property.loadCalled();
      this.updatePropertyCall(propertyName);
    });
  }

  private loadCalled() {
    if (!this.loaded) {
      this.loaded = true;
      this.value = PossibleErrors.notSync;
    }
  }

  private loadMultipleProperties(properties: string) {
    if (properties === "*") {
      this.loadAllProperties();
    } else {
      properties
        .replace(/\s/g, "")
        .split(",")
        .forEach((completeProperties: string) => {
          this.loadNavigational(completeProperties);
        });
    }
  }

  private loadNavigational(completePropertyName: string) {
    const properties: Array<string> = completePropertyName.split("/");
    let navigationalOfficeMockObject: OfficeMockObject = this;

    // Iterating through navigational properties
    for (let i = 0; i < properties.length - 1; i++) {
      const property = properties[i];

      const retrievedProperty: OfficeMockObject | undefined =
        navigationalOfficeMockObject.properties.get(property);
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
      this.updatePropertyCall(scalarPropertyName);

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

  private parseObjectPropertyIntoArray(objectData: ObjectData): string[] {
    let composedProperties: string[] = [];

    Object.keys(objectData).forEach((propertyName: string) => {
      const property: OfficeMockObject | undefined =
        this.properties.get(propertyName);

      if (property) {
        const propertyValue: ObjectData = objectData[propertyName];
        if (property.isObject) {
          const composedProperty: string[] =
            property.parseObjectPropertyIntoArray(propertyValue);
          if (composedProperty.length !== 0) {
            composedProperty.forEach((prop: string) => {
              composedProperties = composedProperties.concat(
                propertyName + "/" + prop
              );
            });
          }
        } else if (propertyValue) {
          composedProperties = composedProperties.concat(propertyName);
        }
      } else {
        throw new Error(
          `Property ${propertyName} needs to be present in object model before load is called.`
        );
      }
    });

    return composedProperties;
  }

  private populate(objectData: ObjectData) {
    Object.keys(objectData).forEach((propertyName: string) => {
      const property = objectData[propertyName];
      const dataType: string = typeof property;

      if (dataType === "object" && !Array.isArray(property)) {
        this.addMock(propertyName);
        this[propertyName].populate(property);
      } else {
        this.setValue(propertyName, property);
      }
    });
  }

  private resetValue(value: unknown) {
    this.value = PossibleErrors.notLoaded;
    this.valueBeforeLoaded = value;
    this.loaded = false;
  }

  /**
   * Sets a property of any type or function to the object
   * @param propertyName Property name to the property to be added
   * @param value Value this added property will have
   */
  private setValue(propertyName: string, value: any) {
    if (typeof value === "function") {
      this[propertyName] = value;
    } else {
      if (!this.properties.has(propertyName)) {
        const officeMockObject = new OfficeMockObject();
        officeMockObject.isObject = false;
        this.properties.set(propertyName, officeMockObject);
      }
      this.properties.get(propertyName)?.resetValue(value);
      this[propertyName] = this.properties.get(propertyName)?.value;
    }
  }

  private updatePropertyCall(propertyName: string) {
    if (this.properties.get(propertyName)?.isObject) {
      this[propertyName] = this.properties.get(propertyName);
    } else if (isValidError(this[propertyName])) {
      // It is a known error
      this[propertyName] = this.properties.get(propertyName)?.value;
    }
  }

  private properties: Map<string, OfficeMockObject>;
  private loaded: boolean;
  private value: unknown;
  private valueBeforeLoaded: unknown;
  private isObject: boolean | undefined;
  /* eslint-disable-next-line */
  [key: string]: any;
}

// Represents the Object to be used when populating Office JS with data.
class ObjectData {
  /* eslint-disable-next-line */
  [key: string]: any;
}
