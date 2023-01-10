import { OfficeApp } from "office-addin-manifest";
import { getHostType } from "./host";
import { isValidError, PossibleErrors } from "./possibleErrors";
import { ObjectData } from "./objectData";

/**
 * Creates an office-js mockable object
 * @param object Object structure to provide initial values for the mock object (Optional)
 * @param host Host tested by the object (Optional)
 */
export class OfficeMockObject {
  constructor(object?: ObjectData, host?: OfficeApp | undefined) {
    this._properties = new Map<string, OfficeMockObject>();
    this._loaded = false;
    if (host) {
      this._host = host;
    } else {
      this._host = getHostType(object);
    }
    this.resetValue(undefined);
    if (object) {
      this.populate(object);
    }
  }

  /**
   * Mock replacement of the load method in the Office.js API
   * @param propertyArgument Argument of the load call. Will load any properties in the argument
   */
  load(propertyArgument: string | string[] | ObjectData) {
    if (this._host === OfficeApp.Outlook) {
      return;
    }
    let properties: string[] = [];

    if (propertyArgument === undefined) { 
      // an empty load call mean load all properties
      properties = ["*"];
    } else if (typeof propertyArgument === "string") {
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
   * Mock replacement for the sync method in the Office.js API
   */
  async sync() {
    this._properties.forEach(
      async (property: OfficeMockObject, key: string) => {
        await property.sync();
        this.updatePropertyCall(key);
      }
    );
    if (this._loaded) {
      this._value = this._valueBeforeLoaded;
    }
  }

  /**
   * addMock(name) will add a property named “name”, with a new OfficeMockObject as its value, to the object
   * @param objectName Object name of the object to be added
   */
  private addMock(objectName: string) {
    if (this[objectName] !== undefined) {
      throw new Error("Mock object already exists");
    }

    const officeMockObject = new OfficeMockObject(undefined, this._host);
    officeMockObject._isObject = true;
    this._properties.set(objectName, officeMockObject);
    this[objectName] = this._properties.get(objectName);
  }

  private loadAllProperties() {
    this._properties.forEach((property, propertyName: string) => {
      property.loadCalled();
      this.updatePropertyCall(propertyName);
    });
  }

  private loadCalled() {
    if (!this._loaded) {
      this._loaded = true;
      this._value = PossibleErrors.notSync;
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
        navigationalOfficeMockObject._properties.get(property);
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
    if (this._properties.has(scalarPropertyName)) {
      this._properties.get(scalarPropertyName)?.loadCalled();
      this.updatePropertyCall(scalarPropertyName);

      this._properties
        .get(scalarPropertyName)
        ?._properties.forEach((property: OfficeMockObject) => {
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
        this._properties.get(propertyName);

      if (property) {
        const propertyValue: ObjectData = objectData[propertyName];
        if (property._isObject) {
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

      if (
        dataType === "object" &&
        !Array.isArray(property) &&
        !(property instanceof Date)
      ) {
        this.addMock(propertyName);
        this[propertyName].populate(property);
      } else {
        this.setValue(propertyName, property);
      }
    });
  }

  private resetValue(value: unknown) {
    if (this._host === OfficeApp.Outlook) {
      this._value = value;
    } else {
      this._value = PossibleErrors.notLoaded;
      this._valueBeforeLoaded = value;
      this._loaded = false;
    }
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
      if (!this._properties.has(propertyName)) {
        const officeMockObject = new OfficeMockObject(undefined, this._host);
        officeMockObject._isObject = false;
        this._properties.set(propertyName, officeMockObject);
      }
      this._properties.get(propertyName)?.resetValue(value);
      this[propertyName] = this._properties.get(propertyName)?._value;
    }
  }

  private updatePropertyCall(propertyName: string) {
    if (this._properties.get(propertyName)?._isObject) {
      this[propertyName] = this._properties.get(propertyName);
    } else if (isValidError(this[propertyName])) {
      // It is a known error
      this[propertyName] = this._properties.get(propertyName)?._value;
    }
  }

  private _properties: Map<string, OfficeMockObject>;
  private _loaded: boolean;
  private _value: unknown;
  private _valueBeforeLoaded: unknown;
  private _isObject: boolean | undefined;
  private _host: OfficeApp | undefined;
  /* eslint-disable-next-line */
  [key: string]: any;
}
