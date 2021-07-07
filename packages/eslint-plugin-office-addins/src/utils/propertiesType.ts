const propertiesTypeJson = require("./data/propertiesType.json");

export enum PropertyType {
  navigational,
  scalar,
  undefined, // Can be scalar or navigational. Depends of the context.
  notProperty,
}

const navigationProperties: Set<string> = new Set<string>(
  propertiesTypeJson.navigational
);
const scalarProperties: Set<string> = new Set<string>(
  propertiesTypeJson.scalar
);
const undefinedProperties: Set<string> = new Set<string>(
  propertiesTypeJson.undefined
);

export function getPropertyType(propertyName: string): PropertyType {
  if (navigationProperties.has(propertyName)) {
    return PropertyType.navigational;
  } else if (scalarProperties.has(propertyName)) {
    return PropertyType.scalar;
  } else if (undefinedProperties.has(propertyName)) {
    return PropertyType.undefined;
  } else {
    return PropertyType.notProperty;
  }
}
