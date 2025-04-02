import * as propertiesJson from "./data/properties.json";

export enum PropertyType {
  navigational,
  scalar,
  ambiguous, // Can be scalar or navigational. Depends of the context.
  notProperty,
}

const navigationProperties: Set<string> = new Set<string>(
  propertiesJson.navigational,
);
const scalarProperties: Set<string> = new Set<string>(propertiesJson.scalar);
const ambiguousProperties: Set<string> = new Set<string>(
  propertiesJson.ambiguous,
);

export function getPropertyType(propertyName: string): PropertyType {
  if (navigationProperties.has(propertyName)) {
    return PropertyType.navigational;
  } else if (scalarProperties.has(propertyName)) {
    return PropertyType.scalar;
  } else if (ambiguousProperties.has(propertyName)) {
    return PropertyType.ambiguous;
  } else {
    return PropertyType.notProperty;
  }
}
