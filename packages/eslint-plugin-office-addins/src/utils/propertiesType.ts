import * as propertiesTypeJson from "./data/propertiesType.json";

export enum PropertyType {
  navigational,
  scalar,
  ambiguous, // Can be scalar or navigational. Depends of the context.
  notProperty,
}

const navigationProperties: Set<string> = new Set<string>(
  propertiesTypeJson.navigational
);
const scalarProperties: Set<string> = new Set<string>(
  propertiesTypeJson.scalar
);
const ambiguousProperties: Set<string> = new Set<string>(
  propertiesTypeJson.ambiguous
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
