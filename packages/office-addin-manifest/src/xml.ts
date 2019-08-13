// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

export type Xml = any;

/**
 * Given an xml element, returns the value of the attribute with the specified name.
 * @param xml Xml object
 * @param name Attribute name
 * @returns The attribute value or undefined
 * @example Given the the following xml, the attribute name "DefaultValue" will return the value "abc".
 *   <First DefaultValue="abc">1</First>
 */
export function getXmlAttributeValue(xml: Xml, name: string): string | undefined {
    try {
      return xml.$[name];
    } catch (err) {
      // reading xml values is resilient to errors but you can uncomment the next line for debugging if attributes are missing
      // console.error(`Unable to get xml attribute value "${name}". ${err}`);
    }
}

/**
 * Given an xml object, returns the first inner element with the specified name, or undefined.
 * @param xml Xml object
 * @param name Element name
 * @returns Xml object or undefined
 * @example Given the the following xml, the name "Second" will return the xml object for <Second>...</Second>.
 *   <Current>
 *     <First>1</First>
 *     <Second>2</Second>
 *   </Current>
 */
export function getXmlElement(xml: Xml, name: string): Xml | undefined {
  try {
    const element = xml[name];

    if (element instanceof Array) {
      return element[0];
    }
  } catch (err) {
    // reading xml values is resilient to errors but you can uncomment the next line for debugging if elements are missing
    // console.error(`Unable to get xml element "${name}". ${err}`);
  }
}

/**
 * Given an xml object, returns the attribute value for the first inner element with the specified name, or undefined.
 * @param xml Xml object
 * @param elementName Element name
 * @param attributeName Attribute name
 * @example Given the the following xml, the element name "First" and attribute name "DefaultValue" will return the value "abc".
 *   <Current>
 *     <First DefaultValue="abc">1</First>
 *   </Current>
 */
export function getXmlElementAttributeValue(xml: Xml, elementName: string, attributeName: string = "DefaultValue"): string | undefined {
  const element: Xml = getXmlElement(xml, elementName);
  if (element) {
    return getXmlAttributeValue(element, attributeName);
  }
}

/**
 * Given an xml object, returns an array with the inner elements with the specified name.
 * @param xml Xml object
 * @param name Element name
 * @returns Array of xml objects;
 * @example Given the the following xml, the name "Item" will return an array with the two items.
 *   <Items>
 *     <Item>1</Item>
 *     <Item>2</Item>
 *   </Items>
 */
export function getXmlElements(xml: Xml, name: string): Xml[] {
  try {
    const elements = xml[name];
    return (elements instanceof Array) ? elements : [];
  } catch (err) {
    return [];
  }
}

/**
 * Given an xml object, for the specified element, returns the values of the inner elements with the specified item element name.
 * @param xml The xml object.
 * @param name The name of the inner xml element.
 * @example Given the the following xml, the container name "Items" and item name "Item" will return ["One", "Two"].
 * If the attributeName is "AnotherValue", then it will return ["First", "Second"].
 *   <Items>
 *     <Item DefaultValue="One" AnotherValue="First">1</Item>
 *     <Item DefaultValue="Two" AnotherValue="Second">2</Item>
 *   </Current>
 */
export function getXmlElementsAttributeValue(xml: Xml, name: string, itemElementName: string, attributeName: string = "DefaultValue"): string[] {
  const values: string[] = [];

  try {
    const xmlElements: Xml[] = xml[name][0][itemElementName];

    xmlElements.forEach((xmlElement: Xml) => {
      const elementValue = getXmlAttributeValue(xmlElement, attributeName);
      if (elementValue !== undefined) {
        values.push(elementValue);
      }
    });
  } catch (err) {
    // do nothing
  }

  return values;
}

/**
 * Given an xml object, for the specified element, returns the values of the inner elements with the specified item element name.
 * @param xml The xml object.
 * @param name The name of the inner xml element.
 * @example Given the the following xml, the container name "Items" and item name "Item" will return ["1", "2"].
 *   <Items>
 *     <Item>1</Item>
 *     <Item>2</Item>
 *   </Current>
 */
export function getXmlElementsValue(xml: Xml, name: string, itemElementName: string): string[] {
  const values: string[] = [];

  getXmlElements(xml, name).forEach((xmlElement) => {
    const elementValue = getXmlElementValue(xmlElement, itemElementName);
    if (elementValue !== undefined) {
      values.push(elementValue);
    }
  });

  return values;
}

/**
 * Returns the value of the first inner xml element with the specified name.
 * @param xml The xml object.
 * @param name The name of the inner xml element.
 * @example Given the the following xml, the name "Second" will return the value "2".
 *   <Current>
 *     <First>1</First>
 *     <Second>2</Second>
 *   </Current>
 */
export function getXmlElementValue(xml: Xml, name: string): string | undefined {
  try {
    const element = xml[name];

    if (element instanceof Array) {
      return element[0];
    }
  } catch (err) {
    // reading xml values is resilient to errors but you can uncomment the next line for debugging if elements are missing
    // console.error(`Unable to get xml element value "${name}". ${err}`);
    }
}

/**
 * Given an xml object, set the attribute value for the specified element name.
 * @param xml Xml object
 * @param elementName Element name
 * @param attributeValue Attribute value
 * @param attributeName Attribute name
 */
export function setXmlElementAttributeValue(xml: Xml, elementName: string, attributeValue: string | undefined, attributeName: string = "DefaultValue") {
  xml[elementName][0].$[attributeName] = attributeValue;
}

/**
 * Given an xml object, set the inner xml element
 * @param xml Xml object
 * @param elementName Element name
 * @param elementValue Element value
 */
export function setXmlElementValue(xml: Xml, elementName: string, elementValue: any) {
  xml[elementName] = elementValue;
}
