const uuid = require('uuid/v1');

export function getXmlAttributeValue(xml: any, name: string): string | undefined {
    try {
      return xml.$[name];
    } catch (err) {
      // console.error(`Unable to get xml attribute value "${name}". ${err}`);
    }
  }

export function getXmlElementAttributeValue(xml: any, elementName: string, attributeName: string = "DefaultValue"): string | undefined {
  const element = getXmlElementValue(xml, elementName);
  if (element) {
    return getXmlAttributeValue(element, attributeName);
  }
}

export function getXmlElementValue(xml: any, name: string): string | undefined {
  try {
    const element = xml[name];

    if (element) {
      return element[0];
    }
  } catch (err) {
      // console.error(`Unable to get xml element value "${name}". ${err}`);
    }
  }

export function setModifiedXmlData(xml: any, guid: string | undefined, displayName: string | undefined) {
  if (guid) {
    if (guid === "random") {
      guid = uuid();
    }
    setXmlElementValue(xml, "Id", guid);
  }

  if (displayName) {
    setElementAttributeValue(xml, "DisplayName", displayName);
  }
}

export function setXmlElementValue(xml: any, elementName: string, input: any) {
  const element = getXmlElementValue(xml, elementName);

  if (element) {
  try {
    xml[elementName] = input;
    } catch (err) {
      console.error(`Unable to write value to xml element: ${err}`);
    }
  return xml;
  }
}

export function setElementAttributeValue(xml: any, elementName: string, input: string | undefined, attributeName: string = "DefaultValue") {
  const element = getXmlElementValue(xml, elementName);

  if (element) {
    try {
      xml[elementName][0].$[attributeName] = input;
    } catch (err) { console.error(`Unable to write value to xml attribute: ${err}`); }
  }
}
