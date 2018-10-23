export function getXmlAttributeValue(xml: any, name: string): string | undefined {
    try {
      return xml.$[name];
    } catch (err) {
      // reading xml values is resilient to errors but you can uncomment the next line for debugging if attributes are missing
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
    // reading xml values is resilient to errors but you can uncomment the next line for debugging if elements are missing
    // console.error(`Unable to get xml element value "${name}". ${err}`);
    }
}

export function setXmlElementAttributeValue(xml: any, elementName: string, input: string | undefined, attributeName: string = "DefaultValue") {
  xml[elementName][0].$[attributeName] = input;
}

export function setXmlElementValue(xml: any, elementName: string, input: any) {
  xml[elementName] = input;
}
