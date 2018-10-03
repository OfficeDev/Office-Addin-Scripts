const uuid = require('uuid/v1');

export function xmlAttributeValue(xml: any, name: string): string | undefined {
    try {
      return xml.$[name];
    } catch (err) {
      // console.error(`Unable to get xml attribute value "${name}". ${err}`);
    }
  }
  
export function xmlElementAttributeValue(xml: any, elementName: string, attributeName: string = "DefaultValue"): string | undefined {
    const element = xmlElementValue(xml, elementName);
    if (element) {
      return xmlAttributeValue(element, attributeName);
    }
  }
  
export function xmlElementValue(xml: any, name: string): string | undefined {
  try {
    const element = xml[name];
  
    if (element) {
      return element[0];
    }
  } catch (err) {
      // console.error(`Unable to get xml element value "${name}". ${err}`);
    }
  }

  export function setXmlElementValue(xml: any, element: string, input: any)
{
  try {
    // check to see if element specified is 'Id' and the input is 'random', in which case
    // we randomly generate a guid
    if (element == "Id" && input == 'random') {
      input = uuid();
    }
    xml.OfficeApp[element] = input;
    } catch (err) {
    console.error(`Unable to write value to xml element: ${err}`);
  }
  return xml;
}

export function setElementAttributeValue(xml: any, elementName: string, input: string, attributeName: string = "DefaultValue")
{
  const element = xmlElementValue(xml.OfficeApp, elementName);

  if (element) {
    try{
      xml.OfficeApp[elementName][0].$[attributeName] = input;
    }
    catch (err) {
      console.error(`Unable to write attribute to xml: ${err}`);
    }
  }
}