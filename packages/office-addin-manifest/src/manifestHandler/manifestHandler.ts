export abstract class manifestHandler {
  /* eslint-disable no-unused-vars */
  abstract getAttributeValue(name: string): string | undefined;
  abstract getElement(name: string): any | undefined;
  abstract getElementAttributeValue(elementName: string, attributeName: string): string | undefined;
  abstract getElements(name: string): any[];
  abstract getElementsAttributeValue(name: string, itemElementName: string, attributeName: string): string[];
  abstract getElementsValue(name: string, itemElementName: string): string[];
  abstract getElementValue(name: string): string | undefined;
  abstract setElementAttributeValue(
    elementName: string,
    attributeValue: string | undefined,
    attributeName: string
  ): void;
  abstract setElementValue(elementName: string, elementValue: any): void;
  /* eslint-enable no-unused-vars */
}
