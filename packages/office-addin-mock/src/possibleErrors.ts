export enum PossibleErrors {
  /* eslint-disable no-unused-vars */
  notLoaded = "Error, property was not loaded",
  notSync = "Error, context.sync() was not called",
  /* eslint-enable no-unused-vars */
}

export function isValidError(str: string): boolean {
  let foundError = false;
  Object.values(PossibleErrors).forEach((possibleError: string) => {
    if (str === possibleError) {
      foundError = true;
    }
  });
  return foundError;
}
