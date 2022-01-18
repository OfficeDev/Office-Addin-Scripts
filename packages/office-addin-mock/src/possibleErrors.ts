export enum PossibleErrors {
  /* eslint-disable @typescript-eslint/no-unused-vars */
  notLoaded = "Error, property was not loaded",
  notSync = "Error, context.sync() was not called",
  /* eslint-enable @typescript-eslint/no-unused-vars */
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
