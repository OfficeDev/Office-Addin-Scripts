export function logErrorMessage(err: any) {
    console.error(`Error: ${err instanceof Error ? err.message : err}`);
}
