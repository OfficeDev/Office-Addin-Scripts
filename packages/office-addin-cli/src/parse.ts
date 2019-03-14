export function parseNumber(optionValue: any, errorMessage: string = "The value should be a number."): number | undefined {
    switch (typeof(optionValue)) {
        case "number": {
            return optionValue;
        }
        case "string": {
            let result;

            try {
                result = parseFloat(optionValue);
            } catch (err) {
                throw new Error(errorMessage);
            }

            if (Number.isNaN(result)) {
                throw new Error(errorMessage);
            }

            return result;
        }
        case "undefined": {
            return undefined;
        }
        default: {
            throw new Error(errorMessage);
        }
    }
}
