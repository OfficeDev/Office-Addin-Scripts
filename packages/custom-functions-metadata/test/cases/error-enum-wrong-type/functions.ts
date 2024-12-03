// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Enum for planets with descriptions and tooltips.
 * @customenum {string}
 */
enum PLANETS {
  /** mercury is the first planet from the sun */
  mercury = 0,
  /** venus is the second planet from the sun */
  venus = 1,
}

/**
 * Enum members have different types.
 * @customenum {number}
 */
enum NUMBERS {
  /** One */
  One = 1,
  /** Two */
  Two = "two",
}

/**
 * Enum with a wrong type.
 * @customenum {wrongtype}
 */
enum WRONGTYPE {
  /** One */
  One = 1,
  /** Two */
  Two = 2,
}