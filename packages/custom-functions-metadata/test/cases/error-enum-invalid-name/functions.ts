// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/**
 * Invalid enum name starting with a '_'
 * @customenum
 */
enum _PLANETS {
  /** mercury is the first planet from the sun */
  mercury = "mercuryvalue",
  /** venus is the second planet from the sun */
  venus = "venusvalue",
}

/**
 * Invalid enum name with unsupported characters
 * @customenum {number}
 */
enum NUMBERS$ {
  /** One */
  One = 1,
  /** Two */
  Two = 2,
}

/**
 * Invalid enum name with 129 characters
 * @customenum
 */
enum InvalidEnumName123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234 {
  /** One */
  One = 1,
  /** Two */
  Two = 2,
}