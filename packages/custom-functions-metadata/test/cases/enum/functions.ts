/**
 * Enum for planets with descriptions and tooltips.
 * @customenum
 * @enum {string}
 */
enum PLANETS {
  /** mercury is the first planet from the sun */
  mercury = "mercuryvalue",
  /** venus is the second planet from the sun */
  venus = "venusvalue",
  /** earth is the third planet from the sun */
  earth = "earthvalue",
  /** mars is the fourth planet from the sun */
  mars = "marsvalue",
  /** jupiter is the fifth planet from the sun */
  jupiter = "jupitervalue",
  /** saturn is the sixth planet from the sun */
  saturn = "saturnvalue",
  /** uranus is the seventh planet from the sun */
  uranus = "uranusvalue",
  /** neptune is the eighth planet from the sun */
  neptune = "neptunevalue",
}

/**
 * Test string enum
 * @customfunction
 * @param first
 * @param second param of enum type planets
 * @returns
 */
export function testStringEnum(first: number, second: PLANETS): any {
  return second;
}

/**
 * Enum for numbers with descriptions and tooltips.
 * @customenum
 * @enum {number}
 */
enum NUMBERS {
  /** One */
  One = 1,
  /** Two */
  Two = 2,
  /** Three */
  Three = 3,
  /** Four */
  Four = 4,
  /** Five */
  Five = 5,
}

/**
 * Test number enum
 * @customfunction
 * @param first
 * @param second param of enum type numbers
 * @returns
 */
export function testNumberEnum(first: number, second: NUMBERS[]): any {
  const sum = second.reduce((acc, num) => acc + num, 0);
  return first + sum;
}
