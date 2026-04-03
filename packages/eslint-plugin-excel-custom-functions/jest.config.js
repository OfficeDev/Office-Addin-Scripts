module.exports = {
  testEnvironment: "node",
  transform: {
    "^.+\\.tsx?$": ["ts-jest"],
  },
  testRegex: "./tests/.+\\.test\\.ts$",
  moduleFileExtensions: ["ts", "tsx", "js", "jsx", "json"],
};
