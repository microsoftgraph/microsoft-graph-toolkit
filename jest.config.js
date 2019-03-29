module.exports = {
  collectCoverage: true,
  coverageReporters: [
    "cobertura",
    "html"
  ],
  reporters: ["jest-junit"],
  moduleFileExtensions: [
    "js",
    "json",
    "jsx",
    "ts",
    "tsx",
    "node"
  ],
  testPathIgnorePatterns:
  [
    "/samples/"
  ],
  testEnvironment: "node",
  testRegex: "(/tests/.*|(\\.|/)(test|spec))\\.tsx?$",
  transform: {
    "^.+\\.tsx?$": "ts-jest"
  },
};
