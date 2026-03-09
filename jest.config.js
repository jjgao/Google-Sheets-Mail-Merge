module.exports = {
  testEnvironment: 'node',
  setupFiles: ['<rootDir>/jest.setup.js'],
  testMatch: ['**/__tests__/**/*.test.js'],
  clearMocks: true,
  transform: {
    '^.+\\.gs$': '<rootDir>/jest.gastransform.js',
  },
  moduleFileExtensions: ['js', 'json', 'gs'],
  collectCoverageFrom: ['*.gs', '!UI.gs'],
};
