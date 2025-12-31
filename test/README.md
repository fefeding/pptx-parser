# Test Directory

This directory contains test files for the ppt-parser library.

## Test Files

- **utils.test.ts** - Tests for utility functions (ID generation, unit conversion, XML parsing)
- **parser.test.ts** - Tests for PPTX parsing functionality
- **serializer.test.ts** - Tests for PPTX serialization functionality
- **integration.test.ts** - Integration tests for end-to-end workflows
- **types.test.ts** - Tests for TypeScript type definitions

## Running Tests

```bash
# Run tests in watch mode
npm test

# Run tests once
npm run test:run

# Run tests with coverage
npm run test:coverage
```

## Test Structure

All tests use Vitest as the test runner with jsdom environment to simulate browser APIs (like DOMParser needed for XML parsing).
