# K2 Microsoft Teams broker

This is a K2 Broker for Microsoft Teams (using TypeScript).

# Features

  - Full object model intellisense for making development easier
  - Sample broker code that accesses jsonplaceholder.
  - Sample unit tests with mocks and code coverage.
  - RollupJS configuration for TypeScript.

## Getting Started

This template requires [Node.js](https://nodejs.org/) v12.14.1+ to run.

Install the dependencies and devDependencies, build the dist/index.js (K2 upload file), or run tests:

```bash
npm install
```

## Run tests
```bash
npm run test
```

## Building your bundled JS
When you're ready to build your broker, run the following command

```bash
npm run build
```

You will find the results in the [dist/index.js](./dist/index.js).

## Creating a service type
Once you have a bundled .js file, upload it to your repository (anonymously
accessible) and register the service type using the system SmartObject located
at System > Management > SmartObjects > SmartObjects > JavaScript Service
Provider and run the Create From URL method.

### License

MIT, found in the [LICENSE](./LICENSE) file.

[www.k2.com](https://www.k2.com)
