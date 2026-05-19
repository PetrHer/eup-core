module.exports = {
    testEnvironment: 'jest-environment-jsdom',
    transform: {
        '^.+\\.tsx?$': ['ts-jest', {
            tsconfig: {
                module: 'CommonJS',
                esModuleInterop: true,
            },
            diagnostics: false,
        }],
    },
    moduleNameMapper: {
        '\\.(css|scss|sass|less)$': 'identity-obj-proxy',
        '@microsoft/(.*)': 'identity-obj-proxy',
        '@syncfusion/(.*)': 'identity-obj-proxy',
        '@pnp/(.*)': 'identity-obj-proxy',
    },
    testMatch: ['**/?(*.)+(spec|test).[jt]s?(x)'],
    testPathIgnorePatterns: ['/node_modules/', '<rootDir>/lib/'],
    moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json'],
    collectCoverage: true,
    coveragePathIgnorePatterns: [
        '/node_modules/',
        '\\.tsx$',
    ],
    coverageThreshold: {
        global: {
            lines: 70,
            functions: 70,
        }
    },
};
