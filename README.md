# jtfg-eup-core

Starter scaffold for a Node.js module.

## Usage

1. Copy your module files into `src/`.
2. Update `package.json` metadata (`description`, `author`, version, etc.).
3. Run `npm install` if you add dependencies.

## Development

- Entry point: `src/index.ts`
- Build output: `dist/`

## Publish To GitHub Packages

1. Ensure the package is scoped: `@petrher/jtfg-eup-core`.
2. Set a GitHub token in your shell:
	- PowerShell: `$env:NODE_AUTH_TOKEN="<YOUR_GITHUB_TOKEN>"`
	- bash: `export NODE_AUTH_TOKEN=<YOUR_GITHUB_TOKEN>`
3. Bump the version in `package.json`.
4. Publish:
	- `npm publish`

Your `.npmrc` already points the `@petrher` scope to GitHub Packages.
