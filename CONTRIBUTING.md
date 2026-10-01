# Contributing

Thanks for helping improve Office MCP. The project is MIT-licensed and accepts focused bug fixes, tests, documentation improvements, and Office capability additions.

## Before you start

- Search existing issues and pull requests before opening a duplicate.
- For a new tool or breaking API change, open a feature discussion first. Explain the user task, supported Office versions, cross-platform behavior, and why the existing tools cannot express it.
- Do not submit private documents, credentials, customer data, or copyrighted templates without permission. Create synthetic fixtures instead.
- Report security issues privately as described in [SECURITY.md](SECURITY.md).

## Local setup

Use Node.js 20 or newer. Clone the repository, then run:

```sh
npm ci
npm run check
npm run smoke
```

`npm run check` covers TypeScript, portable tests, and the build. The smoke test starts the built stdio server and exercises its MCP tools. On Windows, it also tests rendering when desktop PowerPoint is installed. `npm run test:native` requires installed desktop Microsoft Office on Windows and is not a CI requirement.

Portable tests should not require Office, LibreOffice, credentials, or network access. Add a test for each new tool's successful result and at least one invalid-input or failure case. Native features should have a manual/integration test where feasible, plus a clear capability and limitation entry in `README.md` and `office_capabilities`.

## Pull request checklist

1. Keep the MCP stdio protocol clean: logs go to stderr, not stdout.
2. Validate all file paths against `OFFICE_MCP_ROOTS`; preserve inputs and avoid overwriting by default.
3. Keep native PowerPoint file work hidden. Do not take over a user's active Office session without an explicit opt-in.
4. Run `npm run check` and `npm run smoke`; run `npm run test:native` when changing the COM bridge and Office is available.
5. Update user-facing documentation, capability reporting, and release notes for observable changes.
6. State what was not tested, especially native Office versions and rendered output.

The current root-level dependency overrides and bundled Excel/PowerPoint dependencies are intentional. They keep the published tarball on patched transitive versions; test a freshly installed tarball and run `npm audit --omit=dev` before changing them.
