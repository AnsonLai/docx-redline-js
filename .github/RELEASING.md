# Releasing @ansonlai/docx-redline-js via GitHub

This repo is configured so **only GitHub Releases** publish to npm.

## Workflows

- `.github/workflows/ci.yml`
  - Runs on PRs and pushes to `main`.
  - Executes: `npm ci`, `npm run test:isolation`, `npm test`, `npm run build`.
- `.github/workflows/publish.yml`
  - Runs only when a GitHub Release is **published**.
  - Validates release tag matches `package.json` version (`vX.Y.Z`).
  - Executes full verification and then `npm publish --access public --provenance`.

## One-time setup

1. Push this repo to GitHub.
2. Configure npm auth for CI:
   - Preferred: npm Trusted Publishing (GitHub OIDC).
     - npm package settings -> Trusted Publishers -> add this GitHub repo/workflow.
   - Alternative: create `NPM_TOKEN` secret in GitHub with publish rights.
3. (Recommended) Protect `main` branch and require CI checks.

## Release protocol (example: 0.1.1)

1. Update code.
2. Bump version and create tag locally:
   - `npm version patch -m "release: %s"`
3. Push commit and tag:
   - `git push`
   - `git push --tags`
4. In GitHub, create/publish a Release from tag `v0.1.1`.
5. GitHub Actions `Publish` workflow runs and publishes to npm.

## Why this prevents random publishes

- CI never publishes.
- Publish workflow is gated to `release.published` events only.
- Tag/version mismatch fails the publish job.
