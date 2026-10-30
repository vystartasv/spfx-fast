# Release and publish

## GitHub checks

`.github/workflows/ci.yml` runs `pnpm test` and `pnpm test:package` on pushes and pull requests. Both npm and pnpm work locally:

```sh
pnpm test
pnpm test:package
```

`test:package` packs the library, installs that tarball in a clean consumer, checks its exports and agent docs, generates apps, and builds both SPFx fixtures.

## Publish from GitHub

`.github/workflows/publish.yml` runs when a GitHub Release is published. It verifies the `vX.Y.Z` tag matches `package.json`, reruns tests, validates the package tarball, then publishes that exact tarball to npm.

Set up npm trusted publishing once:

1. Open `@spfx-fast/core` on [npmjs.com](https://www.npmjs.com/package/@spfx-fast/core) → **Settings** → **Trusted Publisher**.
2. Choose **GitHub Actions**. Enter owner `vystartasv`, repository `spfx-fast`, workflow `publish.yml`; leave environment blank.
3. Allow direct `npm publish` for this publisher.
4. Create a GitHub Release whose tag matches the package version, such as `v0.2.9` for `package.json` version `0.2.9`.

The workflow uses GitHub OIDC; it needs no npm token or 2FA code. npm requires CLI 11.5.1+ and Node 22.14+. GitHub-hosted runners and public repositories get provenance automatically. The first trusted publish must succeed within two days of configuring the publisher. See [npm trusted publishing](https://docs.npmjs.com/trusted-publishers/).

Do not publish a GitHub Release until its matching package version is ready. A repeated version cannot be republished.

## Local publish

Use this only if publishing manually. Review the version, README, and packed files, then publish:

```sh
pnpm test
pnpm test:package
pnpm pack --dry-run
pnpm publish --access public
```

The same scripts work with npm (`npm test`, `npm run test:package`, `npm pack --dry-run`, `npm publish --access public`). To publish the exact smoke-tested tarball, set `SPFX_FAST_PACKAGE_ARCHIVE=/tmp/spfx-fast-core-<version>.tgz` during `test:package`, then publish that file.

## Recent releases

- **0.2.9:** refreshes the README gallery with current Fluent UI web parts and cache-busted screenshot assets.
- **0.2.8:** compact mobile-first README gallery with a prominent sample dashboard and expandable screenshots.
- **0.2.7:** refreshed Fluent UI v9 examples and screenshots, improved floor-plan map and mobile layout.
- **0.2.6:** run agent CLI commands through the project package manager, without a global install.
- **0.2.5:** installable Agent Skill, quick routes for common SPFx requests, and focused npm search terms.
- **0.2.4:** cached floor-plan components, resumable upload recovery, list permission helpers, tenant validation.
- **0.2.3:** drive sync dashboard, file/version/sharing actions, calendar/device examples, app-catalog deployment guide.
- **0.2.2:** public GitHub links.
- **0.2.1:** shorter npm quick start.
- **0.2.0:** 100 use cases, 24 form presets, CLI catalog search, and roadmap.

Tenant results and remaining checks: [validation log](TENANT-VALIDATION.md). Local builds cannot verify tenant permissions or live SharePoint behavior.
