# Contributing

Contributions should save another SPFx developer or agent from repeating discovery. Prefer a small verified fix, recipe, or debugging note with its applicability and limits. Keep tenant and user data private.

## Code and documentation changes

1. Work on a focused branch and keep the change limited to the owning package or guide.
2. Never commit tenant URLs or IDs, user data, credentials, browser profiles, screenshots, or traces.
3. Run the relevant repository checks (`pnpm test` and `pnpm test:package` for package changes).
4. Commit the change with a clear message and open a pull request. Explain the problem, the fix, and any validation limits.

## Knowledge contributions

1. Install or run the `@spfx-fast/knowledge` package and create a local draft:

   ```sh
   npx --yes @spfx-fast/knowledge learn init
   ```

   To reuse a draft from `@spfx-fast/playwright`, convert it with `learn import-playwright <file>` and inspect the mapped fields.

2. Fill in symptom, diagnostic evidence, root cause, fix, reproduction, applicability, limitations, and validation level. Use `tenant-validated` only for behavior actually checked against SharePoint; include the date and sample size where applicable.
3. Validate and create a contribution patch:

   ```sh
   npx --yes @spfx-fast/knowledge learn validate test-results/spfx-fast-knowledge/learning-draft.json
   npx --yes @spfx-fast/knowledge learn patch test-results/spfx-fast-knowledge/learning-draft.json
   ```

4. Inspect the patch and all source evidence yourself. The validator catches common identifiers but cannot guarantee anonymity. Keep evidence local; never include tenant URLs/IDs, identities, list content, tokens, browser profiles, screenshots, or traces.
5. Apply the generated patch in a checkout of this repository, confirm the finding adds reusable value, and commit it as a small, focused change (replace the example patch filename with the generated one):

   ```sh
   git apply sharepoint-rest-path.patch
   git add knowledge/findings/
   git commit -m "Add SharePoint finding: REST path"
   ```

   Then open a pull request. Committing useful discoveries is encouraged; the commit and PR may reveal the submitter's GitHub identity even though the finding contains no author field.
6. Maintainers should check accuracy, scope, privacy, attribution, duplication, and whether the finding belongs in a durable guide, example, or recipe. Keep unverified claims labeled; reject or revise unsupported universal wording.

The package performs no network request or upload. Do not paste private customer evidence into issues or pull requests.
