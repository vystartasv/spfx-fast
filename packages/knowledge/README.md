# @spfx-fast/knowledge

The shareable SPFx specialist skill and field debugging playbook. This keeps SPFx as one skill, with focused references loaded only when the task needs them. It is independent of the runtime and Playwright packages.

```sh
npx --yes @spfx-fast/knowledge --help
npx --yes @spfx-fast/knowledge learn init
npx --yes @spfx-fast/knowledge learn import-playwright test-results/spfx-fast-playwright/learning-draft.json
npx --yes @spfx-fast/knowledge learn validate test-results/spfx-fast-knowledge/learning-draft.json
npx --yes @spfx-fast/knowledge learn patch test-results/spfx-fast-knowledge/learning-draft.json
```

`learn patch` writes a reviewable Git patch that adds one sanitized JSON finding under `knowledge/findings/`. It does not upload, open a PR, include identity, or include evidence attachments. Apply the patch in a checkout of `spfx-fast`, review it, then submit it through the normal pull request flow. GitHub may associate the PR with the submitter's account.

Use `learn import-playwright` to convert an existing Playwright learning draft into the contribution format. Review the mapped diagnostic, cause, and fix fields before validating; conversion is mechanical and does not infer correctness.

Install `skill/SKILL.md` into an agent's skill directory, or point the agent configuration at that file. The focused debugging reference is in `docs/DEBUGGING.md`.
