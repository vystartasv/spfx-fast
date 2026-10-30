# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Stack

TypeScript toolkit for SPFx 1.24+: framework-independent core, React hooks/components, and authenticated SPFx REST/Graph adapters. Generated starters use Fluent UI React v9 and SharePoint patterns.

## Users

Agents build SPFx apps; developers refine them.

## Product Purpose

Turn short app requests into working SPFx apps. Reuse data access, paging, caching, sync, and UI patterns.

## Positioning

Agent-first SharePoint toolkit with discoverable app generation, local-first data, and SPFx authentication.

## Operating Context

Apps use the signed-in user's SharePoint and Graph access. Agents use the shipped examples, capabilities, schemas, and templates. Each agent environment must load the skill; `npm install` alone does not register it.

## Capabilities and Constraints

- Target SPFx 1.24+; validate each release with its supported React version.
- Keep the data core framework-independent; provide React and SPFx adapters.
- Use SPFx clients for authenticated REST/Graph. The toolkit does not own credentials or permissions.
- Keep fetched data local and reactive. First reads need a network; browsers can clear or evict storage.
- SharePoint controls permissions, writes, and concurrency. Honor throttling and bound agent results.
- Type common APIs; expose raw REST/Graph for gaps. Isolate legacy APIs and label CSOM-only gaps.
- Use Fluent UI/native SharePoint patterns for generated starter UI; allow developers to replace components.
- Optimize agent discovery and time to a working app.
- Make examples copyable and verifiable product features.

## Product Principles

- Agents discover capabilities; they do not guess endpoints.
- Serve warm reads locally; sync incrementally when supported.
- Generate configurable list apps as ordinary, editable code.
- Bound automated work; document permissions, limits, failures, and gaps.
- Measure the full path from request to running authenticated app.
