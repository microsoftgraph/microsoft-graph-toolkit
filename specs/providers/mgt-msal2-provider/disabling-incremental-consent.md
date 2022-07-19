# Disabling Incremental Consent

## Status

| Date                | Version | Author           | Status |
| ------------------- | ------- | ---------------- | ------ |
| February 17th, 2022 | v1.0    | Sébastien Levert | Final  |

## Scope

The following providers are affected by this specification

- `mgt-msal2-provider`

## Context

When utilizing the `mgt-msal2-provider` we want to allow developers to turn off the automatic consent prompt. This serves the scenario of LoB applications where the administrator want to control the scopes allowed for the users of the app. This has the impact of never requesting any scopes from the users and creates a better and controlled experience.

## Design

When a developer turns on `isIncrementalConsentDisabled` on `mgt-msal2-provider`, the source of truth of all scopes to be required are the ones assigned to the Provider via its `scopes`. This includes any component that is configured to do incremental consent (`mgt-file`, `mgt-file-list`, etc.) via any Microsoft Graph calls (`graph.files.ts`, etc.). In this case, we should avoid requesting additional permissions and assume the required permissions are already available on the authentication provider scopes. 

## Implementation

We want to think about the future of incremental consent and we need a global way to identify if we want to disable incremental consent. The setting should be global and can be used within any components that has access to the `IProvider` base provider. As there is only a single provider available on a page, we can safely assume this setting is global to the overall application. By doing this, providers that want to implement the ability to disable the incremental consent will be able to set the provider setting and rely on it when calling into Graph via the `GraphHelper`. As we are funneling all calls to add scopes to the `prepScopes` method, we can use this method as the gate keeper to consent.