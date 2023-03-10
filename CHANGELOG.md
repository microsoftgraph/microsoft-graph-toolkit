## v3.0.0-preview.1 (2023-03-02)

### New feature:

- add quick messaging to fluent person-card (#1958) by Nickii Miaro
- update agenda component to the fluent UI spec (#1867) by Musale Martin
- update teams-channel-picker to fluent UI designs (#1805) by Musale Martin
- report custom element name collisions (#2053) by Gavin Barron
- add theme management tools (#2037) by Gavin Barron
- mgt-picker component for generic picking of entities from Microsoft Graph (#1937) by Nickii Miaro
- update File List component to Fluent UI (#1833) by Nickii Miaro
- update TeamsFxProvider.ts for v3.0.0 (#1983) by rentu
- add typing to events for react components (#1981) by Gavin Barron
- add spfx utils for disambiguation (#1914) by Gavin Barron
- add tests and example jest config (#1987) by Gavin Barron
- add support for GCC and other sovereign clouds (#1928) by Musale Martin
- add nodejs 16 support (#1911) by Gavin Barron
- upgrade sample to angular 14 (#1968) by Gavin Barron
- add custom element disambiguation #1852) by Gavin Barron
- update people-picker component to fluentui design (#1801) by Musale Martin
- update the people component to Fluent UI (#1786) by Musale Martin
- update mgt-login to new fluent-ui designs (#1807) by Gavin Barron
- update Person Card to latest Fluent UI (#1797) by Nickii Miaro
- update File component to latest Fluent design (#1802) by Nickii Miaro
- update person component to latest Fluent UI design (#1773) by Nickii Miaro
- updated Graph Client to v3 (#1040) by Nikola Metulev
- person-card fluent controls upgrade (#1253) by Nicolas Vogt
- added option to disable incremental consent (#1316) by amrutha95
- update fluentui registration (#1338) by Beth Pan
- added multi-user cache functionality and enabled multi user login in Msal2Provider (#1299) by amrutha95
- msal2 multi-account UI (#1041) by Nicolas Vogt

### Bugs fixed:
- fix: set the search icon to be on the same level with the input field (#2043) by Musale Martin
- fix: update typescript and ts-node versions for proxy samples (#2020) by Gavin Barron
- fix: correct typing problems in sample vue app (#2021) by Gavin Barron
- fix: restore provided msal public client behavior (#1931) by Gavin Barron
- fix: people picker RTL renders, focus and storybook loading errors (#1864) by Musale Martin
- fix: lock responselike resolutions to v2.0.0 (#1851) by Nickii Miaro

### BREAKING CHANGES:

- In mgt-agenda for eventClick the clicked MicrosoftGraph.Event moves from a property of e.detail to be the value of e.detail
- All events for mgt-task now emit a CustomEvent<ITask>
- Numerous changes to design tokens that may break styling customizations
- @microsoft/microsoft-graph-client now uses v3.0.2, upgraded from v2.2.1, solutions using the graph client from the provider will experience breaking changes.
