# Changelog

All notable changes to this project will be documented in this file. See [standard-version](https://github.com/conventional-changelog/standard-version) for commit guidelines.

## [4.0.0](https://github.com/microsoftgraph/microsoft-graph-toolkit/compare/v3.0.0...v4.0.0) (2023-07-18)


### ⚠ BREAKING CHANGES

* removes use of user.read.all and group.read.all scopes for team/channel reading
* all events for mgt-task now emit a CustomEvent<ITask>

### Features

* add custom CSS properties for the people picker flyout text ([#2162](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2162)) ([9b600de](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/9b600de1e7285d933e6e4cf642bbda7b6afddbba))
* add custom focus ring color ([#2334](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2334)) ([9df3f19](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/9df3f1953797cccc0ceb65341d59ca0aabfc7d18))
* add custom focus ring color ([#2334](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2334)) ([9f8953a](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/9f8953a4fdd9fafb19551d487197a2bb88d6738e))
* add dedicated icon for business phone ([#1988](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1988)) ([92206b9](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/92206b90e0d556af27a1138c6b2a02804b8bfb9c))
* add new CSS properties for the people picker ([#1926](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1926)) ([041c71e](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/041c71e6cc43880f4e29ce7bdf7f3defef9ccf4d)), closes [#1888](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1888)
* add quick messaging to fluent person-card ([#1958](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1958)) ([94e390b](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/94e390be65d3d2fe3febe913daf6a8aae6eb5c7b))
* add spfx utils for disambiguation ([#1914](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1914)) ([6317ccf](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/6317ccf63a7de136c138d47d8930eeff46b0cf70))
* add support for GCC and other sovereign clouds ([#1928](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1928)) ([1edf635](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/1edf6351d055866caae09e6dcceaa660cc09382c))
* add tests and example jest config ([#1987](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1987)) ([9e00738](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/9e00738e94ca98a44e4eed1cd8d4ed3390d82cfe))
* add theme management tools ([#2037](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2037)) ([9961926](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/99619268bbe016856c86c82d4cb584f80b0a41ae))
* add token overrides to theme switching ([#2328](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2328)) ([c49f70c](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c49f70ca84aac0130601df133bb7f77e09a7c023))
* added mgt-taxonomy-picker component ([#2172](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2172)) ([00f6565](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/00f65655e5cd30b228af5e4f35004362c5451882))
* added spec for mgt-taxonomy-picker ([#2156](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2156)) ([670a3cb](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/670a3cba1dd4db0d2ae16a2399492180df0c8205))
* adds auth page for storybook use ([cfeceeb](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/cfeceeb2a18cea9e26d2fc4bd5f58c44131f7b70))
* allow programmatically theming a component without the theme-toggle component ([#2199](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2199)) ([40b87f9](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/40b87f93075607e64e0dce5ece8484ada4243f31))
* allow TeamsfxProvider to be constructed with TeamssUerCredential ([#1954](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1954)) ([22b204a](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/22b204a0785298ad167df49bfe7310df371937a0))
* cache the endpoint url for the mock-provider in session storage ([#1979](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1979)) ([b2cbbf2](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/b2cbbf2655ebe1d3ab41f0f3554b47fe820d7c75))
* mark Teams, TeamsMsal2 and Msal providers as deprecated in v2.x.x ([#2232](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2232)) ([52efa2e](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/52efa2e73337676f3b329c237645af13cd9089bb))
* mark Teams, TeamsMsal2 and Msal providers as deprecated in v2.x.x ([#2232](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2232)) ([7b701d0](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/7b701d0120f7ee376bc930b9f3b50b89a4a9e75e))
* mgt-loader version warning ([#2059](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2059)) ([e1b352e](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/e1b352e49bcf63f99f872cd2113855e81f418aaf))
* mgt-picker component for generic picking of entities from Microsoft Graph ([#1937](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1937)) ([11e6807](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/11e6807259b623c7042236697bc3653f5d83af0e))
* mgt-picker selected-value attribute ([#2363](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2363)) ([abd00b4](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/abd00b4a1c59cc102658cf1acb14930d35731ecc))
* migrate to eslint ([#2125](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2125)) ([2970a44](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/2970a446109b79f879c53f8f646a9ac304488ec4))
* move completed items to bottom of the list in mgt-todo ([#2215](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2215)) ([1c9face](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/1c9facefa15720a3cc73fe2482cd7c9fee9c296e))
* nodejs 16 support ([#1911](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1911)) ([a573bdf](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/a573bdf940fc154e6cbdc7e7b392298e05f14ab8))
* optionally pass the group-id value from tasks to people picker ([#2200](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2200)) ([711a233](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/711a23381fa1d2a99fa7e944a580081e3bfe13f3))
* preview component support ([#2356](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2356)) ([38a13e1](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/38a13e106fbb79f5df75fea631a693d19ce31477))
* provide dynamic element name ([91d5fa4](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/91d5fa49a25dd6edca6bf55219664472db7332aa))
* remove legacy mgt-dark and mgt-light theme and stories ([#2284](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2284)) ([4065d2f](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/4065d2f7e3bffd8a364a54c3f6534196827831d1))
* remove the Teams, TeamsMsal2 and the Msal providers in v3.x.x ([#2231](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2231)) ([c007abc](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c007abc82d351bf71ea23e50dd7b1c43e532e21c))
* report custom element name collisions ([#2053](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2053)) ([66c2e6c](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/66c2e6c9362ecc64e0841afa68a827854e7e77c4))
* Storybook authentication ([#2048](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2048)) ([9af2bf5](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/9af2bf5f04b1366dc385850f8dd630c647f5b93c))
* Support for new component mgt-search-results ([#2047](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2047)) ([87e6c2f](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/87e6c2f35a9916d26803f3772adc9647670c26fb))
* typed events ([#1981](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1981)) ([bbd5da4](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/bbd5da4896c522fd916aa90b7006bfd89003a926))
* Update agenda component to the fluent UI spec ([#1867](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1867)) ([2771544](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/27715440d8c0e67ed84875c7b4ae7d2ac0ebdb97))
* update File component to latest Fluent design ([#1802](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1802)) ([f074652](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f0746522d4d54521539071a3b7cfd1e232a371ab))
* update File List component to Fluent UI ([#1833](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1833)) ([21e4246](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/21e4246feb1260087b3366fe885868316e2e4f77))
* update mgt-login to new fluent-ui designs ([#1807](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1807)) ([710d6b9](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/710d6b9b0ac200724b78748207f68f0090b05f1d))
* Update people-picker component to fluentui design ([#1801](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1801)) ([7c782ff](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/7c782ffd8554c89a0c4b84ae497c0dd060b5cd49))
* Update Person Card to latest Fluent UI ([#1797](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1797)) ([069309d](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/069309d4b892075c4f4187f3baea45a30aa3c44b))
* Update teams-channel-picker to fluent UI designs ([#1805](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1805)) ([b2e5d8b](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/b2e5d8bce2b3c3a0d6cb4ef39f05a24c06ab69a5))
* update TeamsFxProvider.ts for v3.0.0 ([#1983](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1983)) ([d7247a0](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/d7247a036d6bb0f16ca0e9b6bb8c364e02fc0cf3))
* Update the people component to Fluent UI ([#1786](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1786)) ([8c99782](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/8c99782615c624971bfb11467601cd75b5e156f9))
* update Todo component to new Fluent designs ([#1967](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1967)) ([c7bf047](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c7bf047e66df6472fe63299fae5c381047466f90))
* use fluent UI to theme the tasks component ([#2150](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2150)) ([3677936](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/3677936e82c44cb59b7846fe2902da63b12ddea4))
* use fluentui to theme the person component. ([#2072](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2072)) ([fc4f1e0](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/fc4f1e0aafbb9b8ccdd198593190b2b5dc2792f3))
* use fluentui tokens for theming file and file-list ([#2044](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2044)) ([89b79e9](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/89b79e95035203368cbe427216ff1463261feb74))


### Bug Fixes

* **a11y:** Add keyboard navigation to mgt-teams-channel-picker ([#2415](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2415)) ([5184217](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/51842179abc26e95abb5c8f2e92735b8eb88257e))
* **a11y:** unset custom color of storybook left chevrons ([#2595](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2595)) ([764bf12](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/764bf12429e0d69f7c4c0cfc305bfe7dbed1421f))
* **ac11y:** set aria-expanded when open/closed ([#2405](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2405)) ([d084665](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/d08466597af59cf896808144ac1cd3a6aca60b11))
* add `InAConferenceCall` activity when availability is `Busy` ([#2585](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2585)) ([bd17195](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/bd17195d22a90ab8a820528c91d8c7d6a89191ca))
* add a title tag to be announced for location svg ([#2285](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2285)) ([8c601c9](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/8c601c9cf949d6cc7f2396a017c1c338f2055694))
* add aria-label text  announce cancelling adding a new task ([#2359](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2359)) ([ffaa0e7](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/ffaa0e7066cb0b8dad0873f7e02189091baf0494))
* add class to people-picker styles story to enable custom css ([#2605](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2605)) ([dcec953](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/dcec953ecdc2e721a280a3a11aca5be1596311dc))
* add combobox attributes ([#2538](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2538)) ([b89f045](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/b89f0451f33bd62bd40e11f7e924ae1f54ce5473))
* add font family to tasks ([#2603](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2603)) ([e380b4a](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/e380b4ae98095069a260f459ec6fd6d156d97762))
* add login custom styles, removes style not in use ([#2587](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2587)) ([7ba98e4](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/7ba98e4359854adb2658207be8c8e64bbeeda730))
* add top to the caching key to allow varying values with different page sizes ([#2406](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2406)) ([00b68e6](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/00b68e6ba986b86021430bf49307111221c6f73d))
* add user name to aria-label for people-picker remove button ([#1912](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1912)) ([d93bde7](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/d93bde79feec89a2057b481a832b9aa1f3f4662a))
* Adding taxonomy-picker as an exported React component and updated used types ([4c06bd2](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/4c06bd26d67c1967a53c4edd4465b125314c7d9e))
* adds customHosts support for non-graph domain requests ([#2592](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2592)) ([1f97215](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/1f97215f03324ebbc255ef57fe7658c6978ff0dc))
* Allow display and searching of people supplied in attributes ([#1839](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1839)) ([00865f3](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/00865f30255e7b86c2df369e98bf9059c4601beb))
* announce more options button on narrator in mgt-tasks ([#2399](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2399)) ([11ef91f](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/11ef91fc3e1be8b398aa461cde32055902161c5d))
* announce suggestion list changes ([#2148](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2148)) ([59301ab](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/59301abd665360aa75d1a3f9f420cacc85336a28))
* announce teams channel results when you type ([#2561](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2561)) ([5260ce0](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/5260ce05c125f59ee56a3cbac7eba18e10de523c))
* announce the label of the people-picker text field ([#2398](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2398)) ([f6ba11f](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f6ba11f0a5d5444c3db2b7125e7db0b494c64df9))
* announce the name of the selected user to remove in narrator ([#2360](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2360)) ([a6af856](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/a6af856ff175d2b013297950c6c0e8f63f29bad5))
* apply theme color to calendar icon of date input ([#2312](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2312)) ([68a05bb](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/68a05bb8470c0021e17baf091ae267d28306d43a))
* aspnet proxy provider sample ([#2594](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2594)) ([362339a](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/362339a2a80be148c1e20602f3ae639682b43137))
* caching story ([#2516](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2516)) ([2bd21d1](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/2bd21d1d7e99b0e7c49f95b2f4c5117d8d5d2be1))
* change override design token logic ([#2384](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2384)) ([ed6c9b3](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/ed6c9b310a44a541ed784ebafb8429a9bfd9b004))
* change sppkg details for mgt-spfx ([c46e418](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c46e418cdc29573d564e4c5588118b4266f1422f))
* check pageSize when filtering by file extensions ([#1965](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1965)) ([ab34e64](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/ab34e64954957950e55823c0f9d9886ee96bd7cb))
* correct compilation errors for CI build ([#1956](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1956)) ([4ac80bb](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/4ac80bba18ac805a972f9765d24f9978d70e88c1))
* correct css overrides for contrast issues ([#1814](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1814)) ([08ab575](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/08ab575f764110008cd773880237fd6f2212068b))
* correct sppkg upload script ([#2552](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2552)) ([8b20d84](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/8b20d8489c0043060714a1ca1401d0c7433a7c05))
* correct typing problems in sample vue app ([#2021](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2021)) ([1bc5f03](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/1bc5f0356cfb03448d63ca60bddd73454ff73b77))
* correctly render date times based on user time and preferred timezone ([#1515](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1515)) ([728e578](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/728e5780d7ad0712169427313b2cf959b408e60b))
* disable autocomplete ([#2481](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2481)) ([c9d2195](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c9d21956f81c4d1e7fdc795b6e6c79fb2f3a2a3b))
* disambiguated tagname and query selectors ([#2475](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2475)) ([f9f99e6](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f9f99e69779f25530efffa94b0c6936944d12abb))
* editor tabs keyboard navigation ([#2371](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2371)) ([669b1ea](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/669b1ea8f1046837eab8d5f87d595ca738f57f67))
* email resolution on paste for people picker ([#1791](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1791)) ([964204d](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/964204d08b3f4aa77e6b7b79b989395fca5c8550))
* enable keyboard navigation in the picker ([#2324](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2324)) ([622f000](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/622f0007e9c76422e454edd11d276138d423fd2c))
* enable text spacings on login, channel-picker and teams-channel-picker ([#2413](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2413)) ([08819b1](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/08819b1c4a9388e791aed7cb02c338fd898179d8))
* ensure login pop-up renders inside window ([d084665](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/d08466597af59cf896808144ac1cd3a6aca60b11))
* ensure todo tasks are rendered in mgt-tasks ([#2480](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2480)) ([46afd78](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/46afd781f1b673945c90f5993e6a904641f346a0))
* execute beta queries and eliminate re-renders ([#2391](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2391)) ([70bef48](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/70bef481aecb1155a960b386d3dc87d0c2f74ede))
* files compact view in person card ([#2597](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2597)) ([6985717](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/69857170d87c37c1ce0c0f51baff34922f168e97))
* fix dynamic group id people picker story ([#2023](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2023)) ([666e7a4](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/666e7a4d9cfba25fa9798db13f86fece4a7fde9d))
* fix person-card to use fluent-card ([#2487](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2487)) ([6d3254d](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/6d3254d6c179278ab8a348063c00d1f40201dfe9))
* getGroupImage Endpoint Fix ([#1883](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1883)) ([cd06fe2](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/cd06fe204b6e7532d510359bae05005779534fc9))
* hide storybook canvas button ([#2145](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2145)) ([617989e](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/617989e84563ee5a3ae964927ef3204946229d03)), closes [#1642](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1642)
* include ImageNotFound error code to avoid unnecessary refetching ([#1854](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1854)) ([b8d6891](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/b8d6891e7a02ee63c41a5cf334a935002bbbef34))
* include mgt-mock-provider as dependency to mgt ([#2336](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2336)) ([b166aaf](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/b166aaf80d8c8f737b1e94b3026220d052d027c5))
* keyboard navigation of login account list ([#2289](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2289)) ([af34d15](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/af34d15385db8cf229c7e056af29c6c8380ecc61))
* local and calendar time alignment ([#1903](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1903)) ([41de8c4](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/41de8c49b77390a714eea9248a5843edf138e8e0))
* make people-picker aria-label work in firefox ([#1963](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1963)) ([ae3c0b1](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/ae3c0b142eb4ff15b513b0303a37d60c49ee6dc4))
* make spfx version script work for rc version ([#2396](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2396)) ([2d953fa](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/2d953faa7ea81b90c89a50eb9c1c402d5359526c))
* make tasks header navigable with the keyboard. ([#2313](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2313)) ([4747189](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/4747189c6d66a2a856274b931dac43e86002ff52))
* markdown table names ([#2473](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2473)) ([511e05b](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/511e05b932bb117ff602ba3a880ebdf05ea73b03))
* max picker list height ([#2431](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2431)) ([7a22138](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/7a22138cdd308938c771c3e9523dfeafbfed6d85))
* mgt-person narration ([#2493](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2493)) ([c14af08](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c14af0808c3b828a1ba5982e8c73544a2997bb6d))
* navigate mgt-people using left/right arrow keys when it is focused  ([#2283](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2283)) ([edab5f1](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/edab5f19e69459f7e10476324f9ecdd2df5bdb17))
* new task select rendering ([#2368](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2368)) ([f55a88a](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f55a88a58f024fb15f361eb967fad7277ae495d6))
* open file/folder when you press enter/backspace on focused item ([#2325](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2325)) ([e7efa21](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/e7efa2192e663466148afd815eed755af7b09656))
* people picker default selections ([#2579](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2579)) ([49b81bf](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/49b81bfeb342c00999408078a08a7b9b2ac557a4))
* people picker deletion ([#1877](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1877)) ([4164940](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/41649409887809838f0b531c29c3c21d3b2eb32e))
* people picker option labels ([#2207](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2207)) ([721dbe4](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/721dbe4bbf1fdffcf8debd000cc5ac1f6307f11f))
* People picker RTL renders, focus and storybook loading errors ([#1864](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1864)) ([21549f8](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/21549f8b77f84844a09598f0aa691f69285a4a6d))
* people picker single select mode ([#2541](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2541)) ([7032e88](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/7032e88c23584c23f698a27ad9d72c66ac3188c7))
* people-picker set focus on list navigation ([#2219](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2219)) ([6488705](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/648870562ae3c5868deacdf58214a323c465612a))
* people-picker uses show-max attribute ([#2527](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2527)) ([8691055](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/8691055e2eaabb3a69ea84c74496096a37817886))
* PeoplePicker add removePerson safety input check ([#2033](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2033)) ([6c678a8](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/6c678a8b4c3d8ccd4fed3ee4e9b6344ef5dce6a9))
* person card color ([#2533](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2533)) ([a83ae28](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/a83ae28fbb314312294d93e88f7449ff840ccc02))
* person component responsive issue ([#2297](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2297)) ([fbd397d](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/fbd397db8cd4d3c0867f80b8d35289306a4c51d8))
* person text visibility in custom properties ([#2298](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2298)) ([71aef6b](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/71aef6bca5bc7d605be9d540c27556c64d520707))
* person-card hover icons alignment ([#2531](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2531)) ([5bdaea6](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/5bdaea6b42e2a74b2991f3b8ede2c8baa0910081))
* prevent undefined in people-picker option labels for line two ([#2211](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2211)) ([274af2c](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/274af2c60e8c018c03bfc07d90dc514c66936438))
* prevent unsafe access to dom nodes in getDocumentDirection ([#1788](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1788)) ([03eddcf](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/03eddcf3d317fe3c726f16bec73b35f3ee2028ac))
* properly validate typed in emails and set them on comma or semi-colon key presses ([#1978](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1978)) ([2f03ee5](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/2f03ee552e7739fca674e27345f029ffd2598e2a))
* proxy provider sample ([#2515](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2515)) ([70211aa](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/70211aafc0422a6de14205937ac6dee611492793))
* proxy sample node ([#2491](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2491)) ([8ca93c9](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/8ca93c9acf9b377ba8b5051f664a8eef161b9e7c))
* react peer dependencies ([#2254](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2254)) ([f446898](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f446898fb619a36c3beca78374ce6dc40344ba82))
* react peer dependencies ([#2254](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2254)) ([d27858c](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/d27858cacec989e6de07be947800688c7c7d453e))
* remove disabled users from org tree in person card ([#1973](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1973)) ([911c07c](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/911c07c756179696565c883e3f83ed5d6b85cd93))
* remove mgt-agenda background colour ([#2476](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2476)) ([e82bd1f](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/e82bd1f64738cc734906c34f090e3658a8e11829))
* remove react peer dependency ([#2389](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2389)) ([4025f6f](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/4025f6fd7b7ae3080bc22c1514954e144875e9fe))
* remove the SdkVersion header when redirected to a non-graph endpoint ([#1947](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1947)) ([f9e2145](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f9e21459d93c032bffb39c59de02400e77d67442))
* remove the search field from the playground ([#1948](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1948)) ([4d39db7](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/4d39db75e76c7369425c9aabeb915fee0853c367))
* remove theme-toggle capability from custom CSS property and templating stories ([#2326](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2326)) ([fbbb1e3](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/fbbb1e3252a1f192da88509313bbc46a4426246e))
* removing unused mgt-teams-channel-picker tokens ([#2518](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2518)) ([6c39ea1](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/6c39ea18412ac85ab3a5e46e3e325ee763bb0ea6))
* request state update after setting selectedPeople ([#2163](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2163)) ([8f7eac6](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/8f7eac63aeb96622652f2f4c93cda3452a3e148d))
* resolving storybook a11y issues ([#2129](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2129)) ([0f68f60](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/0f68f60eaba27633d1c2ff83796e2a10e3a249de))
* restore provided msal public client behavior ([#1931](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1931)) ([dd41eb1](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/dd41eb1e0425695599952733d6564006c463d775))
* returning a JSON parsed version of the cache for insights types ([#2524](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2524)) ([0bbb487](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/0bbb487192cd05e5b5f6fe09b183e02d53a1099b))
* revert dot-options to use fluent controls ([#2424](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2424)) ([59ef61a](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/59ef61a702e6e42c92d710be3c888af77cb8d60c))
* rework parameter passing to click events in contact card section ([#1809](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1809)) ([5606784](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/560678406a31f314575373a36c0e14b5c90677f7))
* sample electron app fixes ([#2482](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2482)) ([cde48c8](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/cde48c8036c1f37d0f6e168814d94a9b782b454c))
* scopes used to query team channels ([#2519](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2519)) ([3da6333](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/3da63336591bf3c8d6b813986e5073bfd026c484))
* select first list available to display on mgt-todo ([#2456](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2456)) ([c5a5493](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c5a5493b6b10cd387dc83512965ebb22210280e2))
* set default contrast colors to allow highlighting in high contrast mode ([#2281](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2281)) ([4fc6460](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/4fc64602e878ba2fc2058093667f829735a2defd))
* set react peer dependency as range ([#2393](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2393)) ([2ee0078](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/2ee0078242cbe3a7f9ba46b615af5e19c7e7bbdf))
* set the avatar sizes for different mgt-person types with a single CSS prop ([#2457](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2457)) ([3e16476](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/3e1647609e61dbfca493409fd2aec25d8b38c70a))
* set the hover and focus color on dropdown item ([#1960](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1960)) ([39774a0](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/39774a093890837af79be0583f738e58b1c4042a)), closes [#1950](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1950)
* set the person details in all blocks ([#1959](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1959)) ([1b67cf6](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/1b67cf69f73560090e72461a997def9d17a34561)), closes [#1593](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1593)
* set the search icon to be on the same level with the input field ([#2043](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2043)) ([e105885](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/e105885b888a5bd93664219b568124fc258d9fbe))
* set the teams-channel-picker dropdown to overlay all other elements ([#2337](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2337)) ([1ce13ae](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/1ce13ae56494fa9257c53435d8d45129279765c3))
* show the dropdown when you focus on the people picker with tab key ([#1902](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1902)) ([7bd958e](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/7bd958eaad40c0ead80fa8a6db84155a7c95b053))
* storybook deployment ([#2553](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2553)) ([e979212](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/e979212f841e8ffae3e86a74fa15edaaa0cc0191))
* storybook footer accessibility ([#2369](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2369)) ([6dafa61](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/6dafa61c74ff666f9cfbd0d25de6e65321817630))
* styling of nested disambiguated components ([#2479](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2479)) ([3a60ed9](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/3a60ed998407f38d572a19c7e878e991934c0a38))
* suggested people aria labels ([#2335](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2335)) ([c14b6a3](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c14b6a3f9f8286f73efba4dd619a50491fe5a7d1))
* suggested people aria labels ([#2335](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2335)) ([071f3d5](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/071f3d517324627baecb1bb6bd8645e845b796c1))
* support target-id and initial-id attributes in mgt-todo ([#2407](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2407)) ([f2d4668](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f2d4668d3dfe9e15aa657602fa00a710fa29a2a3))
* task assignment button ([#2528](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2528)) ([63ad055](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/63ad0552e0f5af1c06c8684501d6d6d9e76edaa6))
* templating story for teams messages ([#2517](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2517)) ([3a51a52](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/3a51a525882a4a1eefb37990929042b47d572375))
* update accessibility features of the people picker component ([#1922](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1922)) ([37d1add](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/37d1add7caac2671fc9508d683c7eb4c83399c2a))
* Update mgt-file-upload.ts ([#2358](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2358)) ([f80accf](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f80accf8fcf88b83219647e65911bf4e7f7561e6))
* Update mgt-file-upload.ts ([#2358](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2358)) ([b095694](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/b095694599272726b4a40aa9fac0329fc5de7c79))
* update the text announced for person image, initials and email address ([#1923](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1923)) ([3b1dfba](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/3b1dfba3acdfa664dab045879c8fbbfe0e7a6e91))
* update typescript and ts-node versions for proxy samples ([#2020](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2020)) ([8170289](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/817028974616461e2abfa3a10ab6ec260d736089))
* use fixed graph client version ([#2274](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2274)) ([5dd2cad](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/5dd2cadd1d985b1b88d3bfe7fea5596e2636e1e9))
* use fluentui token to set person/login background ([#2435](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2435)) ([99884f8](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/99884f82009f05fae1b00678e17f7e35413b085b))
* use iterator to load events from event-query ([#2600](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2600)) ([0ba37cc](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/0ba37cc78f0e1efe20e22247149218deb1066c68))
* Use optional chaining for null user object ([#1856](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1856)) ([37b4cf2](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/37b4cf28cf5463a36c88c5decb938b0e6e3e0684))
* use optional chaining for search results hits ([#2447](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2447)) ([da8b7e3](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/da8b7e3d2b6b7cbd6e9f550f2d2b973fa905dc35))
* using mgt-search-results instead of mgt-search-box ([#2395](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2395)) ([f10a96b](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f10a96bb40495acf9066fd9c51f85f0e9883dd45))
* voice over for person in list ([#2206](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2206)) ([c60567d](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c60567d6c65f0252a7e736b91415c4a7f05d4842))

### [3.0.0](https://github.com/microsoftgraph/microsoft-graph-toolkit/compare/v2.11.2...v3.0.0) (2023-06-26)


### ⚠ BREAKING CHANGES

* removes MgtTeamsChannelPickerConfig
* removes use of user.read.all and group.read.all scopes for team/channel reading

### Features

* remove legacy mgt-dark and mgt-light theme and stories ([#2284](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2284)) ([4065d2f](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/4065d2f7e3bffd8a364a54c3f6534196827831d1))


### Bug Fixes

* **a11y:** unset custom color of storybook left chevrons ([#2595](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2595)) ([764bf12](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/764bf12429e0d69f7c4c0cfc305bfe7dbed1421f))
* **a11y:** set aria-expanded when open/closed ([#2405](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2405)) ([d084665](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/d08466597af59cf896808144ac1cd3a6aca60b11))
* Adding taxonomy-picker as an exported React component and updated used types ([4c06bd2](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/4c06bd26d67c1967a53c4edd4465b125314c7d9e))
* announce more options button on narrator in mgt-tasks ([#2399](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2399)) ([11ef91f](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/11ef91fc3e1be8b398aa461cde32055902161c5d))
* announce the label of the people-picker text field ([#2398](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2398)) ([f6ba11f](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f6ba11f0a5d5444c3db2b7125e7db0b494c64df9))
* announce the name of the selected user to remove in narrator ([#2360](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2360)) ([a6af856](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/a6af856ff175d2b013297950c6c0e8f63f29bad5))
* apply theme color to calendar icon of date input ([#2312](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2312)) ([68a05bb](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/68a05bb8470c0021e17baf091ae267d28306d43a))
* aspnet proxy provider sample ([#2594](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2594)) ([362339a](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/362339a2a80be148c1e20602f3ae639682b43137))
* caching story ([#2516](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2516)) ([2bd21d1](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/2bd21d1d7e99b0e7c49f95b2f4c5117d8d5d2be1))
* disable autocomplete ([#2481](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2481)) ([c9d2195](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c9d21956f81c4d1e7fdc795b6e6c79fb2f3a2a3b))
* disambiguated tagname and query selectors ([#2475](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2475)) ([f9f99e6](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f9f99e69779f25530efffa94b0c6936944d12abb))
* enable keyboard navigation in the picker ([#2324](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2324)) ([622f000](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/622f0007e9c76422e454edd11d276138d423fd2c))
* enable text spacings on login, channel-picker and teams-channel-picker ([#2413](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2413)) ([08819b1](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/08819b1c4a9388e791aed7cb02c338fd898179d8))
* ensure todo tasks are rendered in mgt-tasks ([#2480](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2480)) ([46afd78](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/46afd781f1b673945c90f5993e6a904641f346a0))
* files compact view in person card ([#2597](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2597)) ([6985717](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/69857170d87c37c1ce0c0f51baff34922f168e97))
* fix person-card to use fluent-card ([#2487](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2487)) ([6d3254d](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/6d3254d6c179278ab8a348063c00d1f40201dfe9)), closes [#2512](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2512)
* markdown table names ([#2473](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2473)) ([511e05b](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/511e05b932bb117ff602ba3a880ebdf05ea73b03))
* max picker list height ([#2431](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2431)) ([7a22138](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/7a22138cdd308938c771c3e9523dfeafbfed6d85))
* mgt-person narration ([#2493](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2493)) ([c14af08](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c14af0808c3b828a1ba5982e8c73544a2997bb6d))
* navigate mgt-people using left/right arrow keys when it is focused  ([#2283](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2283)) ([edab5f1](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/edab5f19e69459f7e10476324f9ecdd2df5bdb17))
* people picker default selections ([#2579](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2579)) ([49b81bf](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/49b81bfeb342c00999408078a08a7b9b2ac557a4))
* people picker single select mode ([#2541](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2541)) ([7032e88](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/7032e88c23584c23f698a27ad9d72c66ac3188c7))
* people-picker uses show-max attribute ([#2527](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2527)) ([8691055](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/8691055e2eaabb3a69ea84c74496096a37817886))
* person card color ([#2533](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2533)) ([a83ae28](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/a83ae28fbb314312294d93e88f7449ff840ccc02))
* person-card hover icons alignment ([#2531](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2531)) ([5bdaea6](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/5bdaea6b42e2a74b2991f3b8ede2c8baa0910081))
* proxy provider sample ([#2515](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2515)) ([70211aa](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/70211aafc0422a6de14205937ac6dee611492793))
* proxy sample node ([#2491](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2491)) ([8ca93c9](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/8ca93c9acf9b377ba8b5051f664a8eef161b9e7c))
* remove mgt-agenda background colour ([#2476](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2476)) ([e82bd1f](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/e82bd1f64738cc734906c34f090e3658a8e11829))
* removing unused mgt-teams-channel-picker tokens ([#2518](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2518)) ([6c39ea1](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/6c39ea18412ac85ab3a5e46e3e325ee763bb0ea6))
* returning a JSON parsed version of the cache for insights types ([#2524](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2524)) ([0bbb487](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/0bbb487192cd05e5b5f6fe09b183e02d53a1099b))
* revert dot-options to use fluent controls ([#2424](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2424)) ([59ef61a](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/59ef61a702e6e42c92d710be3c888af77cb8d60c))
* sample electron app fixes ([#2482](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2482)) ([cde48c8](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/cde48c8036c1f37d0f6e168814d94a9b782b454c))
* scopes used to query team channels ([#2519](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2519)) ([3da6333](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/3da63336591bf3c8d6b813986e5073bfd026c484))
* select first list available to display on mgt-todo ([#2456](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2456)) ([c5a5493](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c5a5493b6b10cd387dc83512965ebb22210280e2))
* set default contrast colors to allow highlighting in high contrast mode ([#2281](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2281)) ([4fc6460](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/4fc64602e878ba2fc2058093667f829735a2defd))
* set the avatar sizes for different mgt-person types with a single CSS prop ([#2457](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2457)) ([3e16476](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/3e1647609e61dbfca493409fd2aec25d8b38c70a))
* storybook deployment ([#2553](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2553)) ([e979212](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/e979212f841e8ffae3e86a74fa15edaaa0cc0191))
* styling of nested disambiguated components ([#2479](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2479)) ([3a60ed9](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/3a60ed998407f38d572a19c7e878e991934c0a38))
* support target-id and initial-id attributes in mgt-todo ([#2407](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2407)) ([f2d4668](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f2d4668d3dfe9e15aa657602fa00a710fa29a2a3))
* task assignment button ([#2528](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2528)) ([63ad055](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/63ad0552e0f5af1c06c8684501d6d6d9e76edaa6))
* templating story for teams messages ([#2517](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2517)) ([3a51a52](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/3a51a525882a4a1eefb37990929042b47d572375))
* use fluentui token to set person/login background ([#2435](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2435)) ([99884f8](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/99884f82009f05fae1b00678e17f7e35413b085b))
* use iterator to load events from event-query ([#2600](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2600)) ([0ba37cc](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/0ba37cc78f0e1efe20e22247149218deb1066c68))
* use optional chaining for search results hits ([#2447](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2447)) ([da8b7e3](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/da8b7e3d2b6b7cbd6e9f550f2d2b973fa905dc35))

## [3.0.0-rc.3](https://github.com/microsoftgraph/microsoft-graph-toolkit/compare/v3.0.0-rc.2...v3.0.0-rc.3) (2023-06-05)

## [3.0.0-rc.2](https://github.com/microsoftgraph/microsoft-graph-toolkit/compare/v3.0.0-rc.1...v3.0.0-rc.2) (2023-06-05)


### Bug Fixes

* add aria-label text  announce cancelling adding a new task ([#2359](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2359)) ([ffaa0e7](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/ffaa0e7066cb0b8dad0873f7e02189091baf0494))
* make spfx version script work for rc version ([#2396](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2396)) ([2d953fa](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/2d953faa7ea81b90c89a50eb9c1c402d5359526c))

## [3.0.0-rc.1](https://github.com/microsoftgraph/microsoft-graph-toolkit/compare/v2.11.1...v3.0.0-rc.1) (2023-06-02)


### Features

* add token overrides to theme switching ([#2328](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2328)) ([c49f70c](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c49f70ca84aac0130601df133bb7f77e09a7c023))
* added mgt-taxonomy-picker component ([#2172](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2172)) ([00f6565](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/00f65655e5cd30b228af5e4f35004362c5451882))
* mgt-picker selected-value attribute ([#2363](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2363)) ([abd00b4](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/abd00b4a1c59cc102658cf1acb14930d35731ecc))
* preview component support ([#2356](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2356)) ([38a13e1](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/38a13e106fbb79f5df75fea631a693d19ce31477))
* remove the Teams, TeamsMsal2 and the Msal providers in v3.x.x ([#2231](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2231)) ([c007abc](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c007abc82d351bf71ea23e50dd7b1c43e532e21c))


### Bug Fixes

* add a title tag to be announced for location svg ([#2285](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2285)) ([8c601c9](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/8c601c9cf949d6cc7f2396a017c1c338f2055694))
* change override design token logic ([#2384](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2384)) ([ed6c9b3](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/ed6c9b310a44a541ed784ebafb8429a9bfd9b004))
* editor tabs keyboard navigation ([#2371](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2371)) ([669b1ea](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/669b1ea8f1046837eab8d5f87d595ca738f57f67))
* execute beta queries and eliminate re-renders ([#2391](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2391)) ([70bef48](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/70bef481aecb1155a960b386d3dc87d0c2f74ede))
* include mgt-mock-provider as dependency to mgt ([#2336](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2336)) ([b166aaf](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/b166aaf80d8c8f737b1e94b3026220d052d027c5))
* keyboard navigation of login account list ([#2289](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2289)) ([af34d15](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/af34d15385db8cf229c7e056af29c6c8380ecc61))
* make tasks header navigable with the keyboard. ([#2313](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2313)) ([4747189](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/4747189c6d66a2a856274b931dac43e86002ff52))
* new task select rendering ([#2368](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2368)) ([f55a88a](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f55a88a58f024fb15f361eb967fad7277ae495d6))
* open file/folder when you press enter/backspace on focused item ([#2325](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2325)) ([e7efa21](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/e7efa2192e663466148afd815eed755af7b09656))
* person component responsive issue ([#2297](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2297)) ([fbd397d](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/fbd397db8cd4d3c0867f80b8d35289306a4c51d8))
* person text visibility in custom properties ([#2298](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2298)) ([71aef6b](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/71aef6bca5bc7d605be9d540c27556c64d520707))
* remove theme-toggle capability from custom CSS property and templating stories ([#2326](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2326)) ([fbbb1e3](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/fbbb1e3252a1f192da88509313bbc46a4426246e))
* set react peer dependency as range ([#2393](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2393)) ([2ee0078](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/2ee0078242cbe3a7f9ba46b615af5e19c7e7bbdf))
* set the teams-channel-picker dropdown to overlay all other elements ([#2337](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2337)) ([1ce13ae](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/1ce13ae56494fa9257c53435d8d45129279765c3))
* storybook footer accessibility ([#2369](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2369)) ([6dafa61](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/6dafa61c74ff666f9cfbd0d25de6e65321817630))
* using mgt-search-results instead of mgt-search-box ([#2395](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2395)) ([f10a96b](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f10a96bb40495acf9066fd9c51f85f0e9883dd45))

## v3.0.0-preview.2 (2023-05-09)

### New feature:

- use fluent UI to theme the tasks component (#2150) by Musale Martin
- move completed items to bottom of the list in mgt-todo (#2215) by Nickii Miaro
- Support for new component mgt-search-results (#2047) by Sébastien Levert
- update Todo component to new Fluent designs (#1967) by Nickii Miaro
- use fluentui to theme the person component. (#2072) by Musale Martin
- allow programmatically theming a component without the theme-toggle component (#2199) by Musale Martin
- migrate to eslint (#2125) by Gavin Barron
- added spec for mgt-taxonomy-picker (#2156) by Anoop T
- add custom CSS properties for the people picker flyout text (#2162) by Musale Martin
- Storybook authentication (#2048) by Sébastien Levert
- use fluentui tokens for theming file and file-list (#2044) by Musale Martin

### Bugs fixed:

- use fixed graph client version (#2274) by Gavin Barron
- react peer dependencies (#2254) by Gavin Barron
- remove mgt-spfx dependency from react webpart sample (#2196) by Gavin Barron
- fix mgt-spfx-utils packing (#2195) by Gavin Barron

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
