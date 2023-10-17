# Changelog

All notable changes to this project will be documented in this file. See [standard-version](https://github.com/conventional-changelog/standard-version) for commit guidelines.

## [3.1.3](https://github.com/microsoftgraph/microsoft-graph-toolkit/compare/v3.1.2...v3.1.3) (2023-10-06)


### Bug Fixes

* **a11y:** mgt-file and mgt-picker visibility issues in dark-mode ([#2667](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2667)) ([239bfb0](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/239bfb0978569ae55343afd0c998edb370d9ea98))
* add Group entity to IDynamicPerson type and introduce typeguards to find the entity type ([#2688](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2688)) ([b3bc50d](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/b3bc50dda72e1865ff3314f2bd2f8b93739d6710))
* add spaces to presence hover text in mgt-person ([#2693](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2693)) ([f50e6ab](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/f50e6ab514ad9b1e4deb0113e4f50558c8ec49c3))
* disable todo checkboxes and inputs in read-only mode ([#2745](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2745)) ([d19f078](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/d19f078abe52d6b0f3e9e9899db3a783c54a9133))
* ensure batch url resources start with / ([#2740](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2740)) ([247f37a](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/247f37ae57074ec73d9ba17c0f5e96a54ab7149a))
* ensure msal public client application is initialized ([#2702](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2702)) ([b9fcfe7](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/b9fcfe7ee154c98ad59d43dc9c5678354a1c423d))
* ensure people-picker search works when userIds are supplied([#2736](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2736)) ([a724b05](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/a724b05163ff1344c346df691177ae5803587051))
* initials rendering in mgt-person ([#2764](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2764)) ([882aaf6](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/882aaf6ac20666640eb511f36139826b8ec720e0))
* **MgtProfile:** Fix handling of null values for educations & work positions ([#2717](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2717)) ([ba381c8](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/ba381c83ac182e855edce9217c7c072197fa165d))
* typing for template props data context ([#2754](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2754)) ([c9023c2](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/c9023c22a3f2d8cef4c52416ab8e9fd872c48926))
* update mgt-taxonomy-picker colors to match mgt-picker ([#2747](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2747)) ([be7add8](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/be7add83c77809f6210d1bf877d8009840da5c47))

## [3.1.2](https://github.com/microsoftgraph/microsoft-graph-toolkit/compare/v3.1.1...v3.1.2) (2023-09-05)


### Bug Fixes

* adds pointer cursor to logged in accounts ([#2674](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2674)) ([11e5a1c](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/11e5a1cb319812afdb2355c4973e9b072ea680e0))
* disable open on click behavior ([#2685](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2685)) ([10b25f9](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/10b25f97ff9666bb9e0c0fadf5b453ced9da84fc))
* set custom css token --default-font-family to apply to all elements in DOM ([#2677](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2677)) ([cb69e01](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/cb69e0160b41773f65d88a4ea457fc18b7360f51))
* use correct scope for group member resolution ([#2690](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2690)) ([ca313c1](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/ca313c1c7855dde9385981518392351662365428))

## [3.1.1](https://github.com/microsoftgraph/microsoft-graph-toolkit/compare/v3.1.0...v3.1.1) (2023-08-17)


### Bug Fixes

* dismiss login flyout when moving out of the popup ([#2637](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2637)) ([263f36f](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/263f36f56665e7e041be1defe80677d31f65af7b))
* use pointer cursor when person card enabled in mgt-person. ([#2652](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2652)) ([48ea18b](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/48ea18bc28c39d65629d3095db538d2c01fe7109))

## [3.1.0](https://github.com/microsoftgraph/microsoft-graph-toolkit/compare/v3.0.1...v3.1.0) (2023-07-28)


### Features

* add canary url to allowed endpoints for graph ([#2635](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2635)) ([ec621cd](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/ec621cd37a8eb7f91356944fb29ad85699ee89fa))


### Bug Fixes

* **a11y:** add distinct name definitions for copy code buttons in storybook overview ([#2622](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2622)) ([4e52f41](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/4e52f41398b1a5771ada9fe54a67dba4e4b84397))
* add a title text if displaying images only. ([#2625](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2625)) ([28703c9](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/28703c96dd5366b8543fab2a160994db34a9c13b))
* check the file type being uploaded before performing upload ([#2584](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2584)) ([7fb265c](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/7fb265c744cbbf53d49fe6d5c3d45bc8493eb2fd))
* remove agenda tooltip ([#2621](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2621)) ([27e1fc9](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/27e1fc9c61e1957dd8addacb3b667de9546fb358))

## [3.0.1](https://github.com/microsoftgraph/microsoft-graph-toolkit/compare/v3.0.0...v3.0.1) (2023-07-18)


### Bug Fixes

* **a11y:** unset custom color of storybook left chevrons ([#2595](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2595)) ([764bf12](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/764bf12429e0d69f7c4c0cfc305bfe7dbed1421f))
* add `InAConferenceCall` activity when availability is `Busy` ([#2585](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2585)) ([bd17195](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/bd17195d22a90ab8a820528c91d8c7d6a89191ca))
* add class to people-picker styles story to enable custom css ([#2605](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2605)) ([dcec953](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/dcec953ecdc2e721a280a3a11aca5be1596311dc))
* add font family to tasks ([#2603](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2603)) ([e380b4a](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/e380b4ae98095069a260f459ec6fd6d156d97762))
* add login custom styles, removes style not in use ([#2587](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2587)) ([7ba98e4](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/7ba98e4359854adb2658207be8c8e64bbeeda730))
* adds customHosts support for non-graph domain requests ([#2592](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2592)) ([1f97215](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/1f97215f03324ebbc255ef57fe7658c6978ff0dc))
* announce teams channel results when you type ([#2561](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2561)) ([5260ce0](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/5260ce05c125f59ee56a3cbac7eba18e10de523c))
* aspnet proxy provider sample ([#2594](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2594)) ([362339a](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/362339a2a80be148c1e20602f3ae639682b43137))
* correct sppkg upload script ([#2552](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2552)) ([8b20d84](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/8b20d8489c0043060714a1ca1401d0c7433a7c05))
* files compact view in person card ([#2597](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2597)) ([6985717](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/69857170d87c37c1ce0c0f51baff34922f168e97))
* people picker default selections ([#2579](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2579)) ([49b81bf](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/49b81bfeb342c00999408078a08a7b9b2ac557a4))
* use iterator to load events from event-query ([#2600](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/2600)) ([0ba37cc](https://github.com/microsoftgraph/microsoft-graph-toolkit/commit/0ba37cc78f0e1efe20e22247149218deb1066c68))

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
