# Design Tokens

## Status

| Date               | Version | Author           | Status |
| ------------------ | ------- | ---------------- | ------ |
| October 10th, 2022 | v1.0    | Sébastien Levert | Draft  |

## Scope

All the Microsoft Graph Toolkit components are affected by this specification. This is a living document and should be updated with new components or redesign of existing components. This is the central location for defining our design tokens across all of the Microsoft Graph Toolkit.

## Context

When building the Microsoft Graph Toolkit, we are using defined themes that help our solution to be more accessible depending on the preferred user's mode : light or dark. These 2 themes come with a series of colors that make them visually appealing while providing the right set of color contrasts. Though, while building the Microsoft Graph Toolkit, new tokens were introduced and we identified where we might have had gaps or repeated tokens that could be rationalized. This document describes the outcome of this exercise and will be the best approach to deliver theming for our customers.

## Available Tokens

### Common Tokens

| Token name                          | Theme: mgt-light                                      | Theme: mgt-dark |
| ----------------------------------- | ----------------------------------------------------- | --------------- |
| --color                             | <span style="background-color:#000000">#000000</span> | #ffffff         |
| --color-sub1                        | #323130                                               | #f3f2f1         |
| --color-sub2                        | #717171                                               | #c8c6c4         |
| --color-sub3                        | #727170                                               | #727170         |
| --background-color                  |                                                       |                 |
| --background-color--active          |                                                       |                 |
| --background-color--hover           |                                                       |                 |
| --background-color-sub1             |                                                       |                 |
| --background-color-sub2             |                                                       |                 |
| --box-shadow                        |                                                       |                 |
| --box-shadow-color                  |                                                       |                 |
| --default-font-size                 | 14px                                                  | 14px            |
| --icon-color                        |                                                       |                 |
| --input-background-color            |                                                       |                 |
| --input-background-color--hover     |                                                       |                 |
| --input-border                      |                                                       |                 |
| --input-border-top                  |                                                       |                 |
| --input-border-right                |                                                       |                 |
| --input-border-bottom               |                                                       |                 |
| --input-border-left                 |                                                       |                 |
| --input-border-color--hover         |                                                       |                 |
| --input-border-color--focus         |                                                       |                 |
| --line-seperator-color              |                                                       |                 |
| --list-background-color             |                                                       |                 |
| --list-item-background-color--hover |                                                       |                 |
| --placeholder-color                 |                                                       |                 |
| --placeholder-color--focus          |                                                       |                 |
| --tab-line-color                    |                                                       |                 |
| --tab-line-color--hover             |                                                       |                 |
| --tab-background-color--hover       |                                                       |                 |
| --theme-primary-color               | #0078d7                                               | #0078d7         |
| --theme-dark-color                  | #005a9e                                               | #005a9e         |
| --title-color-main                  |                                                       |                 |
| --title-color-subtitle              |                                                       |                 |
| --title-color-sub2                  |                                                       |                 |

### mgt-agenda


| Token name                 | Common token | Theme: mgt-light | Theme: mgt-dark |
| -------------------------- | ------------ | ---------------- | --------------- |
| --agenda-header-color      | --color-sub1 | Inherits         | Inherits        |
| --event-background-color   |              |                  |                 |
| --event-border             |              |                  |                 |
| --event-box-shadow         |              |                  |                 |
| --event-margin             |              |                  |                 |
| --event-padding            |              |                  |                 |
| --event-time-color         |              |                  |                 |
| --event-time-font-size     |              |                  |                 |
| --event-subject-color      |              |                  |                 |
| --event-subject-font-size  |              |                  |                 |
| --event-location-font-size |              |                  |                 |
| --event-location-color     |              |                  |                 |