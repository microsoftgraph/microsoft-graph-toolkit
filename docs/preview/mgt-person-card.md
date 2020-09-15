# MgtPersonCard Component

The Person Card has been heavily redesigned to pack in even more Graph features. In V2, the MgtPersonCard includes sections for:

* Section Overview with a Teams quick messenger and compact views for each other section.
* Contact details - email, phone, office, etc.
* Organization - Managers and coworkers
* Emails - where both users are participants
* Shared files *(work in progress)*
* Personal profile - languages, education, interests, birthday, etc.

## Additional Details (work in progress)

The V1 Person Card provided an additional details area below the contact details to show custom third-party content. Going forward, we've expanded this small area into an entire custom section. The custom section tab will be hidden by default, but when the additional details templates are provided a new section will appear.

Specify the following templates to light up a custom section:

| Data type | Data context | Description |
| - | - | - |
| additional-details-compact | `person`: The person details object | The template used to specify the compact view of the custom section in the overview section. |
| additional-details-full | `person`: The person details object | The template used to specify the full view of the custom section. |
| additional-details-icon | `person`: The person details object | The template used to specify the tab icon of the custom section. |

## Templates (work in progress)

The Person-Card component uses [templates](../templates.md) that allow you to add or replace portions of the component. To specify a template, include a `<template>` element inside of a component and set the `data-type` value to one of the following.

| Data type | Data context | Description |
| - | - | - |
| no-data | null | The template used when no data is available.
| default | `person`: The person details object <br> `personImage`: The URL of the image | The default template replaces the entire component with your own. |
| person-details | `person`: The person details object <br> `personImage`: The URL of the image | The template used to render the top part of the person card. |
| contact-section | `person`: The person details object | The template used to override the contact details section. |
| organization-section | `person`: The person details object | The template used to override the organization section. |
| messages-section | `person`: The person details object | The template used to override the email messages section. |
| files-section | `person`: The person details object | The template used to override the files section. |
| profile-section | `person`: The person details object | The template used to override the profile section. |
