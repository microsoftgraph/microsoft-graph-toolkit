# Agenda Control

## Description

The `mgt-agenda` web component can be used to represent events in a user or group calendar. By default, the calendar displays the current logged in user events for the current day. The component can also use any endpoint that returns events from the Microsoft Graph. 

## Example

```html
<mgt-agenda group-by-day></mgt-agenda>
```

![mgt-agenda](./images/mgt-agenda.png)

## Properties

By default, the `mgt-agenda` component fetches events from the `/me/calendarview` endpoint and displays events for the current day. Use the following attributes to change this behavior:

| attribute | Description |
| --- | --- |
| `group-by-day` | a boolean value to group events by day - by default events are not grouped |
| `date` | a string representing the start date of the events to fetch from the Microsoft Graph. Value should be in a format that can be parsed by the [Date constructor](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Date) - value has no effect if `event-query` attribute is set |
| `days` | a number of days to fetch from the Microsoft Graph - default is 3 - value has no effect if `event-query` attribute is set. |
| `event-query` | a string representing an alternative query to be used when fetching events from the Microsoft Graph. Optionally, add the delegated scope at the end of the string by delimiting it with `|` (ex: `"/groups/GROUP-ID-GUID/calendar/calendarView | group.read.all"`) |
| `events` | an array of events to get or set the list of events rendered by the component - use this property to access the events loaded by the component. Set this value to load your own events - if value is set by developer, the `date`, `days`, or `event-query` attributes have no effect |

Ex:

```html
<mgt-agenda
  group-by-day
  date="May 7, 2019"
  days="3"
  event-query="/me/events?orderby=start/dateTime"
  ></mgt-agenda>
```

## Custom properties

The `mgt-agenda` component defines these custom properties

```css
mgt-agenda {
  --event-box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.092);
  --event-margin: 0px 10px 14px 10px;
  --event-padding: 8px 0px;
  --event-background: #ffffff;
  --event-border: solid 2px rgba(0, 0, 0, 0);

  --agenda-header-margin: 40px 10px 14px 10px;
  --agenda-header-font-size: 24px;
  --agenda-header-color: #333333;

  --event-time-font-size: 12px;
  --event-time-color: #000000;

  --event-subject-font-size: 19px;
  --event-subject-color: #333333;

  --event-location-font-size: 12px;
  --event-location-color: #000000;
}
```

## Templates

The `mgt-agenda` support several [templates](../style.md) that allow you to replace certain parts of the component. To specify a template, simply include a `<template>` element inside of a component and set the `data-type` value to one of the following:


### `default` (or when no value is provided)

The default template replaces the entire component with your own. The `events` are passed to the templates as data context

### `no-data`

The template used when no events are available. No data context is passed

### `loading`

The template used when data is loading. No data context is passed

### `event`

The template used to render each event. The `event` is passed to the template as data context

### `header`

The template used to render the header for each day. The `header` is passed to the template as data context

### `other`

The template used to render additinal content for each event. The `event` is passed to the template as data context

Ex:

```html
<mgt-agenda>
  <template data-type="event">
    <button class="eventButton">
      <div class="event-subject">{{ event.subject }}</div>
      <div data-for="attendee in event.attendees">
        <mgt-person
          person-query="{{ attendee.emailAddress.name }}"
          show-name
          show-email>
        </mgt-person>
      </div>
    </button>
  </template>
  <template data-type="no-data">
    Where is the data
  </template>
</mgt-agenda>
```

## Graph scopes

This component uses the following Microsoft Graph APIs and permissions:

| resource | permission/scope |
| - | - |
| [/me/calendarview](https://docs.microsoft.com/en-us/graph/api/calendar-list-calendarview?view=graph-rest-1.0) | `Calendars.Read` |

The component allows you to specify a different Microsoft Graph query to call (such as `/groups/{id}/calendar/calendarView`). In this case, append the scope at the end of the string, delimited by `|`

## Authentication

The login control leverages the global authentication provider described in the [authentication documentation](./../providers.md). 

