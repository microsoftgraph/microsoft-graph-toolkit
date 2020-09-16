# MgtTodo Component

The Todo component enables the user to view, add, remove, complete, or edit tasks. It works with tasks from the Microsoft To-Do API.  

## Properties

| Attribute | Property | Description |
| -- | -- | -- |
| read-only | readOnly | A Boolean to set the task interface to be read only (no adding or removing tasks). Default is `false`. |
| hide-header | hideHeader | A Boolean to show or hide the header of the component. Default is `false`. |
| hide-options | hideOptions | A Boolean to show or hide the options in tasks. Default is `false`.
| initial-id="folder_id" | initialId | A string ID to set the initially displayed folder to the provided ID. |
| target-id="folder_id"| targetId | A string ID to lock the tasks interface to the provided folder ID. |
| N/A | isNewTaskVisible  | Determines whether new task view is visible at render. |
| N/A | taskFilter  | An optional function to filter which tasks are shown to the user. |

The following example shows only tasks from the folder with ID *12345* and does not allow the user to create new tasks.

```html
<mgt-todo read-only initial-id="12345"></mgt-todo>
```

## Custom CSS variables

````css
mgt-todo {
  --tasks-background-color
  --tasks-header-padding
  --tasks-title-padding
  --tasks-plan-title-font-size
  --tasks-plan-title-padding

  --tasks-new-button-width
  --tasks-new-button-height
  --tasks-new-button-color
  --tasks-new-button-background
  --tasks-new-button-border
  --tasks-new-button-hover-background
  --tasks-new-button-active-background

  --task-margin
  --task-background
  --task-border
  --task-header-color
  --task-header-margin

  --task-new-margin
  --task-new-border
  --task-new-input-margin
  --task-new-input-padding
  --task-new-input-font-size
  --task-new-select-border

  --task-new-add-button-background
  --task-new-add-button-disabled-background
  --task-new-cancel-button-color

  --task-complete-background
  --task-complete-border

  --task-icon-alignment: flex-start (default) | center | flex-end
  --task-icon-background
  --task-icon-background-completed
  --task-icon-border
  --task-icon-border-completed
  --task-icon-border-radius
  --task-icon-color
  --task-icon-color-completed
}
````

## Events
| Event | Detail | Description |
| --- | --- | --- |
| taskAdded | The detail contains the respective `task` object | Fires when a new task has been created. |
| taskChanged | The detail contains the respective `task` object | Fires when task metadata has been changed, such as marking completed. |
| taskClick | The detail contains the respective `task` object | Fires when the user clicks or taps on a task. |
| taskRemoved | The detail contains the respective `task` object | Fires when an existing task has been deleted. |

## Templates

The `tasks` component supports several [templates](../templates.md) that allow you to replace certain parts of the component. To specify a template, include a `<template>` element inside a component and set the `data-type` value to one of the following:

| Data type     | Data context              | Description                                                       |
| ---------     | ------------------------- | ----------------------------------------------------------------- |
| task     | task: a to-do task object | replaces the whole default task. |
| task-details | task: a to-do task object | template replaces the details section of the task. |

The following example defines a template for the tasks component.

```html
<mgt-todo>
    <template data-type="task-details">
        <div>
            Importance Level: {{task.importance}}
        </div>
    </template>
</mgt-todo>
```

## Microsoft Graph permissions

This control uses the following Microsoft Graph APIs and permissions.

| Resource | Permission |
| - | - |
| /me/todo/lists/ | Tasks.ReadWrite |
| /me/todo/lists/{todoTaskListId}/tasks | Tasks.ReadWrite |
| /me/todo/lists/{todoTaskListId}/tasks/{taskId} | Tasks.ReadWrite |
