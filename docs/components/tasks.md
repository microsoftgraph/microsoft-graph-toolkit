# Tasks Control

## Description

The Tasks Control enables the user to view, add, remove, complete, or edit tasks. It works with any tasks in Microsoft Planner or Microsoft To-Do.

## Example

````html
    <mgt-tasks></mgt-tasks>
````

## Properties

| Attribute | Description |
| -- | -- |
| `data-source="todo|planner"` | Sets the Data source for tasks, either Microsoft To-Do, or Microsoft Planner |
| `read-only` | Sets the task interface to be read only (no adding or removing tasks) |
| `initial-id="planner_id/folder_id"` | Sets the initially displayed planner or folder to the provided ID |
| `initial-bucket-id="bucket_id"` | Sets the initially displayed bucket (Planner Data-Source Only) to the provided ID |
| `target-id="planner_id/folder_id"` | Locks the tasks interface to the provided planner or folder ID |
| `target-bucket-id="bucket_id"` | Locks the tasks interface to the provided bucket id (Planner Data-Source Only) |

ex: 
````html
<mgt-tasks read-only initial-id="12345"></mgt-tasks>
````

## Custom CSS Variables

````css
mgt-tasks {
--tasks-header-padding
--tasks-header-margin 

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

--tasks-new-task-name-margin

--task-margin
--task-box-shadow
--task-background
--task-border

--task-header-color
--task-header-margin

--task-detail-icon-margin

--task-new-margin
--task-new-border
--task-new-line-margin
--tasks-new-line-border
--task-new-input-margin
--task-new-input-padding
--task-new-input-font-size
--task-new-input-active-border
--task-new-select-border

--task-new-add-button-background
--task-new-add-button-disabled-background
--task-new-cancel-button-color

--task-complete-background
--task-complete-border
--task-complete-header-color
--task-complete-detail-color
--task-complete-detail-icon-color
}
````

## Graph Scopes

For the O365 Planner Data-Source, fetching and reading tasks requires the `'Groups.Read.All'` permission, and for adding / updating / removing tasks, the `'Groups.ReadWrite.All'` permission is required.

For the Microsoft Todo Data-Source, the `'Tasks.Read'` permission is required for fetching and reading tasks, while for adding / updating / removing tasks, the `'Tasks.ReadWrite'` permission is required.

## Authentication

The login control leverages the global authentication provider described in the [authentication documentation](./../providers.md).
