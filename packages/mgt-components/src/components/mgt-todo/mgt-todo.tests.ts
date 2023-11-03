/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT License.
 * -------------------------------------------------------------------------------------------
 */

import { expect, fixture, html } from '@open-wc/testing';
import { MockProvider, Providers } from '@microsoft/mgt-element';
import { MgtTodo, registerMgtTodoComponent } from './mgt-todo';

describe('mgt-todo - tests', () => {
  before(() => {
    registerMgtTodoComponent();
    Providers.globalProvider = new MockProvider();
  });

  it('has required scopes', () => {
    expect(MgtTodo.requiredScopes).to.have.members(['tasks.readwrite', 'tasks.read']);
  });

  it('should render', async () => {
    const todo = await fixture(html`
      <mgt-todo></mgt-todo>
    `);

    await expect(todo).shadowDom.to.equal(
      `<mgt-picker resource="me/todo/lists" scopes="tasks.read, tasks.readwrite" key-name="displayName" placeholder="Select a task list" selected-value="Tasks">
      </mgt-picker>
      <div class="task new-task incomplete" dir="ltr">
        <fluent-text-field autocomplete="off" appearance="outline" class="new-task outline" id="new-task-name-input" aria-label="New Task Input" placeholder="Add a task" current-value="" type="text">
          <div slot="start" class="start">
            <span class="add-icon">
              <svg width="16" height="15" viewBox="0 0 16 15" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
                <path d="M8.39563 0.5C8.39563 0.223858 8.17177 0 7.89563 0C7.61949 0 7.39563 0.223858 7.39563 0.5V7H0.89563C0.619487 7 0.39563 7.22386 0.39563 7.5C0.39563 7.77614 0.619487 8 0.89563 8H7.39563V14.5C7.39563 14.7761 7.61949 15 7.89563 15C8.17177 15 8.39563 14.7761 8.39563 14.5V8H14.8956C15.1718 8 15.3956 7.77614 15.3956 7.5C15.3956 7.22386 15.1718 7 14.8956 7H8.39563V0.5Z" fill=""></path>
              </svg>
            </span>
          </div>
        </fluent-text-field>
      </div>
      <div class="tasks" dir="ltr">
      </div>`,
      { ignoreAttributes: ['src'] }
    );
  });
});
