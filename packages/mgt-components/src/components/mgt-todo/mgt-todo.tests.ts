/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT License.
 * -------------------------------------------------------------------------------------------
 */

import { expect } from '@open-wc/testing';
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
});
