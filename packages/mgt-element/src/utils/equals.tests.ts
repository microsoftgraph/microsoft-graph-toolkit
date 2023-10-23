/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { it } from '@jest/globals';
import { equals } from './equals';

describe('objectEquals', () => {
  const circularObject = { a: undefined };
  circularObject.a = circularObject;

  const circularArray: unknown[] = [];
  circularArray[0] = circularArray;

  // Any other object that is not an object literal or an array will compare by reference
  const simpleDate = new Date(0);

  it.each([
    [{}, {}],
    [
      { a: 1, b: true, c: 'foo' },
      { c: 'foo', b: true, a: 1 }
    ],
    [{ a: [1, 2, 3] }, { a: [1, 2, 3] }],
    [{ a: { b: { c: 1 } } }, { a: { b: { c: 1 } } }],
    [{ a: [1, [2, [3]]] }, { a: [1, [2, [3]]] }],
    [circularObject, circularObject],
    [circularObject, { a: circularObject }],
    [circularArray, circularArray],
    [
      { a: circularObject, b: circularArray },
      { a: circularObject, b: circularArray }
    ],
    [{ a: simpleDate }, { a: simpleDate }]
  ])('should return true between %p and %p', (o1: unknown, o2: unknown) => {
    expect(equals(o1, o2)).toBe(true);
  });

  it.each([
    [{ a: {} }, { a: [] }],
    [{ a: [1, 2, 3] }, { a: [3, 2, 1] }],
    [{ a: [1, [2, [3]]] }, { a: [1, [2, [4]]] }],
    [{ a: { b: [{ c: 1 }, { d: [2, 3] }] } }, { a: { b: [{ c: 1 }, { d: [3, 2] }] } }],
    [{ a: new Date() }, { a: new Date() }],
    [circularObject, circularArray],
    [circularObject, { b: circularObject }]
  ])('should return false between %p and %p', (o1: unknown, o2: unknown) => {
    expect(equals(o1, o2)).toBe(false);
  });
});
