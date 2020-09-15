/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Crude implementation of equivalence between the two specified arguments.
 *
 * The primary intent of this function is for comparing data contexts, which
 * are expected to be object literals with potentially nested structures and
 * where leaf values are primitives.
 */
export function equals(o1: any, o2: any) {
  return equalsInternal(o1, o2, new Set());
}

/**
 * Not exposed as it would undesirably leak implementation detail (`refs` argument).
 *
 * The `refs` argument is used to avoid infinite recursion due to circular references.
 *
 * @see equals
 */
function equalsInternal(o1: any, o2: any, refs: Set<object>) {
  const o1Label = Object.prototype.toString.call(o1);
  const o2Label = Object.prototype.toString.call(o2);
  if (o1Label === o2Label && o1Label === '[object Object]' && !refs.has(o1)) {
    refs.add(o1);
    for (const k in o1) {
      if (!equalsInternal(o1[k], o2[k], refs)) {
        return false;
      }
    }
    for (const k in o2) {
      if (!o1.hasOwnProperty(k)) {
        return false;
      }
    }
    return true;
  }
  if (o1Label === o2Label && o1Label === '[object Array]' && !refs.has(o1)) {
    refs.add(o1);
    if (o1.length !== o2.length) {
      return false;
    }
    for (let i = 0; i < o1.length; i++) {
      if (!equalsInternal(o1[i], o2[i], refs)) {
        return false;
      }
    }
    return true;
  }
  // Everything else requires strict equality (e.g. primitives, functions, dates)
  return o1 === o2;
}

/**
 * Compares two arrays if the elements are equals
 * Should be used for arrays of primitive types
 *
 * @export
 * @template T the type of the elements in the array (should be primitive)
 * @param {T[]} arr1
 * @param {T[]} arr2
 * @returns true if both arrays contain the same items or if both arrays are null or empty
 */
export function arraysAreEqual<T>(arr1: T[], arr2: T[]) {
  if (arr1 === arr2) {
    return true;
  }

  if (!arr1 || !arr2) {
    return false;
  }

  if (arr1.length !== arr2.length) {
    return false;
  }

  if (arr1.length === 0) {
    return true;
  }

  const setArr1 = new Set(arr1);

  for (const i of arr2) {
    if (!setArr1.has(i)) {
      return false;
    }
  }

  return true;
}
