import * as Utils from '../src/utils/Utils';

describe('objectEquals', () => {
  const circularObject: any = {};
  circularObject.a = circularObject;

  const circularArray: any = [];
  circularArray[0] = circularArray;

  // Any other object that is not an object literal or an array will compare by reference
  const simpleDate = new Date(0);

  it.each([
    [{}, {}],
    [{ a: 1, b: true, c: 'foo' }, { c: 'foo', b: true, a: 1 }],
    [{ a: [1, 2, 3] }, { a: [1, 2, 3] }],
    [{ a: { b: { c: 1 } } }, { a: { b: { c: 1 } } }],
    [{ a: [1, [2, [3]]] }, { a: [1, [2, [3]]] }],
    [circularObject, circularObject],
    [circularArray, circularArray],
    [{ a: circularObject, b: circularArray }, { a: circularObject, b: circularArray }],
    [{ a: simpleDate }, { a: simpleDate }]
  ])('should return true between %p and %p', (o1: any, o2: any) => {
    expect(Utils.equals(o1, o2)).toBe(true);
  });

  it.each([
    [{ a: {} }, { a: [] }],
    [{ a: [1, 2, 3] }, { a: [3, 2, 1] }],
    [{ a: [1, [2, [3]]] }, { a: [1, [2, [4]]] }],
    [{ a: { b: [{ c: 1 }, { d: [2, 3] }] } }, { a: { b: [{ c: 1 }, { d: [3, 2] }] } }],
    [{ a: new Date(0) }, { a: new Date(0) }]
  ])('should return false between %p and %p', (o1: any, o2: any) => {
    expect(Utils.equals(o1, o2)).toBe(false);
  });
});
