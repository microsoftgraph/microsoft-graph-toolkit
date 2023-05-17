import * as jspath from 'jspath/lib/jspath.js';
import { isEmpty } from 'lodash-es';

export class ObjectHelper {
  /**
   * Get object proeprty value by its deep path.
   * @param object the object containg the property path
   * @param path the property path to get
   * @param delimiter if multiple matches are found, sepcifiy the delimiter character to use to separate values in the returned string
   * @returns the property value as string if found, 'undefined' otherwise
   */
  public static byPath(object: object, path: string, delimiter?: string): string {
    let isValidPredicate = true;

    // Test if the provided path is a valid predicate https://www.npmjs.com/package/jspath#documentation
    try {
      jspath.compile(`.${path}`);
    } catch (e) {
      isValidPredicate = false;
    }

    if (path && object && isValidPredicate) {
      try {
        // jsPath always returns an array. See https://www.npmjs.com/package/jspath#result
        const value: object[] = jspath.apply(`.${path}`, object);

        // Empty array returned by jsPath
        if (isEmpty(value)) {
          // i.e the value to look for does not exist in the provided object
          return undefined;
        }

        // Check if value is an object
        // - Arrays of objects will return '[object Object],[object Object]' etc.
        if (value.toString().indexOf('[object Object]') !== -1) {
          // Returns the stringified array
          return JSON.stringify(value);
        }

        if (delimiter && value.length > 1) {
          return value.join(delimiter);
        }

        // Use the default behavior of the toString() method. Arrays of simple values (string, integer, etc.) will be separated by a comma (',')
        return value.toString();
      } catch (error) {
        // Case when unexpected string or tokens are passed in the path
        return null;
      }
    } else {
      return undefined;
    }
  }
}
