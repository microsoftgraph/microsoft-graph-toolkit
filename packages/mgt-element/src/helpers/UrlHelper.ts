export class UrlHelper {
  /**
   * Test if the provided string is a valid URL
   * @param url the URL to check
   */
  public static isValidUrl(url: string): boolean {
    // eslint-disable-next-line no-useless-escape
    return /((([A-Za-z]{3,9}:(?:\/\/)?)(?:[\-;:&=\+\$,\w]+@)?[A-Za-z0-9\.\-]+|(?:www\.|[\-;:&=\+\$,\w]+@)[A-Za-z0-9\.\-]+)((?:\/[\+~%\/\.\w\-_]*)?\??(?:[\-\+=&;%@\.\w_]*)#?(?:[\.\!\/\\\w]*))?)/.test(
      url
    );
  }

  /**
   * Get the value of a querystring
   * @param  {String} field The field to get the value of
   * @param  {String} url   The URL to get the value from (optional)
   * @return {String}       The field value
   */
  public static getQueryStringParam(field: string, url: string): string {
    const href = url ? url : window.location.href;
    const reg = new RegExp('[?&#]' + field + '=([^&#]*)', 'i');
    const qs = reg.exec(href);
    return qs ? qs[1] : '';
  }

  /**
   * @param {String} field The field name of the query string to remove
   * @param {String} sourceURL The source URL
   * @return {String}       The updated URL
   */
  public static removeQueryStringParam(field: string, sourceURL: string): string {
    let rtn = sourceURL.split('?')[0];
    let param = null;
    let paramsArr = [];
    const hash = window.location.hash;
    const queryString = sourceURL.indexOf('?') !== -1 ? sourceURL.split('?')[1] : '';

    if (queryString !== '') {
      paramsArr = queryString.split('&');
      for (let i = paramsArr.length - 1; i >= 0; i -= 1) {
        param = paramsArr[i].split('=')[0];
        if (param === field) {
          paramsArr.splice(i, 1);
        }
      }

      if (paramsArr.length > 0) {
        rtn = rtn + '?' + paramsArr.join('&').replace(hash, '') + hash;
      }
    }
    return rtn;
  }

  /**
   * Add or replace a query string parameter
   * @param url The current URL
   * @param param The query string parameter to add or replace
   * @param value The new value
   */
  public static addOrReplaceQueryStringParam(url: string, param: string, value: string): string {
    param = param.replace(/[.~*()]/g, ''); // // Ensure param is safe from DOS attacks - so we strip away RegEx special characters
    const re = new RegExp('[\\?&]' + param + '=([^&#]*)');
    const match = re.exec(url);
    let delimiter;
    let newString;

    if (match === null) {
      // Append new param
      const hash = window.location.hash && window.location.hash !== '' ? window.location.hash : '#';
      const hasQuestionMark = /\?/.test(url);
      delimiter = hasQuestionMark ? '&' : '?';
      newString = url.replace(hash, '') + delimiter + param + '=' + encodeURIComponent(value) + hash;
    } else {
      delimiter = match[0].charAt(0);
      newString = url.replace(re, delimiter + param + '=' + encodeURIComponent(value));
    }

    return newString;
  }

  /**
   * Gets the current query string parameters
   * @returns query string parameters as object
   */
  public static getQueryStringParams(): { [parameter: string]: string } {
    const queryStringParameters: { [parameter: string]: string } = {};
    const urlParams = new URLSearchParams(window.location.search);
    urlParams.forEach((value, key) => {
      queryStringParameters[key] = value;
    });

    return queryStringParameters;
  }

  /**
   * Decodes a provided string
   * @param encodedStr the string to decode
   */
  public static decode(encodedStr: string) {
    const domParser = new DOMParser();
    const htmlContent: Document = domParser.parseFromString(`<!doctype html><body>${encodedStr}</body>`, 'text/html');
    return htmlContent.body.textContent;
  }
}

export enum PageOpenBehavior {
  'Self' = 'self',
  'NewTab' = 'newTab'
}

export enum QueryPathBehavior {
  'URLFragment' = 'fragment',
  'QueryParameter' = 'queryparameter'
}
