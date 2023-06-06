import { ObjectHelper } from '../helpers/ObjectHelper';
import { StringHelper } from '../helpers/StringHelper';
import { ITokenService } from './ITokenService';

export enum BuiltinTokenNames {
  /**
   * The input query text configured in the search results Web Part
   */
  inputQueryText = 'inputQueryText',

  /**
   * Similar as 'inputQueryText' to match the SharePoint search token
   */
  searchTerms = 'searchTerms'
}

export class TokenService implements ITokenService {
  /**
   * This regex only matches expressions enclosed with single, not escaped, curly braces '{}'
   */
  private genericTokenRegexp = /{[^{]+?[^\\]}/gi;

  /**
   * The list of static tokens values set by the Web Part as context
   */
  private tokenValuesList: { [key: string]: string } = {
    [BuiltinTokenNames.inputQueryText]: undefined,
    [BuiltinTokenNames.searchTerms]: undefined
  };

  public setTokenValue(token: string, value: string) {
    // Check if the token is in the whitelist
    if (Object.keys(this.tokenValuesList).indexOf(token) !== -1) {
      this.tokenValuesList[token] = value;
    } else {
      console.log(`The token '${token}' not allowed.`);
    }
  }

  public getTokenValue(token: string): string {
    return this.tokenValuesList[token];
  }

  public resolveTokens(inputString: string): string {
    if (inputString) {
      // Look for static tokens in the specified string
      const tokens = inputString.match(this.genericTokenRegexp);

      if (tokens !== null && tokens.length > 0) {
        tokens.forEach(token => {
          // Take the expression inside curly brackets
          const tokenName = token.substr(1).slice(0, -1);

          // Check if the property exists in the object
          // 'undefined' => token hasn't been initialized in the TokenService instance. We left the token expression untouched (ex: {token}). Ex: no filters component connected, etc.
          // 'null' => token has been initialized but set with a null value. We replace by an empty string as we don't want the string 'null' litterally in the output.
          // '' (empty string) => replaced in the original string with an empty string as well.
          const tokenValue = ObjectHelper.byPath(this.tokenValuesList, tokenName);

          if (tokenValue !== undefined) {
            if (tokenValue !== null) {
              inputString = inputString.replace(new RegExp(StringHelper.escapeRegExp(token), 'gi'), tokenValue);
            } else {
              // If the property value is 'null', we replace by an empty string. 'null' means it has been already set but resolved as empty.
              inputString = inputString.replace(new RegExp(StringHelper.escapeRegExp(token), 'gi'), '');
            }
          }
        });
      }

      // Replace manually escaped curly braces
      inputString = inputString.replace(/\\({|})/gi, '$1');
    }

    return inputString;
  }
}
