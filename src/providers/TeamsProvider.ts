import { IProvider, ProviderState } from './IProvider';
import { Graph } from '../Graph';
import * as microsoftTeams from '@microsoft/teams-js';

declare global {
  interface Window {
    msCrypto: any;
  }
}

export class TeamsProvider extends IProvider {
  private _idToken: string;

  private _clientId: string;
  private _loginPopupUrl: string;

  private _provider: any;

  get provider() {
    return this._provider;
  }

  get isLoggedIn(): boolean {
    return !!this._idToken;
  }

  get clientId(): string {
    return this._clientId;
  }

  set clientId(value: string) {
    this._clientId = value;
  }

  get loginPopupUrl(): string {
    return this._loginPopupUrl;
  }

  set loginPopupUrl(value: string) {
    this._loginPopupUrl = value;
  }

  static async isAvailable(): Promise<boolean> {
    const ms = 500;
    return Promise.race([
      new Promise<boolean>((resolve, reject) => {
        try {
          microsoftTeams.initialize();
          microsoftTeams.getContext(function(context) {
            if (context) {
              resolve(true);
            } else {
              resolve(false);
            }
          });
        } catch (reason) {
          resolve(false);
        }
      }),
      new Promise<boolean>((resolve, reject) => {
        let id = setTimeout(() => {
          clearTimeout(id);
          resolve(false);
        }, ms);
      })
    ]);
  }

  static auth() {
    microsoftTeams.initialize(); // Get the tab context, and use the information to navigate to Azure AD login page

    var url = new URL(window.location.href);
    if (url.searchParams.get('clientId')) {
      microsoftTeams.getContext(function(context) {
        // Generate random state string and store it, so we can verify it in the callback
        var state = _guid();

        localStorage.setItem('simple.state', state);
        localStorage.removeItem('simple.error'); // See https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols-implicit
        // for documentation on these query parameters

        var url = new URL(window.location.href);
        var clientId = url.searchParams.get('clientId');
        url.searchParams.delete('clientId');
        var queryParams = {
          client_id: clientId,
          response_type: 'id_token token',
          response_mode: 'fragment',
          resource: 'https://graph.microsoft.com/',
          redirect_uri: url.href,
          nonce: _guid(),
          state: state,
          // login_hint pre-fills the username/email address field of the sign in page for the user,
          // if you know their username ahead of time.
          login_hint: context.upn
        }; // Go to the AzureAD authorization endpoint

        var authorizeEndpoint =
          'https://login.microsoftonline.com/common/oauth2/authorize?' + toQueryString(queryParams);
        window.location.assign(authorizeEndpoint);
      }); // Build query string from map of query parameter

      function toQueryString(queryParams) {
        var encodedQueryParams = [];

        for (var key in queryParams) {
          encodedQueryParams.push(key + '=' + encodeURIComponent(queryParams[key]));
        }

        return encodedQueryParams.join('&');
      } // Converts decimal to hex equivalent
      // (From ADAL.js: https://github.com/AzureAD/azure-activedirectory-library-for-js/blob/dev/lib/adal.js)

      function _decimalToHex(number) {
        var hex = number.toString(16);

        while (hex.length < 2) {
          hex = '0' + hex;
        }

        return hex;
      } // Generates RFC4122 version 4 guid (128 bits)
      // (From ADAL.js: https://github.com/AzureAD/azure-activedirectory-library-for-js/blob/dev/lib/adal.js)

      function _guid() {
        // RFC4122: The version 4 UUID is meant for generating UUIDs from truly-random or
        // pseudo-random numbers.
        // The algorithm is as follows:
        //     Set the two most significant bits (bits 6 and 7) of the
        //        clock_seq_hi_and_reserved to zero and one, respectively.
        //     Set the four most significant bits (bits 12 through 15) of the
        //        time_hi_and_version field to the 4-bit version number from
        //        Section 4.1.3. Version4
        //     Set all the other bits to randomly (or pseudo-randomly) chosen
        //     values.
        // UUID                   = time-low "-" time-mid "-"time-high-and-version "-"clock-seq-reserved and low(2hexOctet)"-" node
        // time-low               = 4hexOctet
        // time-mid               = 2hexOctet
        // time-high-and-version  = 2hexOctet
        // clock-seq-and-reserved = hexOctet:
        // clock-seq-low          = hexOctet
        // node                   = 6hexOctet
        // Format: xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx
        // y could be 1000, 1001, 1010, 1011 since most significant two bits needs to be 10
        // y values are 8, 9, A, B
        var cryptoObj = window.crypto || window.msCrypto; // for IE 11

        if (cryptoObj && cryptoObj.getRandomValues) {
          var buffer = new Uint8Array(16);
          cryptoObj.getRandomValues(buffer); //buffer[6] and buffer[7] represents the time_hi_and_version field. We will set the four most significant bits (4 through 7) of buffer[6] to represent decimal number 4 (UUID version number).

          buffer[6] |= 0x40; //buffer[6] | 01000000 will set the 6 bit to 1.

          buffer[6] &= 0x4f; //buffer[6] & 01001111 will set the 4, 5, and 7 bit to 0 such that bits 4-7 == 0100 = "4".
          //buffer[8] represents the clock_seq_hi_and_reserved field. We will set the two most significant bits (6 and 7) of the clock_seq_hi_and_reserved to zero and one, respectively.

          buffer[8] |= 0x80; //buffer[8] | 10000000 will set the 7 bit to 1.

          buffer[8] &= 0xbf; //buffer[8] & 10111111 will set the 6 bit to 0.

          return (
            _decimalToHex(buffer[0]) +
            _decimalToHex(buffer[1]) +
            _decimalToHex(buffer[2]) +
            _decimalToHex(buffer[3]) +
            '-' +
            _decimalToHex(buffer[4]) +
            _decimalToHex(buffer[5]) +
            '-' +
            _decimalToHex(buffer[6]) +
            _decimalToHex(buffer[7]) +
            '-' +
            _decimalToHex(buffer[8]) +
            _decimalToHex(buffer[9]) +
            '-' +
            _decimalToHex(buffer[10]) +
            _decimalToHex(buffer[11]) +
            _decimalToHex(buffer[12]) +
            _decimalToHex(buffer[13]) +
            _decimalToHex(buffer[14]) +
            _decimalToHex(buffer[15])
          );
        } else {
          var guidHolder = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx';
          var hex = '0123456789abcdef';
          var r = 0;
          var guidResponse = '';

          for (var i = 0; i < 36; i++) {
            if (guidHolder[i] !== '-' && guidHolder[i] !== '4') {
              // each x and y needs to be random
              r = (Math.random() * 16) | 0;
            }

            if (guidHolder[i] === 'x') {
              guidResponse += hex[r];
            } else if (guidHolder[i] === 'y') {
              // clock-seq-and-reserved first hex is filtered and remaining hex values are random
              r &= 0x3; // bit and with 0011 to set pos 2 to zero ?0??

              r |= 0x8; // set pos 3 to 1 as 1???

              guidResponse += hex[r];
            } else {
              guidResponse += guidHolder[i];
            }
          }

          return guidResponse;
        }
      }
    } else {
      localStorage.removeItem('simple.error'); // Split the key-value pairs passed from Azure AD
      // getHashParameters is a helper function that parses the arguments sent
      // to the callback URL by Azure AD after the authorization call

      var hashParams = getHashParameters();

      if (hashParams['error']) {
        // Authentication/authorization failed
        localStorage.setItem('simple.error', JSON.stringify(hashParams));
        microsoftTeams.authentication.notifyFailure(hashParams['error']);
      } else if (hashParams['access_token']) {
        // Get the stored state parameter and compare with incoming state
        // This validates that the data is coming from Azure AD
        var expectedState = localStorage.getItem('simple.state');

        if (expectedState !== hashParams['state']) {
          // State does not match, report error
          localStorage.setItem('simple.error', JSON.stringify(hashParams));
          microsoftTeams.authentication.notifyFailure('StateDoesNotMatch');
        } else {
          // Success: return token information to the tab
          //microsoftTeams.authentication.notifySuccess({
          //  idToken: hashParams["id_token"],
          //  accessToken: hashParams["access_token"],
          //  tokenType: hashParams["token_type"],
          //  expiresIn: hashParams["expires_in"]
          //});
          microsoftTeams.authentication.notifySuccess(hashParams['access_token']);
        }
      } else {
        // Unexpected condition: hash does not contain error or access_token parameter
        localStorage.setItem('simple.error', JSON.stringify(hashParams));
        microsoftTeams.authentication.notifyFailure('UnexpectedFailure');
      } // Parse hash parameters into key-value pairs

      function getHashParameters() {
        var hashParams = {};
        location.hash
          .substr(1)
          .split('&')
          .forEach(function(item) {
            var s = item.split('='),
              k = s[0],
              v = s[1] && decodeURIComponent(s[1]);
            hashParams[k] = v;
          });
        return hashParams;
      }
    }
  }

  scopes: string[];
  authority: string;

  constructor(clientId: string, loginPopupUrl: string) {
    super();
    this._clientId = clientId;
    this._loginPopupUrl = loginPopupUrl;

    this._provider = this;
    this.graph = new Graph(this);
  }

  async login(): Promise<void> {
    this._idToken = await this.getAccessToken();
    if (this._idToken) {
      this.setState(this._idToken ? ProviderState.SignedIn : ProviderState.SignedOut);
    }
  }

  async getAccessToken(): Promise<string> {
    if (this._idToken) return this._idToken;
    return new Promise((resolve, reject) => {
      var url = new URL(this._loginPopupUrl, new URL(window.location.href));
      url.searchParams.append('clientId', this._clientId);

      microsoftTeams.authentication.authenticate({
        url: url.href,
        width: 600,
        height: 535,
        successCallback: function(result: any) {
          resolve(result);
        },
        failureCallback: function(reason) {
          console.log(reason);
          reject(reason);
        }
      });
    });
  }

  updateScopes(scopes: string[]) {
    this.scopes = scopes;
  }
}
