import { initialize, getContext, authentication } from '@microsoft/teams-js';

var LoginType;
(function (LoginType) {
    LoginType[LoginType["Popup"] = 0] = "Popup";
    LoginType[LoginType["Redirect"] = 1] = "Redirect";
})(LoginType || (LoginType = {}));

var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
class Graph {
    constructor(provider) {
        this.rootUrl = 'https://graph.microsoft.com/beta';
        // this.token = token;
        this._provider = provider;
    }
    getJson(resource, scopes) {
        return __awaiter(this, void 0, void 0, function* () {
            let response = yield this.get(resource, scopes);
            if (response) {
                return response.json();
            }
            return null;
        });
    }
    get(resource, scopes) {
        return __awaiter(this, void 0, void 0, function* () {
            if (!resource.startsWith('/')) {
                resource = "/" + resource;
            }
            let token;
            try {
                token = yield this._provider.getAccessToken(...scopes);
            }
            catch (error) {
                console.log(error);
                return null;
            }
            if (!token) {
                return null;
            }
            let response = yield fetch(this.rootUrl + resource, {
                headers: {
                    authorization: 'Bearer ' + token
                }
            });
            if (response.status >= 400) {
                // hit limit - need to wait and retry per:
                // https://docs.microsoft.com/en-us/graph/throttling
                if (response.status == 429) {
                    console.log('too many requests - wait ' + response.headers.get('Retry-After') + ' seconds');
                    return null;
                }
                let error = response.json();
                if (error.error !== undefined) {
                    console.log(error);
                }
                console.log(response);
                throw 'error accessing graph';
            }
            return response;
        });
    }
    me() {
        return __awaiter(this, void 0, void 0, function* () {
            let scopes = ['user.read'];
            return this.getJson('/me', scopes);
        });
    }
    getUser(userPrincipleName) {
        return __awaiter(this, void 0, void 0, function* () {
            let scopes = ['user.readbasic.all'];
            return this.getJson(`/users/${userPrincipleName}`, scopes);
        });
    }
    findPerson(query) {
        return __awaiter(this, void 0, void 0, function* () {
            let scopes = ['people.read'];
            let result = yield this.getJson(`/me/people/?$search="${query}"`, scopes);
            return result ? result.value : null;
        });
    }
    myPhoto() {
        let scopes = ['user.read'];
        return this.getBase64('/me/photo/$value', scopes);
    }
    getUserPhoto(id) {
        return __awaiter(this, void 0, void 0, function* () {
            let scopes = ['user.readbasic.all'];
            return this.getBase64(`users/${id}/photo/$value`, scopes);
        });
    }
    getBase64(resource, scopes) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                let response = yield this.get(resource, scopes);
                if (!response) {
                    return null;
                }
                let blob = yield response.blob();
                return new Promise((resolve, reject) => {
                    const reader = new FileReader;
                    reader.onerror = reject;
                    reader.onload = _ => {
                        resolve(reader.result);
                    };
                    reader.readAsDataURL(blob);
                });
            }
            catch (_a) {
                return null;
            }
        });
    }
    calendar(startDateTime, endDateTime) {
        return __awaiter(this, void 0, void 0, function* () {
            let scopes = ['calendars.read'];
            let sdt = `startdatetime=${startDateTime.toISOString()}`;
            let edt = `enddatetime=${endDateTime.toISOString()}`;
            let uri = `/me/calendarview?${sdt}&${edt}`;
            let calendar = yield this.getJson(uri, scopes);
            return calendar ? calendar.value : null;
        });
    }
}

class EventDispatcher {
    constructor() {
        this.eventHandlers = [];
    }
    fire(event) {
        for (let handler of this.eventHandlers) {
            handler(event);
        }
    }
    register(eventHandler) {
        this.eventHandlers.push(eventHandler);
    }
}

var __awaiter$1 = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
// import { UserAgentApplication } from 'msal/lib-es6';
class MsalProvider {
    constructor(config) {
        this._loginChangedDispatcher = new EventDispatcher();
        this.temp = 0;
        if (!config.clientId) {
            throw 'ClientID must be a valid string';
        }
        this.initProvider(config);
    }
    get provider() {
        return this._provider;
    }
    get isLoggedIn() {
        return !!this._idToken;
    }
    get isAvailable() {
        return true;
    }
    initProvider(config) {
        return __awaiter$1(this, void 0, void 0, function* () {
            let msal = yield import('./chunk-96b3023a.js');
            console.log('initProvider');
            this._clientId = config.clientId;
            this.scopes =
                typeof config.scopes !== 'undefined' ? config.scopes : ['user.read'];
            this.authority =
                typeof config.authority !== 'undefined' ? config.authority : null;
            let options = typeof config.options != 'undefined'
                ? config.options
                : { cacheLocation: 'localStorage' };
            this._loginType =
                typeof config.loginType !== 'undefined'
                    ? config.loginType
                    : LoginType.Redirect;
            let callbackFunction = ((errorDesc, token, error, tokenType, state) => {
                this.tokenReceivedCallback(errorDesc, token, error, tokenType, state);
            }).bind(this);
            this._provider = new msal.UserAgentApplication(this._clientId, this.authority, callbackFunction, options);
            console.log(this._provider);
            this.graph = new Graph(this);
            this.tryGetIdTokenSilent();
        });
    }
    login() {
        return __awaiter$1(this, void 0, void 0, function* () {
            console.log('login');
            if (this._loginType == LoginType.Popup) {
                this._idToken = yield this.provider.loginPopup(this.scopes);
                this.fireLoginChangedEvent({});
            }
            else {
                this.provider.loginRedirect(this.scopes);
            }
        });
    }
    tryGetIdTokenSilent() {
        return __awaiter$1(this, void 0, void 0, function* () {
            console.log('tryGetIdTokenSilent');
            try {
                this._idToken = yield this.provider.acquireTokenSilent([this._clientId], this.authority);
                if (this._idToken) {
                    console.log('tryGetIdTokenSilent: got a token');
                    this.fireLoginChangedEvent({});
                }
                return this.isLoggedIn;
            }
            catch (e) {
                console.log(e);
                return false;
            }
        });
    }
    getAccessToken(...scopes) {
        return __awaiter$1(this, void 0, void 0, function* () {
            ++this.temp;
            let temp = this.temp;
            scopes = scopes || this.scopes;
            console.log('getaccesstoken' + ++temp + ': scopes' + scopes);
            let accessToken;
            try {
                accessToken = yield this.provider.acquireTokenSilent(scopes, this.authority);
                console.log('getaccesstoken' + temp + ': got token');
            }
            catch (e) {
                try {
                    console.log('getaccesstoken' + temp + ': catch ' + e);
                    // TODO - figure out for what error this logic is needed so we
                    // don't prompt the user to login unnecessarily
                    if (e.includes('multiple_matching_tokens_detected')) {
                        console.log('getaccesstoken' + temp + ' ' + e);
                        return null;
                    }
                    if (this._loginType == LoginType.Redirect) {
                        this.provider.acquireTokenRedirect(scopes);
                        return new Promise((resolve, reject) => {
                            this._resolveToken = resolve;
                            this._rejectToken = reject;
                        });
                    }
                    else {
                        accessToken = yield this.provider.acquireTokenPopup(scopes);
                    }
                }
                catch (e) {
                    // TODO - figure out how to expose this during dev to make it easy for the dev to figure out
                    // if error contains "'token' is not enabled", make sure to have implicit oAuth enabled in the AAD manifest
                    console.log('getaccesstoken' + temp + 'catch2: ' + e);
                    throw e;
                }
            }
            return accessToken;
        });
    }
    logout() {
        return __awaiter$1(this, void 0, void 0, function* () {
            this.provider.logout();
            this.fireLoginChangedEvent({});
        });
    }
    updateScopes(scopes) {
        this.scopes = scopes;
    }
    tokenReceivedCallback(errorDesc, token, error, tokenType, state) {
        // debugger;
        console.log('tokenReceivedCallback ' + errorDesc + ' | ' + tokenType);
        if (this._provider) {
            console.log(window.location.hash);
            console.log('isCallback: ' + this._provider.isCallback(window.location.hash));
        }
        if (error) {
            console.log(error + ' ' + errorDesc);
            if (this._rejectToken) {
                this._rejectToken(errorDesc);
            }
        }
        else {
            if (tokenType == 'id_token') {
                this._idToken = token;
                this.fireLoginChangedEvent({});
            }
            else {
                if (this._resolveToken) {
                    this._resolveToken(token);
                }
            }
        }
    }
    onLoginChanged(eventHandler) {
        console.log('onloginChanged');
        this._loginChangedDispatcher.register(eventHandler);
    }
    fireLoginChangedEvent(event) {
        console.log('fireLoginChangedEvent');
        this._loginChangedDispatcher.fire(event);
    }
}

var __awaiter$2 = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
class WamProvider {
    constructor(clientId, authority) {
        this.graphResource = 'https://graph.microsoft.com';
        this._loginChangedDispatcher = new EventDispatcher();
        this.clientId = clientId;
        this.authority = authority || 'https://login.microsoftonline.com/common';
        this.graph = new Graph(this);
        this.printRedirectUriToConsole();
    }
    static isAvailable() {
        return !!(window.Windows);
    }
    get isLoggedIn() {
        return !!this.accessToken;
    }
    login() {
        return __awaiter$2(this, void 0, void 0, function* () {
            if (WamProvider.isAvailable()) {
                let webCore = window.Windows.Security.Authentication.Web.Core;
                let wap = yield webCore.WebAuthenticationCoreManager.findAccountProviderAsync('https://login.microsoft.com', this.authority);
                if (!wap) {
                    console.log('no account provider');
                    return;
                }
                let wtr = new webCore.WebTokenRequest(wap, '', this.clientId);
                wtr.properties.insert('resource', this.graphResource);
                let wtrr = yield webCore.WebAuthenticationCoreManager.requestTokenAsync(wtr);
                switch (wtrr.responseStatus) {
                    case webCore.WebTokenRequestStatus.success:
                        let account = wtrr.responseData[0].webAccount;
                        this.accessToken = wtrr.responseData[0].token;
                        this.fireLoginChangedEvent({});
                        break;
                    case webCore.WebTokenRequestStatus.userCancel:
                    case webCore.WebTokenRequestStatus.accountSwitch:
                    case webCore.WebTokenRequestStatus.userInteractionRequired:
                    case webCore.WebTokenRequestStatus.accountProviderNotAvailable:
                    case webCore.WebTokenRequestStatus.providerError:
                        console.log(`status ${wtrr.responseStatus}: error code ${wtrr.responseError} | error message ${wtrr.responseError.errorMessage}`);
                        break;
                }
            }
        });
    }
    printRedirectUriToConsole() {
        if (WamProvider.isAvailable()) {
            let web = window.Windows.Security.Authentication.Web;
            let redirectUri = `ms-appx-web://Microsoft.AAD.BrokerPlugIn/${web.WebAuthenticationBroker.getCurrentApplicationCallbackUri().host.toUpperCase()}`;
            console.log("Use the following redirect URI in your AAD application:");
            console.log(redirectUri);
        }
        else {
            console.log("WAM not supported on this platform");
        }
    }
    logout() {
        throw new Error("Method not implemented.");
    }
    getAccessToken(...scopes) {
        if (this.isLoggedIn) {
            return Promise.resolve(this.accessToken);
        }
        else {
            throw "Not logged in";
        }
    }
    onLoginChanged(eventHandler) {
        this._loginChangedDispatcher.register(eventHandler);
    }
    fireLoginChangedEvent(event) {
        this._loginChangedDispatcher.fire(event);
    }
}

var __awaiter$3 = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
class SharePointProvider {
    constructor(context) {
        this._loginChangedDispatcher = new EventDispatcher();
        this.context = context;
        context.aadTokenProviderFactory.getTokenProvider().then((tokenProvider) => {
            this._provider = tokenProvider;
            this.graph = new Graph(this);
            this.internalLogin();
        });
        this.fireLoginChangedEvent({});
    }
    get provider() {
        return this._provider;
    }
    ;
    get isLoggedIn() {
        return !!this._idToken;
    }
    ;
    internalLogin() {
        return __awaiter$3(this, void 0, void 0, function* () {
            this._idToken = yield this.getAccessToken();
            if (this._idToken) {
                this.fireLoginChangedEvent({});
            }
        });
    }
    getAccessToken() {
        return __awaiter$3(this, void 0, void 0, function* () {
            let accessToken;
            try {
                accessToken = yield this.provider.getToken("https://graph.microsoft.com");
            }
            catch (e) {
                console.log(e);
                throw e;
            }
            return accessToken;
        });
    }
    updateScopes(scopes) {
        this.scopes = scopes;
    }
    onLoginChanged(eventHandler) {
        this._loginChangedDispatcher.register(eventHandler);
    }
    fireLoginChangedEvent(event) {
        this._loginChangedDispatcher.fire(event);
    }
}

var __awaiter$4 = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
class TeamsProvider {
    constructor(clientId, loginPopupUrl) {
        this._loginChangedDispatcher = new EventDispatcher();
        this._clientId = clientId;
        this._loginPopupUrl = loginPopupUrl;
        this._provider = this;
        this.graph = new Graph(this);
    }
    get provider() {
        return this._provider;
    }
    ;
    get isLoggedIn() {
        return !!this._idToken;
    }
    ;
    get clientId() {
        return this._clientId;
    }
    set clientId(value) {
        this._clientId = value;
    }
    get loginPopupUrl() {
        return this._loginPopupUrl;
    }
    set loginPopupUrl(value) {
        this._loginPopupUrl = value;
    }
    static isAvailable() {
        return __awaiter$4(this, void 0, void 0, function* () {
            const ms = 500;
            return Promise.race([
                new Promise((resolve, reject) => {
                    try {
                        initialize();
                        getContext(function (context) {
                            if (context) {
                                resolve(true);
                            }
                            else {
                                resolve(false);
                            }
                        });
                    }
                    catch (reason) {
                        resolve(false);
                    }
                }),
                new Promise((resolve, reject) => {
                    let id = setTimeout(() => {
                        clearTimeout(id);
                        resolve(false);
                    }, ms);
                })
            ]);
        });
    }
    ;
    static auth() {
        initialize(); // Get the tab context, and use the information to navigate to Azure AD login page
        var url = new URL(window.location.href);
        if (url.searchParams.get('clientId')) {
            getContext(function (context) {
                // Generate random state string and store it, so we can verify it in the callback
                var state = _guid();
                localStorage.setItem("simple.state", state);
                localStorage.removeItem("simple.error"); // See https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols-implicit
                // for documentation on these query parameters
                var url = new URL(window.location.href);
                var clientId = url.searchParams.get("clientId");
                url.searchParams.delete("clientId");
                var queryParams = {
                    client_id: clientId,
                    response_type: "id_token token",
                    response_mode: "fragment",
                    resource: "https://graph.microsoft.com/",
                    redirect_uri: url.href,
                    nonce: _guid(),
                    state: state,
                    // login_hint pre-fills the username/email address field of the sign in page for the user, 
                    // if you know their username ahead of time.
                    login_hint: context.upn
                }; // Go to the AzureAD authorization endpoint
                var authorizeEndpoint = "https://login.microsoftonline.com/common/oauth2/authorize?" + toQueryString(queryParams);
                window.location.assign(authorizeEndpoint);
            }); // Build query string from map of query parameter
            function toQueryString(queryParams) {
                var encodedQueryParams = [];
                for (var key in queryParams) {
                    encodedQueryParams.push(key + "=" + encodeURIComponent(queryParams[key]));
                }
                return encodedQueryParams.join("&");
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
                    return _decimalToHex(buffer[0]) + _decimalToHex(buffer[1]) + _decimalToHex(buffer[2]) + _decimalToHex(buffer[3]) + '-' + _decimalToHex(buffer[4]) + _decimalToHex(buffer[5]) + '-' + _decimalToHex(buffer[6]) + _decimalToHex(buffer[7]) + '-' + _decimalToHex(buffer[8]) + _decimalToHex(buffer[9]) + '-' + _decimalToHex(buffer[10]) + _decimalToHex(buffer[11]) + _decimalToHex(buffer[12]) + _decimalToHex(buffer[13]) + _decimalToHex(buffer[14]) + _decimalToHex(buffer[15]);
                }
                else {
                    var guidHolder = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx';
                    var hex = '0123456789abcdef';
                    var r = 0;
                    var guidResponse = "";
                    for (var i = 0; i < 36; i++) {
                        if (guidHolder[i] !== '-' && guidHolder[i] !== '4') {
                            // each x and y needs to be random
                            r = Math.random() * 16 | 0;
                        }
                        if (guidHolder[i] === 'x') {
                            guidResponse += hex[r];
                        }
                        else if (guidHolder[i] === 'y') {
                            // clock-seq-and-reserved first hex is filtered and remaining hex values are random
                            r &= 0x3; // bit and with 0011 to set pos 2 to zero ?0??
                            r |= 0x8; // set pos 3 to 1 as 1???
                            guidResponse += hex[r];
                        }
                        else {
                            guidResponse += guidHolder[i];
                        }
                    }
                    return guidResponse;
                }
            }
        }
        else {
            localStorage.removeItem("simple.error"); // Split the key-value pairs passed from Azure AD
            // getHashParameters is a helper function that parses the arguments sent
            // to the callback URL by Azure AD after the authorization call
            var hashParams = getHashParameters();
            if (hashParams["error"]) {
                // Authentication/authorization failed
                localStorage.setItem("simple.error", JSON.stringify(hashParams));
                authentication.notifyFailure(hashParams["error"]);
            }
            else if (hashParams["access_token"]) {
                // Get the stored state parameter and compare with incoming state
                // This validates that the data is coming from Azure AD
                var expectedState = localStorage.getItem("simple.state");
                if (expectedState !== hashParams["state"]) {
                    // State does not match, report error
                    localStorage.setItem("simple.error", JSON.stringify(hashParams));
                    authentication.notifyFailure("StateDoesNotMatch");
                }
                else {
                    // Success: return token information to the tab
                    //microsoftTeams.authentication.notifySuccess({
                    //  idToken: hashParams["id_token"],
                    //  accessToken: hashParams["access_token"],
                    //  tokenType: hashParams["token_type"],
                    //  expiresIn: hashParams["expires_in"]
                    //});
                    authentication.notifySuccess(hashParams["access_token"]);
                }
            }
            else {
                // Unexpected condition: hash does not contain error or access_token parameter
                localStorage.setItem("simple.error", JSON.stringify(hashParams));
                authentication.notifyFailure("UnexpectedFailure");
            } // Parse hash parameters into key-value pairs
            function getHashParameters() {
                var hashParams = {};
                location.hash.substr(1).split("&").forEach(function (item) {
                    var s = item.split("="), k = s[0], v = s[1] && decodeURIComponent(s[1]);
                    hashParams[k] = v;
                });
                return hashParams;
            }
        }
    }
    login() {
        return __awaiter$4(this, void 0, void 0, function* () {
            this._idToken = yield this.getAccessToken();
            if (this._idToken) {
                this.fireLoginChangedEvent({});
            }
        });
    }
    getAccessToken() {
        return __awaiter$4(this, void 0, void 0, function* () {
            if (this._idToken)
                return this._idToken;
            return new Promise((resolve, reject) => {
                var url = new URL(this._loginPopupUrl, new URL(window.location.href));
                url.searchParams.append('clientId', this._clientId);
                authentication.authenticate({
                    url: url.href,
                    width: 600,
                    height: 535,
                    successCallback: function (result) {
                        resolve(result);
                    },
                    failureCallback: function (reason) {
                        console.log(reason);
                        reject(reason);
                    }
                });
            });
        });
    }
    updateScopes(scopes) {
        this.scopes = scopes;
    }
    onLoginChanged(eventHandler) {
        this._loginChangedDispatcher.register(eventHandler);
    }
    fireLoginChangedEvent(event) {
        this._loginChangedDispatcher.fire(event);
    }
}

var __awaiter$5 = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var Providers;
(function (Providers) {
    function getAvailable() {
        const providers = getProviders();
        for (let provider of providers) {
            return provider;
        }
        return null;
    }
    Providers.getAvailable = getAvailable;
    function addCustomProvider(provider) {
        const providers = getProviders();
        if (provider !== null) {
            providers.push(provider);
            getEventDispatcher().fire({});
        }
        return provider;
    }
    Providers.addCustomProvider = addCustomProvider;
    function addWamProvider(clientId, authority) {
        if (WamProvider.isAvailable()) {
            return addCustomProvider(new WamProvider(clientId, authority));
        }
        return null;
    }
    Providers.addWamProvider = addWamProvider;
    function addMsalProvider(config) {
        return addCustomProvider(new MsalProvider(config));
    }
    Providers.addMsalProvider = addMsalProvider;
    function addSharePointProvider(context) {
        return addCustomProvider(new SharePointProvider(context));
    }
    Providers.addSharePointProvider = addSharePointProvider;
    function addTeamsProvider(clientId, loginPopupUrl) {
        return __awaiter$5(this, void 0, void 0, function* () {
            if (yield TeamsProvider.isAvailable()) {
                return addCustomProvider(new TeamsProvider(clientId, loginPopupUrl));
            }
            return null;
        });
    }
    Providers.addTeamsProvider = addTeamsProvider;
    function onProvidersChanged(event) {
        getEventDispatcher().register(event);
    }
    Providers.onProvidersChanged = onProvidersChanged;
    // TODO - figure out a better way to have a global reference to all providers
    function getProviders() {
        if (!window._msgraph_providers) {
            window._msgraph_providers = [];
        }
        return window._msgraph_providers;
    }
    function getEventDispatcher() {
        if (!window._msgraph_eventDispatcher) {
            window._msgraph_eventDispatcher = new EventDispatcher();
        }
        return window._msgraph_eventDispatcher;
    }
})(Providers || (Providers = {}));

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
const directives = new WeakMap();
const isDirective = (o) => {
    return typeof o === 'function' && directives.has(o);
};

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
/**
 * True if the custom elements polyfill is in use.
 */
const isCEPolyfill = window.customElements !== undefined &&
    window.customElements.polyfillWrapFlushCallback !==
        undefined;
/**
 * Removes nodes, starting from `startNode` (inclusive) to `endNode`
 * (exclusive), from `container`.
 */
const removeNodes = (container, startNode, endNode = null) => {
    let node = startNode;
    while (node !== endNode) {
        const n = node.nextSibling;
        container.removeChild(node);
        node = n;
    }
};

/**
 * @license
 * Copyright (c) 2018 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
/**
 * A sentinel value that signals that a value was handled by a directive and
 * should not be written to the DOM.
 */
const noChange = {};
/**
 * A sentinel value that signals a NodePart to fully clear its content.
 */
const nothing = {};

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
/**
 * An expression marker with embedded unique key to avoid collision with
 * possible text in templates.
 */
const marker = `{{lit-${String(Math.random()).slice(2)}}}`;
/**
 * An expression marker used text-positions, multi-binding attributes, and
 * attributes with markup-like text values.
 */
const nodeMarker = `<!--${marker}-->`;
const markerRegex = new RegExp(`${marker}|${nodeMarker}`);
/**
 * Suffix appended to all bound attribute names.
 */
const boundAttributeSuffix = '$lit$';
/**
 * An updateable Template that tracks the location of dynamic parts.
 */
class Template {
    constructor(result, element) {
        this.parts = [];
        this.element = element;
        let index = -1;
        let partIndex = 0;
        const nodesToRemove = [];
        const _prepareTemplate = (template) => {
            const content = template.content;
            // Edge needs all 4 parameters present; IE11 needs 3rd parameter to be
            // null
            const walker = document.createTreeWalker(content, 133 /* NodeFilter.SHOW_{ELEMENT|COMMENT|TEXT} */, null, false);
            // Keeps track of the last index associated with a part. We try to delete
            // unnecessary nodes, but we never want to associate two different parts
            // to the same index. They must have a constant node between.
            let lastPartIndex = 0;
            while (walker.nextNode()) {
                index++;
                const node = walker.currentNode;
                if (node.nodeType === 1 /* Node.ELEMENT_NODE */) {
                    if (node.hasAttributes()) {
                        const attributes = node.attributes;
                        // Per
                        // https://developer.mozilla.org/en-US/docs/Web/API/NamedNodeMap,
                        // attributes are not guaranteed to be returned in document order.
                        // In particular, Edge/IE can return them out of order, so we cannot
                        // assume a correspondance between part index and attribute index.
                        let count = 0;
                        for (let i = 0; i < attributes.length; i++) {
                            if (attributes[i].value.indexOf(marker) >= 0) {
                                count++;
                            }
                        }
                        while (count-- > 0) {
                            // Get the template literal section leading up to the first
                            // expression in this attribute
                            const stringForPart = result.strings[partIndex];
                            // Find the attribute name
                            const name = lastAttributeNameRegex.exec(stringForPart)[2];
                            // Find the corresponding attribute
                            // All bound attributes have had a suffix added in
                            // TemplateResult#getHTML to opt out of special attribute
                            // handling. To look up the attribute value we also need to add
                            // the suffix.
                            const attributeLookupName = name.toLowerCase() + boundAttributeSuffix;
                            const attributeValue = node.getAttribute(attributeLookupName);
                            const strings = attributeValue.split(markerRegex);
                            this.parts.push({ type: 'attribute', index, name, strings });
                            node.removeAttribute(attributeLookupName);
                            partIndex += strings.length - 1;
                        }
                    }
                    if (node.tagName === 'TEMPLATE') {
                        _prepareTemplate(node);
                    }
                }
                else if (node.nodeType === 3 /* Node.TEXT_NODE */) {
                    const data = node.data;
                    if (data.indexOf(marker) >= 0) {
                        const parent = node.parentNode;
                        const strings = data.split(markerRegex);
                        const lastIndex = strings.length - 1;
                        // Generate a new text node for each literal section
                        // These nodes are also used as the markers for node parts
                        for (let i = 0; i < lastIndex; i++) {
                            parent.insertBefore((strings[i] === '') ? createMarker() :
                                document.createTextNode(strings[i]), node);
                            this.parts.push({ type: 'node', index: ++index });
                        }
                        // If there's no text, we must insert a comment to mark our place.
                        // Else, we can trust it will stick around after cloning.
                        if (strings[lastIndex] === '') {
                            parent.insertBefore(createMarker(), node);
                            nodesToRemove.push(node);
                        }
                        else {
                            node.data = strings[lastIndex];
                        }
                        // We have a part for each match found
                        partIndex += lastIndex;
                    }
                }
                else if (node.nodeType === 8 /* Node.COMMENT_NODE */) {
                    if (node.data === marker) {
                        const parent = node.parentNode;
                        // Add a new marker node to be the startNode of the Part if any of
                        // the following are true:
                        //  * We don't have a previousSibling
                        //  * The previousSibling is already the start of a previous part
                        if (node.previousSibling === null || index === lastPartIndex) {
                            index++;
                            parent.insertBefore(createMarker(), node);
                        }
                        lastPartIndex = index;
                        this.parts.push({ type: 'node', index });
                        // If we don't have a nextSibling, keep this node so we have an end.
                        // Else, we can remove it to save future costs.
                        if (node.nextSibling === null) {
                            node.data = '';
                        }
                        else {
                            nodesToRemove.push(node);
                            index--;
                        }
                        partIndex++;
                    }
                    else {
                        let i = -1;
                        while ((i = node.data.indexOf(marker, i + 1)) !==
                            -1) {
                            // Comment node has a binding marker inside, make an inactive part
                            // The binding won't work, but subsequent bindings will
                            // TODO (justinfagnani): consider whether it's even worth it to
                            // make bindings in comments work
                            this.parts.push({ type: 'node', index: -1 });
                        }
                    }
                }
            }
        };
        _prepareTemplate(element);
        // Remove text binding nodes after the walk to not disturb the TreeWalker
        for (const n of nodesToRemove) {
            n.parentNode.removeChild(n);
        }
    }
}
const isTemplatePartActive = (part) => part.index !== -1;
// Allows `document.createComment('')` to be renamed for a
// small manual size-savings.
const createMarker = () => document.createComment('');
/**
 * This regex extracts the attribute name preceding an attribute-position
 * expression. It does this by matching the syntax allowed for attributes
 * against the string literal directly preceding the expression, assuming that
 * the expression is in an attribute-value position.
 *
 * See attributes in the HTML spec:
 * https://www.w3.org/TR/html5/syntax.html#attributes-0
 *
 * "\0-\x1F\x7F-\x9F" are Unicode control characters
 *
 * " \x09\x0a\x0c\x0d" are HTML space characters:
 * https://www.w3.org/TR/html5/infrastructure.html#space-character
 *
 * So an attribute is:
 *  * The name: any character except a control character, space character, ('),
 *    ("), ">", "=", or "/"
 *  * Followed by zero or more space characters
 *  * Followed by "="
 *  * Followed by zero or more space characters
 *  * Followed by:
 *    * Any character except space, ('), ("), "<", ">", "=", (`), or
 *    * (") then any non-("), or
 *    * (') then any non-(')
 */
const lastAttributeNameRegex = /([ \x09\x0a\x0c\x0d])([^\0-\x1F\x7F-\x9F \x09\x0a\x0c\x0d"'>=/]+)([ \x09\x0a\x0c\x0d]*=[ \x09\x0a\x0c\x0d]*(?:[^ \x09\x0a\x0c\x0d"'`<>=]*|"[^"]*|'[^']*))$/;

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
/**
 * An instance of a `Template` that can be attached to the DOM and updated
 * with new values.
 */
class TemplateInstance {
    constructor(template, processor, options) {
        this._parts = [];
        this.template = template;
        this.processor = processor;
        this.options = options;
    }
    update(values) {
        let i = 0;
        for (const part of this._parts) {
            if (part !== undefined) {
                part.setValue(values[i]);
            }
            i++;
        }
        for (const part of this._parts) {
            if (part !== undefined) {
                part.commit();
            }
        }
    }
    _clone() {
        // When using the Custom Elements polyfill, clone the node, rather than
        // importing it, to keep the fragment in the template's document. This
        // leaves the fragment inert so custom elements won't upgrade and
        // potentially modify their contents by creating a polyfilled ShadowRoot
        // while we traverse the tree.
        const fragment = isCEPolyfill ?
            this.template.element.content.cloneNode(true) :
            document.importNode(this.template.element.content, true);
        const parts = this.template.parts;
        let partIndex = 0;
        let nodeIndex = 0;
        const _prepareInstance = (fragment) => {
            // Edge needs all 4 parameters present; IE11 needs 3rd parameter to be
            // null
            const walker = document.createTreeWalker(fragment, 133 /* NodeFilter.SHOW_{ELEMENT|COMMENT|TEXT} */, null, false);
            let node = walker.nextNode();
            // Loop through all the nodes and parts of a template
            while (partIndex < parts.length && node !== null) {
                const part = parts[partIndex];
                // Consecutive Parts may have the same node index, in the case of
                // multiple bound attributes on an element. So each iteration we either
                // increment the nodeIndex, if we aren't on a node with a part, or the
                // partIndex if we are. By not incrementing the nodeIndex when we find a
                // part, we allow for the next part to be associated with the current
                // node if neccessasry.
                if (!isTemplatePartActive(part)) {
                    this._parts.push(undefined);
                    partIndex++;
                }
                else if (nodeIndex === part.index) {
                    if (part.type === 'node') {
                        const part = this.processor.handleTextExpression(this.options);
                        part.insertAfterNode(node.previousSibling);
                        this._parts.push(part);
                    }
                    else {
                        this._parts.push(...this.processor.handleAttributeExpressions(node, part.name, part.strings, this.options));
                    }
                    partIndex++;
                }
                else {
                    nodeIndex++;
                    if (node.nodeName === 'TEMPLATE') {
                        _prepareInstance(node.content);
                    }
                    node = walker.nextNode();
                }
            }
        };
        _prepareInstance(fragment);
        if (isCEPolyfill) {
            document.adoptNode(fragment);
            customElements.upgrade(fragment);
        }
        return fragment;
    }
}

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
/**
 * The return type of `html`, which holds a Template and the values from
 * interpolated expressions.
 */
class TemplateResult {
    constructor(strings, values, type, processor) {
        this.strings = strings;
        this.values = values;
        this.type = type;
        this.processor = processor;
    }
    /**
     * Returns a string of HTML used to create a `<template>` element.
     */
    getHTML() {
        const endIndex = this.strings.length - 1;
        let html = '';
        for (let i = 0; i < endIndex; i++) {
            const s = this.strings[i];
            // This exec() call does two things:
            // 1) Appends a suffix to the bound attribute name to opt out of special
            // attribute value parsing that IE11 and Edge do, like for style and
            // many SVG attributes. The Template class also appends the same suffix
            // when looking up attributes to create Parts.
            // 2) Adds an unquoted-attribute-safe marker for the first expression in
            // an attribute. Subsequent attribute expressions will use node markers,
            // and this is safe since attributes with multiple expressions are
            // guaranteed to be quoted.
            const match = lastAttributeNameRegex.exec(s);
            if (match) {
                // We're starting a new bound attribute.
                // Add the safe attribute suffix, and use unquoted-attribute-safe
                // marker.
                html += s.substr(0, match.index) + match[1] + match[2] +
                    boundAttributeSuffix + match[3] + marker;
            }
            else {
                // We're either in a bound node, or trailing bound attribute.
                // Either way, nodeMarker is safe to use.
                html += s + nodeMarker;
            }
        }
        return html + this.strings[endIndex];
    }
    getTemplateElement() {
        const template = document.createElement('template');
        template.innerHTML = this.getHTML();
        return template;
    }
}

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
const isPrimitive = (value) => {
    return (value === null ||
        !(typeof value === 'object' || typeof value === 'function'));
};
/**
 * Sets attribute values for AttributeParts, so that the value is only set once
 * even if there are multiple parts for an attribute.
 */
class AttributeCommitter {
    constructor(element, name, strings) {
        this.dirty = true;
        this.element = element;
        this.name = name;
        this.strings = strings;
        this.parts = [];
        for (let i = 0; i < strings.length - 1; i++) {
            this.parts[i] = this._createPart();
        }
    }
    /**
     * Creates a single part. Override this to create a differnt type of part.
     */
    _createPart() {
        return new AttributePart(this);
    }
    _getValue() {
        const strings = this.strings;
        const l = strings.length - 1;
        let text = '';
        for (let i = 0; i < l; i++) {
            text += strings[i];
            const part = this.parts[i];
            if (part !== undefined) {
                const v = part.value;
                if (v != null &&
                    (Array.isArray(v) ||
                        // tslint:disable-next-line:no-any
                        typeof v !== 'string' && v[Symbol.iterator])) {
                    for (const t of v) {
                        text += typeof t === 'string' ? t : String(t);
                    }
                }
                else {
                    text += typeof v === 'string' ? v : String(v);
                }
            }
        }
        text += strings[l];
        return text;
    }
    commit() {
        if (this.dirty) {
            this.dirty = false;
            this.element.setAttribute(this.name, this._getValue());
        }
    }
}
class AttributePart {
    constructor(comitter) {
        this.value = undefined;
        this.committer = comitter;
    }
    setValue(value) {
        if (value !== noChange && (!isPrimitive(value) || value !== this.value)) {
            this.value = value;
            // If the value is a not a directive, dirty the committer so that it'll
            // call setAttribute. If the value is a directive, it'll dirty the
            // committer if it calls setValue().
            if (!isDirective(value)) {
                this.committer.dirty = true;
            }
        }
    }
    commit() {
        while (isDirective(this.value)) {
            const directive = this.value;
            this.value = noChange;
            directive(this);
        }
        if (this.value === noChange) {
            return;
        }
        this.committer.commit();
    }
}
class NodePart {
    constructor(options) {
        this.value = undefined;
        this._pendingValue = undefined;
        this.options = options;
    }
    /**
     * Inserts this part into a container.
     *
     * This part must be empty, as its contents are not automatically moved.
     */
    appendInto(container) {
        this.startNode = container.appendChild(createMarker());
        this.endNode = container.appendChild(createMarker());
    }
    /**
     * Inserts this part between `ref` and `ref`'s next sibling. Both `ref` and
     * its next sibling must be static, unchanging nodes such as those that appear
     * in a literal section of a template.
     *
     * This part must be empty, as its contents are not automatically moved.
     */
    insertAfterNode(ref) {
        this.startNode = ref;
        this.endNode = ref.nextSibling;
    }
    /**
     * Appends this part into a parent part.
     *
     * This part must be empty, as its contents are not automatically moved.
     */
    appendIntoPart(part) {
        part._insert(this.startNode = createMarker());
        part._insert(this.endNode = createMarker());
    }
    /**
     * Appends this part after `ref`
     *
     * This part must be empty, as its contents are not automatically moved.
     */
    insertAfterPart(ref) {
        ref._insert(this.startNode = createMarker());
        this.endNode = ref.endNode;
        ref.endNode = this.startNode;
    }
    setValue(value) {
        this._pendingValue = value;
    }
    commit() {
        while (isDirective(this._pendingValue)) {
            const directive = this._pendingValue;
            this._pendingValue = noChange;
            directive(this);
        }
        const value = this._pendingValue;
        if (value === noChange) {
            return;
        }
        if (isPrimitive(value)) {
            if (value !== this.value) {
                this._commitText(value);
            }
        }
        else if (value instanceof TemplateResult) {
            this._commitTemplateResult(value);
        }
        else if (value instanceof Node) {
            this._commitNode(value);
        }
        else if (Array.isArray(value) ||
            // tslint:disable-next-line:no-any
            value[Symbol.iterator]) {
            this._commitIterable(value);
        }
        else if (value === nothing) {
            this.value = nothing;
            this.clear();
        }
        else {
            // Fallback, will render the string representation
            this._commitText(value);
        }
    }
    _insert(node) {
        this.endNode.parentNode.insertBefore(node, this.endNode);
    }
    _commitNode(value) {
        if (this.value === value) {
            return;
        }
        this.clear();
        this._insert(value);
        this.value = value;
    }
    _commitText(value) {
        const node = this.startNode.nextSibling;
        value = value == null ? '' : value;
        if (node === this.endNode.previousSibling &&
            node.nodeType === 3 /* Node.TEXT_NODE */) {
            // If we only have a single text node between the markers, we can just
            // set its value, rather than replacing it.
            // TODO(justinfagnani): Can we just check if this.value is primitive?
            node.data = value;
        }
        else {
            this._commitNode(document.createTextNode(typeof value === 'string' ? value : String(value)));
        }
        this.value = value;
    }
    _commitTemplateResult(value) {
        const template = this.options.templateFactory(value);
        if (this.value instanceof TemplateInstance &&
            this.value.template === template) {
            this.value.update(value.values);
        }
        else {
            // Make sure we propagate the template processor from the TemplateResult
            // so that we use its syntax extension, etc. The template factory comes
            // from the render function options so that it can control template
            // caching and preprocessing.
            const instance = new TemplateInstance(template, value.processor, this.options);
            const fragment = instance._clone();
            instance.update(value.values);
            this._commitNode(fragment);
            this.value = instance;
        }
    }
    _commitIterable(value) {
        // For an Iterable, we create a new InstancePart per item, then set its
        // value to the item. This is a little bit of overhead for every item in
        // an Iterable, but it lets us recurse easily and efficiently update Arrays
        // of TemplateResults that will be commonly returned from expressions like:
        // array.map((i) => html`${i}`), by reusing existing TemplateInstances.
        // If _value is an array, then the previous render was of an
        // iterable and _value will contain the NodeParts from the previous
        // render. If _value is not an array, clear this part and make a new
        // array for NodeParts.
        if (!Array.isArray(this.value)) {
            this.value = [];
            this.clear();
        }
        // Lets us keep track of how many items we stamped so we can clear leftover
        // items from a previous render
        const itemParts = this.value;
        let partIndex = 0;
        let itemPart;
        for (const item of value) {
            // Try to reuse an existing part
            itemPart = itemParts[partIndex];
            // If no existing part, create a new one
            if (itemPart === undefined) {
                itemPart = new NodePart(this.options);
                itemParts.push(itemPart);
                if (partIndex === 0) {
                    itemPart.appendIntoPart(this);
                }
                else {
                    itemPart.insertAfterPart(itemParts[partIndex - 1]);
                }
            }
            itemPart.setValue(item);
            itemPart.commit();
            partIndex++;
        }
        if (partIndex < itemParts.length) {
            // Truncate the parts array so _value reflects the current state
            itemParts.length = partIndex;
            this.clear(itemPart && itemPart.endNode);
        }
    }
    clear(startNode = this.startNode) {
        removeNodes(this.startNode.parentNode, startNode.nextSibling, this.endNode);
    }
}
/**
 * Implements a boolean attribute, roughly as defined in the HTML
 * specification.
 *
 * If the value is truthy, then the attribute is present with a value of
 * ''. If the value is falsey, the attribute is removed.
 */
class BooleanAttributePart {
    constructor(element, name, strings) {
        this.value = undefined;
        this._pendingValue = undefined;
        if (strings.length !== 2 || strings[0] !== '' || strings[1] !== '') {
            throw new Error('Boolean attributes can only contain a single expression');
        }
        this.element = element;
        this.name = name;
        this.strings = strings;
    }
    setValue(value) {
        this._pendingValue = value;
    }
    commit() {
        while (isDirective(this._pendingValue)) {
            const directive = this._pendingValue;
            this._pendingValue = noChange;
            directive(this);
        }
        if (this._pendingValue === noChange) {
            return;
        }
        const value = !!this._pendingValue;
        if (this.value !== value) {
            if (value) {
                this.element.setAttribute(this.name, '');
            }
            else {
                this.element.removeAttribute(this.name);
            }
        }
        this.value = value;
        this._pendingValue = noChange;
    }
}
/**
 * Sets attribute values for PropertyParts, so that the value is only set once
 * even if there are multiple parts for a property.
 *
 * If an expression controls the whole property value, then the value is simply
 * assigned to the property under control. If there are string literals or
 * multiple expressions, then the strings are expressions are interpolated into
 * a string first.
 */
class PropertyCommitter extends AttributeCommitter {
    constructor(element, name, strings) {
        super(element, name, strings);
        this.single =
            (strings.length === 2 && strings[0] === '' && strings[1] === '');
    }
    _createPart() {
        return new PropertyPart(this);
    }
    _getValue() {
        if (this.single) {
            return this.parts[0].value;
        }
        return super._getValue();
    }
    commit() {
        if (this.dirty) {
            this.dirty = false;
            // tslint:disable-next-line:no-any
            this.element[this.name] = this._getValue();
        }
    }
}
class PropertyPart extends AttributePart {
}
// Detect event listener options support. If the `capture` property is read
// from the options object, then options are supported. If not, then the thrid
// argument to add/removeEventListener is interpreted as the boolean capture
// value so we should only pass the `capture` property.
let eventOptionsSupported = false;
try {
    const options = {
        get capture() {
            eventOptionsSupported = true;
            return false;
        }
    };
    // tslint:disable-next-line:no-any
    window.addEventListener('test', options, options);
    // tslint:disable-next-line:no-any
    window.removeEventListener('test', options, options);
}
catch (_e) {
}
class EventPart {
    constructor(element, eventName, eventContext) {
        this.value = undefined;
        this._pendingValue = undefined;
        this.element = element;
        this.eventName = eventName;
        this.eventContext = eventContext;
        this._boundHandleEvent = (e) => this.handleEvent(e);
    }
    setValue(value) {
        this._pendingValue = value;
    }
    commit() {
        while (isDirective(this._pendingValue)) {
            const directive = this._pendingValue;
            this._pendingValue = noChange;
            directive(this);
        }
        if (this._pendingValue === noChange) {
            return;
        }
        const newListener = this._pendingValue;
        const oldListener = this.value;
        const shouldRemoveListener = newListener == null ||
            oldListener != null &&
                (newListener.capture !== oldListener.capture ||
                    newListener.once !== oldListener.once ||
                    newListener.passive !== oldListener.passive);
        const shouldAddListener = newListener != null && (oldListener == null || shouldRemoveListener);
        if (shouldRemoveListener) {
            this.element.removeEventListener(this.eventName, this._boundHandleEvent, this._options);
        }
        if (shouldAddListener) {
            this._options = getOptions(newListener);
            this.element.addEventListener(this.eventName, this._boundHandleEvent, this._options);
        }
        this.value = newListener;
        this._pendingValue = noChange;
    }
    handleEvent(event) {
        if (typeof this.value === 'function') {
            this.value.call(this.eventContext || this.element, event);
        }
        else {
            this.value.handleEvent(event);
        }
    }
}
// We copy options because of the inconsistent behavior of browsers when reading
// the third argument of add/removeEventListener. IE11 doesn't support options
// at all. Chrome 41 only reads `capture` if the argument is an object.
const getOptions = (o) => o &&
    (eventOptionsSupported ?
        { capture: o.capture, passive: o.passive, once: o.once } :
        o.capture);

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
/**
 * Creates Parts when a template is instantiated.
 */
class DefaultTemplateProcessor {
    /**
     * Create parts for an attribute-position binding, given the event, attribute
     * name, and string literals.
     *
     * @param element The element containing the binding
     * @param name  The attribute name
     * @param strings The string literals. There are always at least two strings,
     *   event for fully-controlled bindings with a single expression.
     */
    handleAttributeExpressions(element, name, strings, options) {
        const prefix = name[0];
        if (prefix === '.') {
            const comitter = new PropertyCommitter(element, name.slice(1), strings);
            return comitter.parts;
        }
        if (prefix === '@') {
            return [new EventPart(element, name.slice(1), options.eventContext)];
        }
        if (prefix === '?') {
            return [new BooleanAttributePart(element, name.slice(1), strings)];
        }
        const comitter = new AttributeCommitter(element, name, strings);
        return comitter.parts;
    }
    /**
     * Create parts for a text-position binding.
     * @param templateFactory
     */
    handleTextExpression(options) {
        return new NodePart(options);
    }
}
const defaultTemplateProcessor = new DefaultTemplateProcessor();

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
/**
 * The default TemplateFactory which caches Templates keyed on
 * result.type and result.strings.
 */
function templateFactory(result) {
    let templateCache = templateCaches.get(result.type);
    if (templateCache === undefined) {
        templateCache = {
            stringsArray: new WeakMap(),
            keyString: new Map()
        };
        templateCaches.set(result.type, templateCache);
    }
    let template = templateCache.stringsArray.get(result.strings);
    if (template !== undefined) {
        return template;
    }
    // If the TemplateStringsArray is new, generate a key from the strings
    // This key is shared between all templates with identical content
    const key = result.strings.join(marker);
    // Check if we already have a Template for this key
    template = templateCache.keyString.get(key);
    if (template === undefined) {
        // If we have not seen this key before, create a new Template
        template = new Template(result, result.getTemplateElement());
        // Cache the Template for this key
        templateCache.keyString.set(key, template);
    }
    // Cache all future queries for this TemplateStringsArray
    templateCache.stringsArray.set(result.strings, template);
    return template;
}
const templateCaches = new Map();

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
const parts = new WeakMap();
/**
 * Renders a template to a container.
 *
 * To update a container with new values, reevaluate the template literal and
 * call `render` with the new result.
 *
 * @param result a TemplateResult created by evaluating a template tag like
 *     `html` or `svg`.
 * @param container A DOM parent to render to. The entire contents are either
 *     replaced, or efficiently updated if the same result type was previous
 *     rendered there.
 * @param options RenderOptions for the entire render tree rendered to this
 *     container. Render options must *not* change between renders to the same
 *     container, as those changes will not effect previously rendered DOM.
 */
const render = (result, container, options) => {
    let part = parts.get(container);
    if (part === undefined) {
        removeNodes(container, container.firstChild);
        parts.set(container, part = new NodePart(Object.assign({ templateFactory }, options)));
        part.appendInto(container);
    }
    part.setValue(result);
    part.commit();
};

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
// IMPORTANT: do not change the property name or the assignment expression.
// This line will be used in regexes to search for lit-html usage.
// TODO(justinfagnani): inject version number at build time
(window['litHtmlVersions'] || (window['litHtmlVersions'] = [])).push('1.0.0');
/**
 * Interprets a template literal as an HTML template that can efficiently
 * render to and update a container.
 */
const html = (strings, ...values) => new TemplateResult(strings, values, 'html', defaultTemplateProcessor);

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
const walkerNodeFilter = 133 /* NodeFilter.SHOW_{ELEMENT|COMMENT|TEXT} */;
/**
 * Removes the list of nodes from a Template safely. In addition to removing
 * nodes from the Template, the Template part indices are updated to match
 * the mutated Template DOM.
 *
 * As the template is walked the removal state is tracked and
 * part indices are adjusted as needed.
 *
 * div
 *   div#1 (remove) <-- start removing (removing node is div#1)
 *     div
 *       div#2 (remove)  <-- continue removing (removing node is still div#1)
 *         div
 * div <-- stop removing since previous sibling is the removing node (div#1,
 * removed 4 nodes)
 */
function removeNodesFromTemplate(template, nodesToRemove) {
    const { element: { content }, parts } = template;
    const walker = document.createTreeWalker(content, walkerNodeFilter, null, false);
    let partIndex = nextActiveIndexInTemplateParts(parts);
    let part = parts[partIndex];
    let nodeIndex = -1;
    let removeCount = 0;
    const nodesToRemoveInTemplate = [];
    let currentRemovingNode = null;
    while (walker.nextNode()) {
        nodeIndex++;
        const node = walker.currentNode;
        // End removal if stepped past the removing node
        if (node.previousSibling === currentRemovingNode) {
            currentRemovingNode = null;
        }
        // A node to remove was found in the template
        if (nodesToRemove.has(node)) {
            nodesToRemoveInTemplate.push(node);
            // Track node we're removing
            if (currentRemovingNode === null) {
                currentRemovingNode = node;
            }
        }
        // When removing, increment count by which to adjust subsequent part indices
        if (currentRemovingNode !== null) {
            removeCount++;
        }
        while (part !== undefined && part.index === nodeIndex) {
            // If part is in a removed node deactivate it by setting index to -1 or
            // adjust the index as needed.
            part.index = currentRemovingNode !== null ? -1 : part.index - removeCount;
            // go to the next active part.
            partIndex = nextActiveIndexInTemplateParts(parts, partIndex);
            part = parts[partIndex];
        }
    }
    nodesToRemoveInTemplate.forEach((n) => n.parentNode.removeChild(n));
}
const countNodes = (node) => {
    let count = (node.nodeType === 11 /* Node.DOCUMENT_FRAGMENT_NODE */) ? 0 : 1;
    const walker = document.createTreeWalker(node, walkerNodeFilter, null, false);
    while (walker.nextNode()) {
        count++;
    }
    return count;
};
const nextActiveIndexInTemplateParts = (parts, startIndex = -1) => {
    for (let i = startIndex + 1; i < parts.length; i++) {
        const part = parts[i];
        if (isTemplatePartActive(part)) {
            return i;
        }
    }
    return -1;
};
/**
 * Inserts the given node into the Template, optionally before the given
 * refNode. In addition to inserting the node into the Template, the Template
 * part indices are updated to match the mutated Template DOM.
 */
function insertNodeIntoTemplate(template, node, refNode = null) {
    const { element: { content }, parts } = template;
    // If there's no refNode, then put node at end of template.
    // No part indices need to be shifted in this case.
    if (refNode === null || refNode === undefined) {
        content.appendChild(node);
        return;
    }
    const walker = document.createTreeWalker(content, walkerNodeFilter, null, false);
    let partIndex = nextActiveIndexInTemplateParts(parts);
    let insertCount = 0;
    let walkerIndex = -1;
    while (walker.nextNode()) {
        walkerIndex++;
        const walkerNode = walker.currentNode;
        if (walkerNode === refNode) {
            insertCount = countNodes(node);
            refNode.parentNode.insertBefore(node, refNode);
        }
        while (partIndex !== -1 && parts[partIndex].index === walkerIndex) {
            // If we've inserted the node, simply adjust all subsequent parts
            if (insertCount > 0) {
                while (partIndex !== -1) {
                    parts[partIndex].index += insertCount;
                    partIndex = nextActiveIndexInTemplateParts(parts, partIndex);
                }
                return;
            }
            partIndex = nextActiveIndexInTemplateParts(parts, partIndex);
        }
    }
}

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
// Get a key to lookup in `templateCaches`.
const getTemplateCacheKey = (type, scopeName) => `${type}--${scopeName}`;
let compatibleShadyCSSVersion = true;
if (typeof window.ShadyCSS === 'undefined') {
    compatibleShadyCSSVersion = false;
}
else if (typeof window.ShadyCSS.prepareTemplateDom === 'undefined') {
    console.warn(`Incompatible ShadyCSS version detected.` +
        `Please update to at least @webcomponents/webcomponentsjs@2.0.2 and` +
        `@webcomponents/shadycss@1.3.1.`);
    compatibleShadyCSSVersion = false;
}
/**
 * Template factory which scopes template DOM using ShadyCSS.
 * @param scopeName {string}
 */
const shadyTemplateFactory = (scopeName) => (result) => {
    const cacheKey = getTemplateCacheKey(result.type, scopeName);
    let templateCache = templateCaches.get(cacheKey);
    if (templateCache === undefined) {
        templateCache = {
            stringsArray: new WeakMap(),
            keyString: new Map()
        };
        templateCaches.set(cacheKey, templateCache);
    }
    let template = templateCache.stringsArray.get(result.strings);
    if (template !== undefined) {
        return template;
    }
    const key = result.strings.join(marker);
    template = templateCache.keyString.get(key);
    if (template === undefined) {
        const element = result.getTemplateElement();
        if (compatibleShadyCSSVersion) {
            window.ShadyCSS.prepareTemplateDom(element, scopeName);
        }
        template = new Template(result, element);
        templateCache.keyString.set(key, template);
    }
    templateCache.stringsArray.set(result.strings, template);
    return template;
};
const TEMPLATE_TYPES = ['html', 'svg'];
/**
 * Removes all style elements from Templates for the given scopeName.
 */
const removeStylesFromLitTemplates = (scopeName) => {
    TEMPLATE_TYPES.forEach((type) => {
        const templates = templateCaches.get(getTemplateCacheKey(type, scopeName));
        if (templates !== undefined) {
            templates.keyString.forEach((template) => {
                const { element: { content } } = template;
                // IE 11 doesn't support the iterable param Set constructor
                const styles = new Set();
                Array.from(content.querySelectorAll('style')).forEach((s) => {
                    styles.add(s);
                });
                removeNodesFromTemplate(template, styles);
            });
        }
    });
};
const shadyRenderSet = new Set();
/**
 * For the given scope name, ensures that ShadyCSS style scoping is performed.
 * This is done just once per scope name so the fragment and template cannot
 * be modified.
 * (1) extracts styles from the rendered fragment and hands them to ShadyCSS
 * to be scoped and appended to the document
 * (2) removes style elements from all lit-html Templates for this scope name.
 *
 * Note, <style> elements can only be placed into templates for the
 * initial rendering of the scope. If <style> elements are included in templates
 * dynamically rendered to the scope (after the first scope render), they will
 * not be scoped and the <style> will be left in the template and rendered
 * output.
 */
const prepareTemplateStyles = (renderedDOM, template, scopeName) => {
    shadyRenderSet.add(scopeName);
    // Move styles out of rendered DOM and store.
    const styles = renderedDOM.querySelectorAll('style');
    // If there are no styles, skip unnecessary work
    if (styles.length === 0) {
        // Ensure prepareTemplateStyles is called to support adding
        // styles via `prepareAdoptedCssText` since that requires that
        // `prepareTemplateStyles` is called.
        window.ShadyCSS.prepareTemplateStyles(template.element, scopeName);
        return;
    }
    const condensedStyle = document.createElement('style');
    // Collect styles into a single style. This helps us make sure ShadyCSS
    // manipulations will not prevent us from being able to fix up template
    // part indices.
    // NOTE: collecting styles is inefficient for browsers but ShadyCSS
    // currently does this anyway. When it does not, this should be changed.
    for (let i = 0; i < styles.length; i++) {
        const style = styles[i];
        style.parentNode.removeChild(style);
        condensedStyle.textContent += style.textContent;
    }
    // Remove styles from nested templates in this scope.
    removeStylesFromLitTemplates(scopeName);
    // And then put the condensed style into the "root" template passed in as
    // `template`.
    insertNodeIntoTemplate(template, condensedStyle, template.element.content.firstChild);
    // Note, it's important that ShadyCSS gets the template that `lit-html`
    // will actually render so that it can update the style inside when
    // needed (e.g. @apply native Shadow DOM case).
    window.ShadyCSS.prepareTemplateStyles(template.element, scopeName);
    if (window.ShadyCSS.nativeShadow) {
        // When in native Shadow DOM, re-add styling to rendered content using
        // the style ShadyCSS produced.
        const style = template.element.content.querySelector('style');
        renderedDOM.insertBefore(style.cloneNode(true), renderedDOM.firstChild);
    }
    else {
        // When not in native Shadow DOM, at this point ShadyCSS will have
        // removed the style from the lit template and parts will be broken as a
        // result. To fix this, we put back the style node ShadyCSS removed
        // and then tell lit to remove that node from the template.
        // NOTE, ShadyCSS creates its own style so we can safely add/remove
        // `condensedStyle` here.
        template.element.content.insertBefore(condensedStyle, template.element.content.firstChild);
        const removes = new Set();
        removes.add(condensedStyle);
        removeNodesFromTemplate(template, removes);
    }
};
/**
 * Extension to the standard `render` method which supports rendering
 * to ShadowRoots when the ShadyDOM (https://github.com/webcomponents/shadydom)
 * and ShadyCSS (https://github.com/webcomponents/shadycss) polyfills are used
 * or when the webcomponentsjs
 * (https://github.com/webcomponents/webcomponentsjs) polyfill is used.
 *
 * Adds a `scopeName` option which is used to scope element DOM and stylesheets
 * when native ShadowDOM is unavailable. The `scopeName` will be added to
 * the class attribute of all rendered DOM. In addition, any style elements will
 * be automatically re-written with this `scopeName` selector and moved out
 * of the rendered DOM and into the document `<head>`.
 *
 * It is common to use this render method in conjunction with a custom element
 * which renders a shadowRoot. When this is done, typically the element's
 * `localName` should be used as the `scopeName`.
 *
 * In addition to DOM scoping, ShadyCSS also supports a basic shim for css
 * custom properties (needed only on older browsers like IE11) and a shim for
 * a deprecated feature called `@apply` that supports applying a set of css
 * custom properties to a given location.
 *
 * Usage considerations:
 *
 * * Part values in `<style>` elements are only applied the first time a given
 * `scopeName` renders. Subsequent changes to parts in style elements will have
 * no effect. Because of this, parts in style elements should only be used for
 * values that will never change, for example parts that set scope-wide theme
 * values or parts which render shared style elements.
 *
 * * Note, due to a limitation of the ShadyDOM polyfill, rendering in a
 * custom element's `constructor` is not supported. Instead rendering should
 * either done asynchronously, for example at microtask timing (for example
 * `Promise.resolve()`), or be deferred until the first time the element's
 * `connectedCallback` runs.
 *
 * Usage considerations when using shimmed custom properties or `@apply`:
 *
 * * Whenever any dynamic changes are made which affect
 * css custom properties, `ShadyCSS.styleElement(element)` must be called
 * to update the element. There are two cases when this is needed:
 * (1) the element is connected to a new parent, (2) a class is added to the
 * element that causes it to match different custom properties.
 * To address the first case when rendering a custom element, `styleElement`
 * should be called in the element's `connectedCallback`.
 *
 * * Shimmed custom properties may only be defined either for an entire
 * shadowRoot (for example, in a `:host` rule) or via a rule that directly
 * matches an element with a shadowRoot. In other words, instead of flowing from
 * parent to child as do native css custom properties, shimmed custom properties
 * flow only from shadowRoots to nested shadowRoots.
 *
 * * When using `@apply` mixing css shorthand property names with
 * non-shorthand names (for example `border` and `border-width`) is not
 * supported.
 */
const render$1 = (result, container, options) => {
    const scopeName = options.scopeName;
    const hasRendered = parts.has(container);
    const needsScoping = container instanceof ShadowRoot &&
        compatibleShadyCSSVersion && result instanceof TemplateResult;
    // Handle first render to a scope specially...
    const firstScopeRender = needsScoping && !shadyRenderSet.has(scopeName);
    // On first scope render, render into a fragment; this cannot be a single
    // fragment that is reused since nested renders can occur synchronously.
    const renderContainer = firstScopeRender ? document.createDocumentFragment() : container;
    render(result, renderContainer, Object.assign({ templateFactory: shadyTemplateFactory(scopeName) }, options));
    // When performing first scope render,
    // (1) We've rendered into a fragment so that there's a chance to
    // `prepareTemplateStyles` before sub-elements hit the DOM
    // (which might cause them to render based on a common pattern of
    // rendering in a custom element's `connectedCallback`);
    // (2) Scope the template with ShadyCSS one time only for this scope.
    // (3) Render the fragment into the container and make sure the
    // container knows its `part` is the one we just rendered. This ensures
    // DOM will be re-used on subsequent renders.
    if (firstScopeRender) {
        const part = parts.get(renderContainer);
        parts.delete(renderContainer);
        if (part.value instanceof TemplateInstance) {
            prepareTemplateStyles(renderContainer, part.value.template, scopeName);
        }
        removeNodes(container, container.firstChild);
        container.appendChild(renderContainer);
        parts.set(container, part);
    }
    // After elements have hit the DOM, update styling if this is the
    // initial render to this container.
    // This is needed whenever dynamic changes are made so it would be
    // safest to do every render; however, this would regress performance
    // so we leave it up to the user to call `ShadyCSSS.styleElement`
    // for dynamic changes.
    if (!hasRendered && needsScoping) {
        window.ShadyCSS.styleElement(container.host);
    }
};

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
/**
 * When using Closure Compiler, JSCompiler_renameProperty(property, object) is
 * replaced at compile time by the munged name for object[property]. We cannot
 * alias this function, so we have to use a small shim that has the same
 * behavior when not compiling.
 */
window.JSCompiler_renameProperty =
    (prop, _obj) => prop;
const defaultConverter = {
    toAttribute(value, type) {
        switch (type) {
            case Boolean:
                return value ? '' : null;
            case Object:
            case Array:
                // if the value is `null` or `undefined` pass this through
                // to allow removing/no change behavior.
                return value == null ? value : JSON.stringify(value);
        }
        return value;
    },
    fromAttribute(value, type) {
        switch (type) {
            case Boolean:
                return value !== null;
            case Number:
                return value === null ? null : Number(value);
            case Object:
            case Array:
                return JSON.parse(value);
        }
        return value;
    }
};
/**
 * Change function that returns true if `value` is different from `oldValue`.
 * This method is used as the default for a property's `hasChanged` function.
 */
const notEqual = (value, old) => {
    // This ensures (old==NaN, value==NaN) always returns false
    return old !== value && (old === old || value === value);
};
const defaultPropertyDeclaration = {
    attribute: true,
    type: String,
    converter: defaultConverter,
    reflect: false,
    hasChanged: notEqual
};
const microtaskPromise = Promise.resolve(true);
const STATE_HAS_UPDATED = 1;
const STATE_UPDATE_REQUESTED = 1 << 2;
const STATE_IS_REFLECTING_TO_ATTRIBUTE = 1 << 3;
const STATE_IS_REFLECTING_TO_PROPERTY = 1 << 4;
const STATE_HAS_CONNECTED = 1 << 5;
/**
 * Base element class which manages element properties and attributes. When
 * properties change, the `update` method is asynchronously called. This method
 * should be supplied by subclassers to render updates as desired.
 */
class UpdatingElement extends HTMLElement {
    constructor() {
        super();
        this._updateState = 0;
        this._instanceProperties = undefined;
        this._updatePromise = microtaskPromise;
        this._hasConnectedResolver = undefined;
        /**
         * Map with keys for any properties that have changed since the last
         * update cycle with previous values.
         */
        this._changedProperties = new Map();
        /**
         * Map with keys of properties that should be reflected when updated.
         */
        this._reflectingProperties = undefined;
        this.initialize();
    }
    /**
     * Returns a list of attributes corresponding to the registered properties.
     * @nocollapse
     */
    static get observedAttributes() {
        // note: piggy backing on this to ensure we're finalized.
        this.finalize();
        const attributes = [];
        // Use forEach so this works even if for/of loops are compiled to for loops
        // expecting arrays
        this._classProperties.forEach((v, p) => {
            const attr = this._attributeNameForProperty(p, v);
            if (attr !== undefined) {
                this._attributeToPropertyMap.set(attr, p);
                attributes.push(attr);
            }
        });
        return attributes;
    }
    /**
     * Ensures the private `_classProperties` property metadata is created.
     * In addition to `finalize` this is also called in `createProperty` to
     * ensure the `@property` decorator can add property metadata.
     */
    /** @nocollapse */
    static _ensureClassProperties() {
        // ensure private storage for property declarations.
        if (!this.hasOwnProperty(JSCompiler_renameProperty('_classProperties', this))) {
            this._classProperties = new Map();
            // NOTE: Workaround IE11 not supporting Map constructor argument.
            const superProperties = Object.getPrototypeOf(this)._classProperties;
            if (superProperties !== undefined) {
                superProperties.forEach((v, k) => this._classProperties.set(k, v));
            }
        }
    }
    /**
     * Creates a property accessor on the element prototype if one does not exist.
     * The property setter calls the property's `hasChanged` property option
     * or uses a strict identity check to determine whether or not to request
     * an update.
     * @nocollapse
     */
    static createProperty(name, options = defaultPropertyDeclaration) {
        // Note, since this can be called by the `@property` decorator which
        // is called before `finalize`, we ensure storage exists for property
        // metadata.
        this._ensureClassProperties();
        this._classProperties.set(name, options);
        // Do not generate an accessor if the prototype already has one, since
        // it would be lost otherwise and that would never be the user's intention;
        // Instead, we expect users to call `requestUpdate` themselves from
        // user-defined accessors. Note that if the super has an accessor we will
        // still overwrite it
        if (options.noAccessor || this.prototype.hasOwnProperty(name)) {
            return;
        }
        const key = typeof name === 'symbol' ? Symbol() : `__${name}`;
        Object.defineProperty(this.prototype, name, {
            // tslint:disable-next-line:no-any no symbol in index
            get() {
                return this[key];
            },
            set(value) {
                // tslint:disable-next-line:no-any no symbol in index
                const oldValue = this[name];
                // tslint:disable-next-line:no-any no symbol in index
                this[key] = value;
                this._requestUpdate(name, oldValue);
            },
            configurable: true,
            enumerable: true
        });
    }
    /**
     * Creates property accessors for registered properties and ensures
     * any superclasses are also finalized.
     * @nocollapse
     */
    static finalize() {
        if (this.hasOwnProperty(JSCompiler_renameProperty('finalized', this)) &&
            this.finalized) {
            return;
        }
        // finalize any superclasses
        const superCtor = Object.getPrototypeOf(this);
        if (typeof superCtor.finalize === 'function') {
            superCtor.finalize();
        }
        this.finalized = true;
        this._ensureClassProperties();
        // initialize Map populated in observedAttributes
        this._attributeToPropertyMap = new Map();
        // make any properties
        // Note, only process "own" properties since this element will inherit
        // any properties defined on the superClass, and finalization ensures
        // the entire prototype chain is finalized.
        if (this.hasOwnProperty(JSCompiler_renameProperty('properties', this))) {
            const props = this.properties;
            // support symbols in properties (IE11 does not support this)
            const propKeys = [
                ...Object.getOwnPropertyNames(props),
                ...(typeof Object.getOwnPropertySymbols === 'function') ?
                    Object.getOwnPropertySymbols(props) :
                    []
            ];
            // This for/of is ok because propKeys is an array
            for (const p of propKeys) {
                // note, use of `any` is due to TypeSript lack of support for symbol in
                // index types
                // tslint:disable-next-line:no-any no symbol in index
                this.createProperty(p, props[p]);
            }
        }
    }
    /**
     * Returns the property name for the given attribute `name`.
     * @nocollapse
     */
    static _attributeNameForProperty(name, options) {
        const attribute = options.attribute;
        return attribute === false ?
            undefined :
            (typeof attribute === 'string' ?
                attribute :
                (typeof name === 'string' ? name.toLowerCase() : undefined));
    }
    /**
     * Returns true if a property should request an update.
     * Called when a property value is set and uses the `hasChanged`
     * option for the property if present or a strict identity check.
     * @nocollapse
     */
    static _valueHasChanged(value, old, hasChanged = notEqual) {
        return hasChanged(value, old);
    }
    /**
     * Returns the property value for the given attribute value.
     * Called via the `attributeChangedCallback` and uses the property's
     * `converter` or `converter.fromAttribute` property option.
     * @nocollapse
     */
    static _propertyValueFromAttribute(value, options) {
        const type = options.type;
        const converter = options.converter || defaultConverter;
        const fromAttribute = (typeof converter === 'function' ? converter : converter.fromAttribute);
        return fromAttribute ? fromAttribute(value, type) : value;
    }
    /**
     * Returns the attribute value for the given property value. If this
     * returns undefined, the property will *not* be reflected to an attribute.
     * If this returns null, the attribute will be removed, otherwise the
     * attribute will be set to the value.
     * This uses the property's `reflect` and `type.toAttribute` property options.
     * @nocollapse
     */
    static _propertyValueToAttribute(value, options) {
        if (options.reflect === undefined) {
            return;
        }
        const type = options.type;
        const converter = options.converter;
        const toAttribute = converter && converter.toAttribute ||
            defaultConverter.toAttribute;
        return toAttribute(value, type);
    }
    /**
     * Performs element initialization. By default captures any pre-set values for
     * registered properties.
     */
    initialize() {
        this._saveInstanceProperties();
        // ensures first update will be caught by an early access of `updateComplete`
        this._requestUpdate();
    }
    /**
     * Fixes any properties set on the instance before upgrade time.
     * Otherwise these would shadow the accessor and break these properties.
     * The properties are stored in a Map which is played back after the
     * constructor runs. Note, on very old versions of Safari (<=9) or Chrome
     * (<=41), properties created for native platform properties like (`id` or
     * `name`) may not have default values set in the element constructor. On
     * these browsers native properties appear on instances and therefore their
     * default value will overwrite any element default (e.g. if the element sets
     * this.id = 'id' in the constructor, the 'id' will become '' since this is
     * the native platform default).
     */
    _saveInstanceProperties() {
        // Use forEach so this works even if for/of loops are compiled to for loops
        // expecting arrays
        this.constructor
            ._classProperties.forEach((_v, p) => {
            if (this.hasOwnProperty(p)) {
                const value = this[p];
                delete this[p];
                if (!this._instanceProperties) {
                    this._instanceProperties = new Map();
                }
                this._instanceProperties.set(p, value);
            }
        });
    }
    /**
     * Applies previously saved instance properties.
     */
    _applyInstanceProperties() {
        // Use forEach so this works even if for/of loops are compiled to for loops
        // expecting arrays
        // tslint:disable-next-line:no-any
        this._instanceProperties.forEach((v, p) => this[p] = v);
        this._instanceProperties = undefined;
    }
    connectedCallback() {
        this._updateState = this._updateState | STATE_HAS_CONNECTED;
        // Ensure first connection completes an update. Updates cannot complete before
        // connection and if one is pending connection the `_hasConnectionResolver`
        // will exist. If so, resolve it to complete the update, otherwise
        // requestUpdate.
        if (this._hasConnectedResolver) {
            this._hasConnectedResolver();
            this._hasConnectedResolver = undefined;
        }
    }
    /**
     * Allows for `super.disconnectedCallback()` in extensions while
     * reserving the possibility of making non-breaking feature additions
     * when disconnecting at some point in the future.
     */
    disconnectedCallback() {
    }
    /**
     * Synchronizes property values when attributes change.
     */
    attributeChangedCallback(name, old, value) {
        if (old !== value) {
            this._attributeToProperty(name, value);
        }
    }
    _propertyToAttribute(name, value, options = defaultPropertyDeclaration) {
        const ctor = this.constructor;
        const attr = ctor._attributeNameForProperty(name, options);
        if (attr !== undefined) {
            const attrValue = ctor._propertyValueToAttribute(value, options);
            // an undefined value does not change the attribute.
            if (attrValue === undefined) {
                return;
            }
            // Track if the property is being reflected to avoid
            // setting the property again via `attributeChangedCallback`. Note:
            // 1. this takes advantage of the fact that the callback is synchronous.
            // 2. will behave incorrectly if multiple attributes are in the reaction
            // stack at time of calling. However, since we process attributes
            // in `update` this should not be possible (or an extreme corner case
            // that we'd like to discover).
            // mark state reflecting
            this._updateState = this._updateState | STATE_IS_REFLECTING_TO_ATTRIBUTE;
            if (attrValue == null) {
                this.removeAttribute(attr);
            }
            else {
                this.setAttribute(attr, attrValue);
            }
            // mark state not reflecting
            this._updateState = this._updateState & ~STATE_IS_REFLECTING_TO_ATTRIBUTE;
        }
    }
    _attributeToProperty(name, value) {
        // Use tracking info to avoid deserializing attribute value if it was
        // just set from a property setter.
        if (this._updateState & STATE_IS_REFLECTING_TO_ATTRIBUTE) {
            return;
        }
        const ctor = this.constructor;
        const propName = ctor._attributeToPropertyMap.get(name);
        if (propName !== undefined) {
            const options = ctor._classProperties.get(propName) || defaultPropertyDeclaration;
            // mark state reflecting
            this._updateState = this._updateState | STATE_IS_REFLECTING_TO_PROPERTY;
            this[propName] =
                // tslint:disable-next-line:no-any
                ctor._propertyValueFromAttribute(value, options);
            // mark state not reflecting
            this._updateState = this._updateState & ~STATE_IS_REFLECTING_TO_PROPERTY;
        }
    }
    /**
     * This private version of `requestUpdate` does not access or return the
     * `updateComplete` promise. This promise can be overridden and is therefore
     * not free to access.
     */
    _requestUpdate(name, oldValue) {
        let shouldRequestUpdate = true;
        // If we have a property key, perform property update steps.
        if (name !== undefined) {
            const ctor = this.constructor;
            const options = ctor._classProperties.get(name) || defaultPropertyDeclaration;
            if (ctor._valueHasChanged(this[name], oldValue, options.hasChanged)) {
                if (!this._changedProperties.has(name)) {
                    this._changedProperties.set(name, oldValue);
                }
                // Add to reflecting properties set.
                // Note, it's important that every change has a chance to add the
                // property to `_reflectingProperties`. This ensures setting
                // attribute + property reflects correctly.
                if (options.reflect === true &&
                    !(this._updateState & STATE_IS_REFLECTING_TO_PROPERTY)) {
                    if (this._reflectingProperties === undefined) {
                        this._reflectingProperties = new Map();
                    }
                    this._reflectingProperties.set(name, options);
                }
            }
            else {
                // Abort the request if the property should not be considered changed.
                shouldRequestUpdate = false;
            }
        }
        if (!this._hasRequestedUpdate && shouldRequestUpdate) {
            this._enqueueUpdate();
        }
    }
    /**
     * Requests an update which is processed asynchronously. This should
     * be called when an element should update based on some state not triggered
     * by setting a property. In this case, pass no arguments. It should also be
     * called when manually implementing a property setter. In this case, pass the
     * property `name` and `oldValue` to ensure that any configured property
     * options are honored. Returns the `updateComplete` Promise which is resolved
     * when the update completes.
     *
     * @param name {PropertyKey} (optional) name of requesting property
     * @param oldValue {any} (optional) old value of requesting property
     * @returns {Promise} A Promise that is resolved when the update completes.
     */
    requestUpdate(name, oldValue) {
        this._requestUpdate(name, oldValue);
        return this.updateComplete;
    }
    /**
     * Sets up the element to asynchronously update.
     */
    async _enqueueUpdate() {
        // Mark state updating...
        this._updateState = this._updateState | STATE_UPDATE_REQUESTED;
        let resolve;
        let reject;
        const previousUpdatePromise = this._updatePromise;
        this._updatePromise = new Promise((res, rej) => {
            resolve = res;
            reject = rej;
        });
        try {
            // Ensure any previous update has resolved before updating.
            // This `await` also ensures that property changes are batched.
            await previousUpdatePromise;
        }
        catch (e) {
            // Ignore any previous errors. We only care that the previous cycle is
            // done. Any error should have been handled in the previous update.
        }
        // Make sure the element has connected before updating.
        if (!this._hasConnected) {
            await new Promise((res) => this._hasConnectedResolver = res);
        }
        try {
            const result = this.performUpdate();
            // If `performUpdate` returns a Promise, we await it. This is done to
            // enable coordinating updates with a scheduler. Note, the result is
            // checked to avoid delaying an additional microtask unless we need to.
            if (result != null) {
                await result;
            }
        }
        catch (e) {
            reject(e);
        }
        resolve(!this._hasRequestedUpdate);
    }
    get _hasConnected() {
        return (this._updateState & STATE_HAS_CONNECTED);
    }
    get _hasRequestedUpdate() {
        return (this._updateState & STATE_UPDATE_REQUESTED);
    }
    get hasUpdated() {
        return (this._updateState & STATE_HAS_UPDATED);
    }
    /**
     * Performs an element update. Note, if an exception is thrown during the
     * update, `firstUpdated` and `updated` will not be called.
     *
     * You can override this method to change the timing of updates. If this
     * method is overridden, `super.performUpdate()` must be called.
     *
     * For instance, to schedule updates to occur just before the next frame:
     *
     * ```
     * protected async performUpdate(): Promise<unknown> {
     *   await new Promise((resolve) => requestAnimationFrame(() => resolve()));
     *   super.performUpdate();
     * }
     * ```
     */
    performUpdate() {
        // Mixin instance properties once, if they exist.
        if (this._instanceProperties) {
            this._applyInstanceProperties();
        }
        let shouldUpdate = false;
        const changedProperties = this._changedProperties;
        try {
            shouldUpdate = this.shouldUpdate(changedProperties);
            if (shouldUpdate) {
                this.update(changedProperties);
            }
        }
        catch (e) {
            // Prevent `firstUpdated` and `updated` from running when there's an
            // update exception.
            shouldUpdate = false;
            throw e;
        }
        finally {
            // Ensure element can accept additional updates after an exception.
            this._markUpdated();
        }
        if (shouldUpdate) {
            if (!(this._updateState & STATE_HAS_UPDATED)) {
                this._updateState = this._updateState | STATE_HAS_UPDATED;
                this.firstUpdated(changedProperties);
            }
            this.updated(changedProperties);
        }
    }
    _markUpdated() {
        this._changedProperties = new Map();
        this._updateState = this._updateState & ~STATE_UPDATE_REQUESTED;
    }
    /**
     * Returns a Promise that resolves when the element has completed updating.
     * The Promise value is a boolean that is `true` if the element completed the
     * update without triggering another update. The Promise result is `false` if
     * a property was set inside `updated()`. If the Promise is rejected, an
     * exception was thrown during the update. This getter can be implemented to
     * await additional state. For example, it is sometimes useful to await a
     * rendered element before fulfilling this Promise. To do this, first await
     * `super.updateComplete` then any subsequent state.
     *
     * @returns {Promise} The Promise returns a boolean that indicates if the
     * update resolved without triggering another update.
     */
    get updateComplete() {
        return this._updatePromise;
    }
    /**
     * Controls whether or not `update` should be called when the element requests
     * an update. By default, this method always returns `true`, but this can be
     * customized to control when to update.
     *
     * * @param _changedProperties Map of changed properties with old values
     */
    shouldUpdate(_changedProperties) {
        return true;
    }
    /**
     * Updates the element. This method reflects property values to attributes.
     * It can be overridden to render and keep updated element DOM.
     * Setting properties inside this method will *not* trigger
     * another update.
     *
     * * @param _changedProperties Map of changed properties with old values
     */
    update(_changedProperties) {
        if (this._reflectingProperties !== undefined &&
            this._reflectingProperties.size > 0) {
            // Use forEach so this works even if for/of loops are compiled to for
            // loops expecting arrays
            this._reflectingProperties.forEach((v, k) => this._propertyToAttribute(k, this[k], v));
            this._reflectingProperties = undefined;
        }
    }
    /**
     * Invoked whenever the element is updated. Implement to perform
     * post-updating tasks via DOM APIs, for example, focusing an element.
     *
     * Setting properties inside this method will trigger the element to update
     * again after this update cycle completes.
     *
     * * @param _changedProperties Map of changed properties with old values
     */
    updated(_changedProperties) {
    }
    /**
     * Invoked when the element is first updated. Implement to perform one time
     * work on the element after update.
     *
     * Setting properties inside this method will trigger the element to update
     * again after this update cycle completes.
     *
     * * @param _changedProperties Map of changed properties with old values
     */
    firstUpdated(_changedProperties) {
    }
}
/**
 * Marks class as having finished creating properties.
 */
UpdatingElement.finalized = true;

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
const legacyCustomElement = (tagName, clazz) => {
    window.customElements.define(tagName, clazz);
    // Cast as any because TS doesn't recognize the return type as being a
    // subtype of the decorated class when clazz is typed as
    // `Constructor<HTMLElement>` for some reason.
    // `Constructor<HTMLElement>` is helpful to make sure the decorator is
    // applied to elements however.
    // tslint:disable-next-line:no-any
    return clazz;
};
const standardCustomElement = (tagName, descriptor) => {
    const { kind, elements } = descriptor;
    return {
        kind,
        elements,
        // This callback is called once the class is otherwise fully defined
        finisher(clazz) {
            window.customElements.define(tagName, clazz);
        }
    };
};
/**
 * Class decorator factory that defines the decorated class as a custom element.
 *
 * @param tagName the name of the custom element to define
 */
const customElement = (tagName) => (classOrDescriptor) => (typeof classOrDescriptor === 'function') ?
    legacyCustomElement(tagName, classOrDescriptor) :
    standardCustomElement(tagName, classOrDescriptor);
const standardProperty = (options, element) => {
    // When decorating an accessor, pass it through and add property metadata.
    // Note, the `hasOwnProperty` check in `createProperty` ensures we don't
    // stomp over the user's accessor.
    if (element.kind === 'method' && element.descriptor &&
        !('value' in element.descriptor)) {
        return Object.assign({}, element, { finisher(clazz) {
                clazz.createProperty(element.key, options);
            } });
    }
    else {
        // createProperty() takes care of defining the property, but we still
        // must return some kind of descriptor, so return a descriptor for an
        // unused prototype field. The finisher calls createProperty().
        return {
            kind: 'field',
            key: Symbol(),
            placement: 'own',
            descriptor: {},
            // When @babel/plugin-proposal-decorators implements initializers,
            // do this instead of the initializer below. See:
            // https://github.com/babel/babel/issues/9260 extras: [
            //   {
            //     kind: 'initializer',
            //     placement: 'own',
            //     initializer: descriptor.initializer,
            //   }
            // ],
            // tslint:disable-next-line:no-any decorator
            initializer() {
                if (typeof element.initializer === 'function') {
                    this[element.key] = element.initializer.call(this);
                }
            },
            finisher(clazz) {
                clazz.createProperty(element.key, options);
            }
        };
    }
};
const legacyProperty = (options, proto, name) => {
    proto.constructor
        .createProperty(name, options);
};
/**
 * A property decorator which creates a LitElement property which reflects a
 * corresponding attribute value. A `PropertyDeclaration` may optionally be
 * supplied to configure property features.
 *
 * @ExportDecoratedItems
 */
function property(options) {
    // tslint:disable-next-line:no-any decorator
    return (protoOrDescriptor, name) => (name !== undefined) ?
        legacyProperty(options, protoOrDescriptor, name) :
        standardProperty(options, protoOrDescriptor);
}

/**
@license
Copyright (c) 2019 The Polymer Project Authors. All rights reserved.
This code may only be used under the BSD style license found at
http://polymer.github.io/LICENSE.txt The complete set of authors may be found at
http://polymer.github.io/AUTHORS.txt The complete set of contributors may be
found at http://polymer.github.io/CONTRIBUTORS.txt Code distributed by Google as
part of the polymer project is also subject to an additional IP rights grant
found at http://polymer.github.io/PATENTS.txt
*/
const supportsAdoptingStyleSheets = ('adoptedStyleSheets' in Document.prototype) &&
    ('replace' in CSSStyleSheet.prototype);

/**
 * @license
 * Copyright (c) 2017 The Polymer Project Authors. All rights reserved.
 * This code may only be used under the BSD style license found at
 * http://polymer.github.io/LICENSE.txt
 * The complete set of authors may be found at
 * http://polymer.github.io/AUTHORS.txt
 * The complete set of contributors may be found at
 * http://polymer.github.io/CONTRIBUTORS.txt
 * Code distributed by Google as part of the polymer project is also
 * subject to an additional IP rights grant found at
 * http://polymer.github.io/PATENTS.txt
 */
// IMPORTANT: do not change the property name or the assignment expression.
// This line will be used in regexes to search for LitElement usage.
// TODO(justinfagnani): inject version number at build time
(window['litElementVersions'] || (window['litElementVersions'] = []))
    .push('2.0.1');
/**
 * Minimal implementation of Array.prototype.flat
 * @param arr the array to flatten
 * @param result the accumlated result
 */
function arrayFlat(styles, result = []) {
    for (let i = 0, length = styles.length; i < length; i++) {
        const value = styles[i];
        if (Array.isArray(value)) {
            arrayFlat(value, result);
        }
        else {
            result.push(value);
        }
    }
    return result;
}
/** Deeply flattens styles array. Uses native flat if available. */
const flattenStyles = (styles) => styles.flat ? styles.flat(Infinity) : arrayFlat(styles);
class LitElement extends UpdatingElement {
    /** @nocollapse */
    static finalize() {
        super.finalize();
        // Prepare styling that is stamped at first render time. Styling
        // is built from user provided `styles` or is inherited from the superclass.
        this._styles =
            this.hasOwnProperty(JSCompiler_renameProperty('styles', this)) ?
                this._getUniqueStyles() :
                this._styles || [];
    }
    /** @nocollapse */
    static _getUniqueStyles() {
        // Take care not to call `this.styles` multiple times since this generates
        // new CSSResults each time.
        // TODO(sorvell): Since we do not cache CSSResults by input, any
        // shared styles will generate new stylesheet objects, which is wasteful.
        // This should be addressed when a browser ships constructable
        // stylesheets.
        const userStyles = this.styles;
        const styles = [];
        if (Array.isArray(userStyles)) {
            const flatStyles = flattenStyles(userStyles);
            // As a performance optimization to avoid duplicated styling that can
            // occur especially when composing via subclassing, de-duplicate styles
            // preserving the last item in the list. The last item is kept to
            // try to preserve cascade order with the assumption that it's most
            // important that last added styles override previous styles.
            const styleSet = flatStyles.reduceRight((set, s) => {
                set.add(s);
                // on IE set.add does not return the set.
                return set;
            }, new Set());
            // Array.from does not work on Set in IE
            styleSet.forEach((v) => styles.unshift(v));
        }
        else if (userStyles) {
            styles.push(userStyles);
        }
        return styles;
    }
    /**
     * Performs element initialization. By default this calls `createRenderRoot`
     * to create the element `renderRoot` node and captures any pre-set values for
     * registered properties.
     */
    initialize() {
        super.initialize();
        this.renderRoot =
            this.createRenderRoot();
        // Note, if renderRoot is not a shadowRoot, styles would/could apply to the
        // element's getRootNode(). While this could be done, we're choosing not to
        // support this now since it would require different logic around de-duping.
        if (window.ShadowRoot && this.renderRoot instanceof window.ShadowRoot) {
            this.adoptStyles();
        }
    }
    /**
     * Returns the node into which the element should render and by default
     * creates and returns an open shadowRoot. Implement to customize where the
     * element's DOM is rendered. For example, to render into the element's
     * childNodes, return `this`.
     * @returns {Element|DocumentFragment} Returns a node into which to render.
     */
    createRenderRoot() {
        return this.attachShadow({ mode: 'open' });
    }
    /**
     * Applies styling to the element shadowRoot using the `static get styles`
     * property. Styling will apply using `shadowRoot.adoptedStyleSheets` where
     * available and will fallback otherwise. When Shadow DOM is polyfilled,
     * ShadyCSS scopes styles and adds them to the document. When Shadow DOM
     * is available but `adoptedStyleSheets` is not, styles are appended to the
     * end of the `shadowRoot` to [mimic spec
     * behavior](https://wicg.github.io/construct-stylesheets/#using-constructed-stylesheets).
     */
    adoptStyles() {
        const styles = this.constructor._styles;
        if (styles.length === 0) {
            return;
        }
        // There are three separate cases here based on Shadow DOM support.
        // (1) shadowRoot polyfilled: use ShadyCSS
        // (2) shadowRoot.adoptedStyleSheets available: use it.
        // (3) shadowRoot.adoptedStyleSheets polyfilled: append styles after
        // rendering
        if (window.ShadyCSS !== undefined && !window.ShadyCSS.nativeShadow) {
            window.ShadyCSS.ScopingShim.prepareAdoptedCssText(styles.map((s) => s.cssText), this.localName);
        }
        else if (supportsAdoptingStyleSheets) {
            this.renderRoot.adoptedStyleSheets =
                styles.map((s) => s.styleSheet);
        }
        else {
            // This must be done after rendering so the actual style insertion is done
            // in `update`.
            this._needsShimAdoptedStyleSheets = true;
        }
    }
    connectedCallback() {
        super.connectedCallback();
        // Note, first update/render handles styleElement so we only call this if
        // connected after first update.
        if (this.hasUpdated && window.ShadyCSS !== undefined) {
            window.ShadyCSS.styleElement(this);
        }
    }
    /**
     * Updates the element. This method reflects property values to attributes
     * and calls `render` to render DOM via lit-html. Setting properties inside
     * this method will *not* trigger another update.
     * * @param _changedProperties Map of changed properties with old values
     */
    update(changedProperties) {
        super.update(changedProperties);
        const templateResult = this.render();
        if (templateResult instanceof TemplateResult) {
            this.constructor
                .render(templateResult, this.renderRoot, { scopeName: this.localName, eventContext: this });
        }
        // When native Shadow DOM is used but adoptedStyles are not supported,
        // insert styling after rendering to ensure adoptedStyles have highest
        // priority.
        if (this._needsShimAdoptedStyleSheets) {
            this._needsShimAdoptedStyleSheets = false;
            this.constructor._styles.forEach((s) => {
                const style = document.createElement('style');
                style.textContent = s.cssText;
                this.renderRoot.appendChild(style);
            });
        }
    }
    /**
     * Invoked on each update to perform rendering tasks. This method must return
     * a lit-html TemplateResult. Setting properties inside this method will *not*
     * trigger the element to update.
     */
    render() {
    }
}
/**
 * Ensure this class is marked as `finalized` as an optimization ensuring
 * it will not needlessly try to `finalize`.
 */
LitElement.finalized = true;
/**
 * Render method used to render the lit-html TemplateResult to the element's
 * DOM.
 * @param {TemplateResult} Template to render.
 * @param {Element|DocumentFragment} Node into which to render.
 * @param {String} Element name.
 * @nocollapse
 */
LitElement.render = render$1;

var css = "/*\n  Your use of the content in the files referenced here is subject to the terms of the license at http://aka.ms/fabric-assets-license\n*/\n:host {\n  --font-size: 14px;\n  --font-weight: 600;\n  --height: '100%';\n  --margin: 0;\n  --padding: 12px 20px;\n  --color: #212121;\n  --background-color: transparent;\n  --background-color--hover: #eaeaea;\n  --popup-content-background-color: white;\n  --popup-command-font-size: 12px; }\n\n.root {\n  position: relative; }\n\n.login-button {\n  display: flex;\n  align-items: center;\n  font-family: var(--default-font-family);\n  font-size: var(--font-size);\n  font-weight: var(--font-weight);\n  height: var(--height);\n  margin: var(--margin);\n  padding: var(--padding);\n  color: var(--color);\n  background-color: var(--background-color);\n  border: none;\n  cursor: pointer;\n  transition: color 0.3s, background-color 0.3s; }\n  .login-button:hover {\n    color: var(--theme-primary-color);\n    background-color: var(--background-color--hover); }\n  .login-button:focus {\n    outline: 0; }\n\n.login-icon + span {\n  margin-left: 6px; }\n\n.popup {\n  display: none;\n  position: absolute;\n  top: 0;\n  right: 0;\n  animation-duration: 300ms;\n  font-family: var(--default-font-family);\n  background: var(--popup-content-background-color);\n  box-shadow: 0 12px 40px 2px rgba(0, 0, 0, 0.08);\n  min-width: 240px; }\n  .popup.show-menu {\n    display: block;\n    animation-name: fade-in; }\n\n.popup-content {\n  display: flex;\n  flex-direction: column;\n  padding: 24px 48px 16px 24px; }\n\n.popup-commands ul {\n  list-style-type: none;\n  margin: 16px 0 0;\n  padding: 0; }\n\n.popup-command {\n  font-family: var(--default-font-family);\n  font-size: var(--popup-command-font-size);\n  font-weight: var(--font-weight);\n  color: var(--theme-primary-color);\n  background-color: transparent;\n  border: none;\n  padding: 0;\n  cursor: pointer;\n  transition: color 0.3s; }\n  .popup-command:hover {\n    color: var(--theme-dark-color); }\n\n@keyframes fade-in {\n  from {\n    opacity: 0; }\n  to {\n    opacity: 1; } }\n";

var css$1 = "/*\n  Your use of the content in the files referenced here is subject to the terms of the license at http://aka.ms/fabric-assets-license\n*/\n:host {\n  --font-size: 14px;\n  --font-weight: 600;\n  --color: #212121;\n  --email-font-size: 12px;\n  --email-color: #333333;\n  --avatar-size--s: 24px;\n  --avatar-size: 48px;\n  --avatar-font-size--s: 16px;\n  --avatar-font-size: 32px;\n  --initials-color: white;\n  --initials-background-color: #b4009e; }\n\nsvg {\n  width: var(--avatar-size--s);\n  height: var(--avatar-size--s); }\n\nimg {\n  border: 0;\n  border-radius: 50%; }\n\n.root {\n  position: relative;\n  display: grid;\n  align-items: center;\n  grid-template-columns: auto 1fr;\n  font-family: var(--default-font-family); }\n\n.user-avatar {\n  grid-column: 1;\n  width: var(--avatar-size);\n  height: var(--avatar-size);\n  margin-right: 12px; }\n  .user-avatar.initials {\n    display: flex;\n    justify-content: center;\n    align-items: center;\n    color: var(--initials-color);\n    background-color: var(--initials-background-color);\n    border-radius: 50%;\n    font-weight: var(--font-weight); }\n  .user-avatar.small {\n    margin-right: 6px;\n    width: var(--avatar-size--s);\n    height: var(--avatar-size--s); }\n  .user-avatar.row-span-2 {\n    grid-row: 1 / 3; }\n  .user-avatar.row-span-3 {\n    grid-row: 1 / 4; }\n\n.avatar-icon {\n  grid-column: 1;\n  font-size: var(--avatar-font-size);\n  margin-right: 12px; }\n  .avatar-icon.small {\n    margin-right: 6px;\n    font-size: var(--avatar-font-size--s); }\n  .avatar-icon.row-span-2 {\n    grid-row: 1 / 3; }\n  .avatar-icon.row-span-3 {\n    grid-row: 1 / 4; }\n\n.user-name {\n  font-size: var(--font-size);\n  font-weight: var(--font-weight);\n  white-space: nowrap;\n  grid-column: 2; }\n\n.user-email {\n  color: var(--email-color);\n  font-size: var(--email-font-size);\n  white-space: nowrap;\n  grid-column: 2; }\n";

var __decorate = (undefined && undefined.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __awaiter$6 = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
let MgtPerson = class MgtPerson extends LitElement {
    constructor() {
        super();
        this.imageSize = 24;
        Providers.onProvidersChanged(_ => this.handleProviderChanged());
        this.loadImage();
    }
    attributeChangedCallback(name, oldval, newval) {
        super.attributeChangedCallback(name, oldval, newval);
        if (name == "person-query" && oldval !== newval) {
            this.personDetails = null;
            this.loadImage();
        }
    }
    static get styles() {
        return css$1;
    }
    handleProviderChanged() {
        let provider = Providers.getAvailable();
        if (provider.isLoggedIn) {
            this.loadImage();
        }
        provider.onLoginChanged(_ => this.loadImage());
    }
    loadImage() {
        return __awaiter$6(this, void 0, void 0, function* () {
            if (!this.personDetails && this.personQuery) {
                let provider = Providers.getAvailable();
                if (provider && provider.isLoggedIn) {
                    if (this.personQuery == "me") {
                        let person = {};
                        yield Promise.all([
                            provider.graph.me().then(user => {
                                if (user) {
                                    person.displayName = user.displayName;
                                    person.email = user.mail;
                                }
                            }),
                            provider.graph.myPhoto().then(photo => {
                                if (photo) {
                                    person.image = photo;
                                }
                            })
                        ]);
                        this.personDetails = person;
                    }
                    else {
                        provider.graph.findPerson(this.personQuery).then(people => {
                            if (people && people.length > 0) {
                                let person = people[0];
                                this.personDetails = person;
                                if (person.scoredEmailAddresses &&
                                    person.scoredEmailAddresses.length) {
                                    this.personDetails.email =
                                        person.scoredEmailAddresses[0].address;
                                }
                                else if (person.emailAddresses &&
                                    person.emailAddresses.length) {
                                    // beta endpoind uses emailAddresses instead of scoredEmailAddresses
                                    this.personDetails.email = (person).emailAddresses[0].address;
                                }
                                if (person.userPrincipalName) {
                                    let userPrincipalName = person.userPrincipalName;
                                    provider.graph.getUserPhoto(userPrincipalName).then(photo => {
                                        this.personDetails.image = photo;
                                        this.requestUpdate();
                                    });
                                }
                            }
                        });
                    }
                }
                else {
                    this.personDetails = null;
                }
            }
        });
    }
    render() {
        return html `
      <div class="root">
        ${this.renderImage()} ${this.renderNameAndEmail()}
      </div>
    `;
    }
    renderImage() {
        if (this.personDetails) {
            if (this.personDetails.image) {
                return html `
          <img
            class="user-avatar ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"
            src=${this.personDetails.image}
          />
        `;
            }
            else {
                return html `
          <div
            class="user-avatar initials ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"
          >
            <style>
              .initials-text {
                font-size: ${this.imageSize * 0.45}px;
              }
            </style>
            <span class="initials-text">
              ${this.getInitials()}
            </span>
          </div>
        `;
            }
        }
        return this.renderEmptyImage();
    }
    renderEmptyImage() {
        return html `
      <i
        class="ms-Icon ms-Icon--Contact avatar-icon ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"
      ></i>
    `;
    }
    renderNameAndEmail() {
        if (!this.personDetails || (!this.showEmail && !this.showName)) {
            return;
        }
        const nameView = this.showName
            ? html `
          <div class="user-name">${this.personDetails.displayName}</div>
        `
            : null;
        const emailView = this.showEmail
            ? html `
          <div class="user-email">${this.personDetails.email}</div>
        `
            : null;
        return html `
      ${nameView} ${emailView}
    `;
    }
    getInitials() {
        if (!this.personDetails) {
            return "";
        }
        let initials = "";
        if (this.personDetails.givenName) {
            initials += this.personDetails.givenName[0].toUpperCase();
        }
        if (this.personDetails.surname) {
            initials += this.personDetails.surname[0].toUpperCase();
        }
        if (!initials && this.personDetails.displayName) {
            const name = this.personDetails.displayName.split(" ");
            for (let i = 0; i < 2 && i < name.length; i++) {
                initials += name[i][0].toUpperCase();
            }
        }
        return initials;
    }
    getImageRowSpanClass() {
        if (this.showEmail && this.showName) {
            return "row-span-2";
        }
        return "";
    }
    getImageSizeClass() {
        if (!this.showEmail || !this.showName) {
            return "small";
        }
        return "";
    }
};
__decorate([
    property({
        attribute: "image-size"
    })
], MgtPerson.prototype, "imageSize", void 0);
__decorate([
    property({
        attribute: "person-query"
    })
], MgtPerson.prototype, "personQuery", void 0);
__decorate([
    property({
        attribute: "show-name",
        type: Boolean
    })
], MgtPerson.prototype, "showName", void 0);
__decorate([
    property({
        attribute: "show-email",
        type: Boolean
    })
], MgtPerson.prototype, "showEmail", void 0);
__decorate([
    property()
], MgtPerson.prototype, "personDetails", void 0);
MgtPerson = __decorate([
    customElement("mgt-person")
], MgtPerson);

var __decorate$1 = (undefined && undefined.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __awaiter$7 = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
let MgtLogin = class MgtLogin extends LitElement {
    constructor() {
        super();
        this._showMenu = false;
        Providers.onProvidersChanged(_ => this.init());
        this.init();
    }
    static get styles() {
        return css;
    }
    updated(changedProps) {
        if (changedProps.get('_showMenu') === false) {
            // get popup bounds
            const popup = this.shadowRoot.querySelector('.popup');
            this._popupRect = popup.getBoundingClientRect();
            // console.log('last', this._popupRect);
            // invert variables
            const deltaX = this._loginButtonRect.left - this._popupRect.left;
            const deltaY = this._loginButtonRect.top - this._popupRect.top;
            const deltaW = this._loginButtonRect.width / this._popupRect.width;
            const deltaH = this._loginButtonRect.height / this._popupRect.height;
            // play back
            popup.animate([
                {
                    transformOrigin: 'top left',
                    transform: `
              translate(${deltaX}px, ${deltaY}px)
              scale(${deltaW}, ${deltaH})
            `,
                    backgroundColor: `#eaeaea`
                },
                {
                    transformOrigin: 'top left',
                    transform: 'none',
                    backgroundColor: `white`
                }
            ], {
                duration: 300,
                easing: 'ease-in-out',
                fill: 'both'
            });
        }
        else if (changedProps.get('_showMenu') === true) {
            // get login button bounds
            const loginButton = this.shadowRoot.querySelector('.login-button');
            this._loginButtonRect = loginButton.getBoundingClientRect();
            // console.log('last', this._loginButtonRect);
            // invert variables
            const deltaX = this._popupRect.left - this._loginButtonRect.left;
            const deltaY = this._popupRect.top - this._loginButtonRect.top;
            const deltaW = this._popupRect.width / this._loginButtonRect.width;
            const deltaH = this._popupRect.height / this._loginButtonRect.height;
            // play back
            loginButton.animate([
                {
                    transformOrigin: 'top left',
                    transform: `
               translate(${deltaX}px, ${deltaY}px)
               scale(${deltaW}, ${deltaH})
             `
                },
                {
                    transformOrigin: 'top left',
                    transform: 'none'
                }
            ], {
                duration: 200,
                easing: 'ease-out',
                fill: 'both'
            });
        }
    }
    firstUpdated() {
        window.onclick = (event) => {
            if (event.target !== this) {
                // get popup bounds
                const popup = this.shadowRoot.querySelector('.popup');
                this._popupRect = popup.getBoundingClientRect();
                // console.log('first', this._popupRect);
                this._showMenu = false;
            }
        };
    }
    init() {
        return __awaiter$7(this, void 0, void 0, function* () {
            const provider = Providers.getAvailable();
            if (provider) {
                provider.onLoginChanged(_ => this.loadState());
                yield this.loadState();
            }
        });
    }
    loadState() {
        return __awaiter$7(this, void 0, void 0, function* () {
            const provider = Providers.getAvailable();
            if (provider && provider.isLoggedIn) {
                this._user = yield provider.graph.me();
            }
        });
    }
    clicked() {
        if (this._user) {
            // get login button bounds
            const loginButton = this.shadowRoot.querySelector('.login-button');
            this._loginButtonRect = loginButton.getBoundingClientRect();
            // console.log('first', this._loginButtonRect);
            this._showMenu = true;
        }
        else {
            this.login();
        }
    }
    login() {
        return __awaiter$7(this, void 0, void 0, function* () {
            const provider = Providers.getAvailable();
            if (provider) {
                yield provider.login();
                yield this.loadState();
            }
        });
    }
    logout() {
        return __awaiter$7(this, void 0, void 0, function* () {
            const provider = Providers.getAvailable();
            if (provider) {
                yield provider.logout();
            }
        });
    }
    render() {
        const content = this._user ? this.renderLoggedIn() : this.renderLoggedOut();
        return html `
      <div class="root">
        <button class="login-button" @click=${this.clicked}>
          ${content}
        </button>
        ${this.renderMenu()}
      </div>
    `;
    }
    renderLoggedOut() {
        return html `
      <i class="login-icon ms-Icon ms-Icon--Contact"></i>
      <span>
        Sign In
      </span>
    `;
    }
    renderLoggedIn() {
        return html `
      <mgt-person person-query="me" show-name />
    `;
    }
    renderMenu() {
        if (!this._user) {
            return;
        }
        return html `
      <div class="popup ${this._showMenu ? 'show-menu' : ''}">
        <div class="popup-content">
          <div>
            <mgt-person person-query="me" show-name show-email />
          </div>
          <div class="popup-commands">
            <ul>
              <li>
                <button class="popup-command" @click=${this.logout}>
                  Sign Out
                </button>
              </li>
            </ul>
          </div>
        </div>
      </div>
    `;
    }
};
__decorate$1([
    property({ attribute: false })
], MgtLogin.prototype, "_user", void 0);
__decorate$1([
    property({ attribute: false })
], MgtLogin.prototype, "_showMenu", void 0);
MgtLogin = __decorate$1([
    customElement('mgt-login')
], MgtLogin);

var css$2 = "/*\n  Your use of the content in the files referenced here is subject to the terms of the license at http://aka.ms/fabric-assets-license\n*/\n.agenda-list {\n  list-style-type: none;\n  padding: 0;\n  margin: 0;\n  color: #000000b7;\n  font-family: var(--default-font-family);\n  font-style: normal;\n  font-weight: normal; }\n\n.agenda-event {\n  border-color: rgba(16, 16, 16, 0.3);\n  border-style: solid;\n  border-width: 0px 0px 1px 0px;\n  margin: 8px 8px 0px 8px;\n  padding: 8px 0px;\n  display: flex; }\n\n.event-time-container {\n  display: flex;\n  flex-direction: column;\n  font-size: 10px;\n  color: rgba(16, 16, 16, 0.6);\n  font-weight: 600;\n  margin: 6px 0px;\n  width: 50px; }\n\n.event-duration {\n  color: rgba(16, 16, 16, 0.3); }\n\n.event-details-container {\n  margin: 2px 4px 4px 16px; }\n\n.event-location {\n  font-size: 10px;\n  color: rgba(16, 16, 16, 0.4); }\n\n.event-attendie-list {\n  display: flex;\n  list-style-type: none;\n  padding: 0;\n  margin: 12px 0 0 0; }\n\n.event-attendie {\n  margin: 0 6px 12px 0; }\n";

var __decorate$2 = (undefined && undefined.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __awaiter$8 = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
let MgtAgenda = class MgtAgenda extends LitElement {
    constructor() {
        super();
        Providers.onProvidersChanged(_ => this.init());
        this.init();
    }
    static get styles() {
        return css$2;
    }
    init() {
        return __awaiter$8(this, void 0, void 0, function* () {
            this._provider = Providers.getAvailable();
            if (this._provider) {
                this._provider.onLoginChanged(_ => this.loadData());
                yield this.loadData();
            }
        });
    }
    loadData() {
        return __awaiter$8(this, void 0, void 0, function* () {
            if (this._provider && this._provider.isLoggedIn) {
                let today = new Date();
                let tomorrow = new Date();
                tomorrow.setDate(today.getDate() + 2);
                this._events = yield this._provider.graph.calendar(today, tomorrow);
            }
        });
    }
    render() {
        if (this._events) {
            // remove slotted elements inserted initially
            while (this.lastChild) {
                this.removeChild(this.lastChild);
            }
            return html `
        <ul class="agenda-list">
          ${this._events.map(event => html `
                <li>
                  ${this.eventTemplateFunction
                ? this.renderEventTemplate(event)
                : this.renderEvent(event)}
                </li>
              `)}
        </ul>
      `;
        }
        else {
            return html `
        <div>no things</div>
      `;
        }
    }
    renderEvent(event) {
        return html `
      <div class="agenda-event">
        <div class="event-time-container">
          <div>${this.getStartingTime(event)}</div>
          <div class="event-duration">${this.getEventDuration(event)}</div>
        </div>
        <div class="event-details-container">
          <div class="event-subject">${event.subject}</div>
          <div class="event-attendies">
            <ul class="event-attendie-list">
              ${event.attendees.slice(0, 5).map(at => html `
                    <li class="event-attendie">
                      <mgt-person
                        person-query=${at.emailAddress.address}
                        image-size="30"
                      ></mgt-person>
                    </li>
                  `)}
            </ul>
          </div>
          <div class="event-location">${event.location.displayName}</div>
        </div>
      </div>
    `;
    }
    renderEventTemplate(event) {
        let content = this.eventTemplateFunction(event);
        if (typeof content === 'string') {
            return html `
        <div>${this.eventTemplateFunction(event)}</div>
      `;
        }
        else {
            let div = document.createElement('div');
            div.slot = event.subject;
            div.appendChild(content);
            this.appendChild(div);
            return html `
        <slot name=${event.subject}></slot>
      `;
        }
    }
    getStartingTime(event) {
        if (event.isAllDay) {
            return 'ALL DAY';
        }
        let dt = new Date(event.start.dateTime);
        dt.setMinutes(dt.getMinutes() - dt.getTimezoneOffset());
        let hours = dt.getHours();
        let minutes = dt.getMinutes();
        let ampm = hours >= 12 ? 'PM' : 'AM';
        hours = hours % 12;
        hours = hours ? hours : 12;
        let minutesStr = minutes < 10 ? '0' + minutes : minutes;
        return `${hours}:${minutesStr} ${ampm}`;
    }
    getEventDuration(event) {
        let dtStart = new Date(event.start.dateTime);
        let dtEnd = new Date(event.end.dateTime);
        let dtNow = new Date();
        let result = '';
        if (dtNow > dtStart) {
            dtStart = dtNow;
        }
        let diff = dtEnd.getTime() - dtStart.getTime();
        var durationMinutes = Math.round(diff / 60000);
        if (durationMinutes > 1440 || event.isAllDay) {
            result = Math.ceil(durationMinutes / 1440) + 'd';
        }
        else if (durationMinutes > 60) {
            result = Math.round(durationMinutes / 60) + 'h';
            let leftoverMinutes = durationMinutes % 60;
            if (leftoverMinutes) {
                result += leftoverMinutes + 'm';
            }
        }
        else {
            result = durationMinutes + 'm';
        }
        return result;
    }
};
__decorate$2([
    property({ attribute: false })
], MgtAgenda.prototype, "_events", void 0);
__decorate$2([
    property()
], MgtAgenda.prototype, "eventTemplateFunction", void 0);
MgtAgenda = __decorate$2([
    customElement('mgt-agenda')
], MgtAgenda);

var __decorate$3 = (undefined && undefined.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
let MgtMsalProvider = class MgtMsalProvider extends LitElement {
    constructor() {
        super();
        this._isInitialized = false;
        this.clientId = '';
        this.validateAuthProps();
    }
    attributeChangedCallback(name, oldval, newval) {
        super.attributeChangedCallback(name, oldval, newval);
        if (this._isInitialized) {
            this.validateAuthProps();
        }
        // console.log("property changed " + name + " = " + newval);
        this.validateAuthProps();
    }
    firstUpdated(changedProperties) {
        this._isInitialized = true;
        this.validateAuthProps();
    }
    validateAuthProps() {
        if (this.clientId) {
            let config = {
                clientId: this.clientId,
            };
            if (this.loginType && this.loginType.length > 1) {
                let loginType = this.loginType.toLowerCase();
                loginType = loginType[0].toUpperCase() + loginType.slice(1);
                let loginTypeEnum = LoginType[loginType];
                config.loginType = loginTypeEnum;
            }
            if (this.authority) {
                config.authority = this.authority;
            }
            Providers.addMsalProvider(config);
        }
    }
};
__decorate$3([
    property({
        type: String,
        attribute: 'client-id'
    })
], MgtMsalProvider.prototype, "clientId", void 0);
__decorate$3([
    property({
        type: String,
        attribute: 'login-type'
    })
], MgtMsalProvider.prototype, "loginType", void 0);
__decorate$3([
    property()
], MgtMsalProvider.prototype, "authority", void 0);
MgtMsalProvider = __decorate$3([
    customElement('mgt-msal-provider')
], MgtMsalProvider);

var __decorate$4 = (undefined && undefined.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
let MgtWamProvider = class MgtWamProvider extends LitElement {
    constructor() {
        super();
        this.validateAuthProps();
    }
    attributeChangedCallback(name, oldval, newval) {
        super.attributeChangedCallback(name, oldval, newval);
        this.validateAuthProps();
    }
    validateAuthProps() {
        if (this.clientId !== undefined) {
            Providers.addWamProvider(this.clientId, this.authority);
        }
    }
};
__decorate$4([
    property({ attribute: 'client-id' })
], MgtWamProvider.prototype, "clientId", void 0);
__decorate$4([
    property({ attribute: 'authority' })
], MgtWamProvider.prototype, "authority", void 0);
MgtWamProvider = __decorate$4([
    customElement('mgt-wam-provider')
], MgtWamProvider);

var __decorate$5 = (undefined && undefined.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __awaiter$9 = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
let MgtTeamsProvider = class MgtTeamsProvider extends LitElement {
    constructor() {
        super();
        this._isInitialized = false;
        this.clientId = '';
        this.loginPopupUrl = '';
        this.validateAuthProps();
    }
    attributeChangedCallback(name, oldval, newval) {
        super.attributeChangedCallback(name, oldval, newval);
        if (this._isInitialized) {
            this.validateAuthProps();
        }
    }
    firstUpdated(changedProperties) {
        return __awaiter$9(this, void 0, void 0, function* () {
            this._isInitialized = true;
            this.validateAuthProps();
            if (yield TeamsProvider.isAvailable()) {
                Providers.addCustomProvider(this._provider);
            }
        });
    }
    validateAuthProps() {
        if (this.clientId && this.loginPopupUrl) {
            if (!this._provider) {
                this._provider = new TeamsProvider(this.clientId, this.loginPopupUrl);
            }
            this._provider.clientId = this.clientId;
            this._provider.loginPopupUrl = this.loginPopupUrl;
        }
    }
};
__decorate$5([
    property({
        type: String,
        attribute: 'client-id'
    })
], MgtTeamsProvider.prototype, "clientId", void 0);
__decorate$5([
    property({
        type: String,
        attribute: 'login-popup-url'
    })
], MgtTeamsProvider.prototype, "loginPopupUrl", void 0);
MgtTeamsProvider = __decorate$5([
    customElement('mgt-teams-provider')
], MgtTeamsProvider);

export { EventDispatcher, Graph, LoginType, MgtAgenda, MgtLogin, MgtMsalProvider, MgtPerson, MgtTeamsProvider, MgtWamProvider, MsalProvider, Providers, SharePointProvider, TeamsProvider, WamProvider };
