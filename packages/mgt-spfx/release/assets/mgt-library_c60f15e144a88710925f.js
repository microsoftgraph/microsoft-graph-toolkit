/*! For license information please see mgt-library_c60f15e144a88710925f.js.LICENSE.txt */
define("78b11c7d-7ca8-47cb-a93c-d3beabb519a1_3.0.1",["@microsoft/microsoft-graph-client"],function(n){return function(e){var t={};function n(a){if(t[a])return t[a].exports;var i=t[a]={i:a,l:!1,exports:{}};return e[a].call(i.exports,i,i.exports,n),i.l=!0,i.exports}return n.m=e,n.c=t,n.d=function(e,t,a){n.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:a})},n.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},n.t=function(e,t){if(1&t&&(e=n(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var a=Object.create(null);if(n.r(a),Object.defineProperty(a,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var i in e)n.d(a,i,function(t){return e[t]}.bind(null,i));return a},n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,"a",t),t},n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},n.p="",n(n.s="mwqp")}({"+53S":function(e,t,n){"use strict";n.d(t,"a",function(){return r});var a=n("/w5G");class i{constructor(e,t){this.target=e,this.propertyName=t}bind(e){e[this.propertyName]=this.target}unbind(){}}function r(e){return new a.a("fast-ref",i,e)}},"+Cud":function(e,t,n){"use strict";function a(e){const t=e.parentElement;if(t)return t;{const t=e.getRootNode();if(t.host instanceof HTMLElement)return t.host}return null}n.d(t,"a",function(){return a})},"+yEz":function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("5ZAu");class i{constructor(){this.targets=new WeakSet}addStylesTo(e){this.targets.add(e)}removeStylesFrom(e){this.targets.delete(e)}isAttachedTo(e){return this.targets.has(e)}withBehaviors(...e){return this.behaviors=null===this.behaviors?e:this.behaviors.concat(e),this}}function r(e){return e.map(e=>e instanceof i?r(e.styles):[e]).reduce((e,t)=>e.concat(t),[])}function o(e){return e.map(e=>e instanceof i?e.behaviors:null).reduce((e,t)=>null===t?e:(null===e&&(e=[]),e.concat(t)),null)}i.create=(()=>{if(a.a.supportsAdoptedStyleSheets){const e=new Map;return t=>new d(t,e)}return e=>new u(e)})();let s=(e,t)=>{e.adoptedStyleSheets=[...e.adoptedStyleSheets,...t]},c=(e,t)=>{e.adoptedStyleSheets=e.adoptedStyleSheets.filter(e=>-1===t.indexOf(e))};if(a.a.supportsAdoptedStyleSheets)try{document.adoptedStyleSheets.push(),document.adoptedStyleSheets.splice(),s=(e,t)=>{e.adoptedStyleSheets.push(...t)},c=(e,t)=>{for(const n of t){const t=e.adoptedStyleSheets.indexOf(n);-1!==t&&e.adoptedStyleSheets.splice(t,1)}}}catch(e){}class d extends i{constructor(e,t){super(),this.styles=e,this.styleSheetCache=t,this._styleSheets=void 0,this.behaviors=o(e)}get styleSheets(){if(void 0===this._styleSheets){const e=this.styles,t=this.styleSheetCache;this._styleSheets=r(e).map(e=>{if(e instanceof CSSStyleSheet)return e;let n=t.get(e);return void 0===n&&(n=new CSSStyleSheet,n.replaceSync(e),t.set(e,n)),n})}return this._styleSheets}addStylesTo(e){s(e,this.styleSheets),super.addStylesTo(e)}removeStylesFrom(e){c(e,this.styleSheets),super.removeStylesFrom(e)}}let l=0;class u extends i{constructor(e){super(),this.styles=e,this.behaviors=null,this.behaviors=o(e),this.styleSheets=r(e),this.styleClass="fast-style-class-"+ ++l}addStylesTo(e){const t=this.styleSheets,n=this.styleClass;e=this.normalizeTarget(e);for(let a=0;a<t.length;a++){const i=document.createElement("style");i.innerHTML=t[a],i.className=n,e.append(i)}super.addStylesTo(e)}removeStylesFrom(e){const t=(e=this.normalizeTarget(e)).querySelectorAll(`.${this.styleClass}`);for(let n=0,a=t.length;n<a;++n)e.removeChild(t[n]);super.removeStylesFrom(e)}isAttachedTo(e){return super.isAttachedTo(this.normalizeTarget(e))}normalizeTarget(e){return e===document?document.body:e}}},"/49y":function(e,t,n){"use strict";n.d(t,"a",function(){return a}),n.d(t,"b",function(){return i});const a="https://graph.microsoft.com",i=new Set([a,"https://graph.microsoft.us","https://dod-graph.microsoft.us","https://graph.microsoft.de","https://microsoftgraph.chinacloudapi.cn","https://canary.graph.microsoft.com"])},"/4Fm":function(e,t,n){"use strict";n.d(t,"l",function(){return i}),n.d(t,"f",function(){return r}),n.d(t,"n",function(){return o}),n.d(t,"j",function(){return s}),n.d(t,"g",function(){return c}),n.d(t,"h",function(){return d}),n.d(t,"e",function(){return l}),n.d(t,"b",function(){return u}),n.d(t,"a",function(){return f}),n.d(t,"c",function(){return p}),n.d(t,"o",function(){return m}),n.d(t,"d",function(){return _}),n.d(t,"p",function(){return h}),n.d(t,"q",function(){return b}),n.d(t,"k",function(){return g}),n.d(t,"m",function(){return v}),n.d(t,"i",function(){return y});var a=n("DNu6");const i=e=>{const t=new Date,n=new Date(t.getFullYear(),t.getMonth(),t.getDate());if(e>=n)return e.toLocaleString("default",{hour:"numeric",minute:"numeric"});const a=new Date(n);if(a.setDate(t.getDate()-t.getDay()),e>=a)return e.toLocaleString("default",{hour:"numeric",minute:"numeric",weekday:"short"});const i=new Date(a);return i.setDate(a.getDate()-7),e>=i?e.toLocaleString("default",{day:"numeric",month:"numeric",weekday:"short"}):e.toLocaleString("default",{day:"numeric",month:"numeric",year:"numeric"})},r=e=>{const t=e.getMonth();return`${e.getDate()} / ${t} / ${e.getFullYear()}`},o=e=>{const t=e.getMonth(),n=e.getDate();return`${s(t)} ${n}`},s=e=>{switch(e){case 0:return"January";case 1:return"February";case 2:return"March";case 3:return"April";case 4:return"May";case 5:return"June";case 6:return"July";case 7:return"August";case 8:return"September";case 9:return"October";case 10:return"November";case 11:return"December";default:return"Month"}},c=e=>{switch(e){case 0:return"Sunday";case 1:return"Monday";case 2:return"Tuesday";case 3:return"Wednesday";case 4:return"Thursday";case 5:return"Friday";case 6:return"Saturday";default:return"Day"}},d=e=>{switch(e){case 1:return 28;case 3:case 5:case 8:case 10:default:return 30;case 0:case 2:case 4:case 6:case 7:case 9:case 11:return 31}},l=(e,t)=>{const n=`${t}`;let a=`${e}`;return a.length<2&&(a="0"+a),new Date(`${n}-${a}-1T12:00:00-${(new Date).getTimezoneOffset()/60}`)},u=(e,t)=>{let n;return function(){window.clearTimeout(n),n=window.setTimeout(()=>e.apply(this,arguments),t)}},f=e=>new Promise((t,n)=>{const a=new FileReader;a.onerror=n,a.onload=()=>{t(a.result)},a.readAsDataURL(e)}),p=e=>e.startsWith("[")?e.match(/([a-zA-Z0-9+._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi).toString():e,m=e=>/^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/.test(e),_=(e,t=2)=>{if(0===e)return"0 Bytes";const n=t<0?0:t,a=Math.floor(Math.log(e)/Math.log(1024));return`${parseFloat((e/Math.pow(1024,a)).toFixed(n))} ${["Bytes","KB","MB","GB","TB","PB","EB","ZB","YB"][a]}`},h=e=>(e&&(e=null==(e=null==(e=null==e?void 0:e.replace(/<ddd\/>/gi,"..."))?void 0:e.replace(/<c0>/gi,"<b>"))?void 0:e.replace(/<\/c0>/gi,"</b>")),e),b=e=>null==e?void 0:e.replace(/\.[^/.]+$/,""),g=e=>new URL(e).pathname.split("/").pop().replace(/-/g," "),v=e=>e||a.a.config.response.invalidationPeriod||a.a.config.defaultInvalidationPeriod,y=()=>a.a.config.response.isEnabled&&a.a.config.isEnabled},"/i08":function(e,t,n){"use strict";n.d(t,"a",function(){return c}),n.d(t,"b",function(){return a}),n.d(t,"c",function(){return i});var a,i,r=n("q/PQ"),o=n("/49y"),s=n("x+GM");class c{get isMultiAccountSupportedAndEnabled(){return!1}set baseURL(e){if(!Object(r.a)(e))throw new Error(`${e} is not a valid Graph URL endpoint.`);this._baseURL=e}get baseURL(){return this._baseURL}set customHosts(e){this._customHosts=e}get customHosts(){return this._customHosts}get isMultiAccountSupported(){return this.isMultipleAccountSupported}get state(){return this._state}get isIncrementalConsentDisabled(){return this._isIncrementalConsentDisabled}set isIncrementalConsentDisabled(e){this._isIncrementalConsentDisabled=e}get name(){return"MgtIProvider"}constructor(){this.isMultipleAccountDisabled=!0,this._loginChangedDispatcher=new s.a,this._activeAccountChangedDispatcher=new s.a,this._baseURL=o.a,this._customHosts=void 0,this._isIncrementalConsentDisabled=!1,this.isMultipleAccountSupported=!1,this._state=i.Loading}setState(e){e!==this._state&&(this._state=e,this._loginChangedDispatcher.fire({detail:this._state}))}onStateChanged(e){this._loginChangedDispatcher.add(e)}removeStateChangedHandler(e){this._loginChangedDispatcher.remove(e)}setActiveAccount(e){this.fireActiveAccountChanged({detail:e})}onActiveAccountChanged(e){this._activeAccountChangedDispatcher.add(e)}removeActiveAccountChangedHandler(e){this._activeAccountChangedDispatcher.remove(e)}fireActiveAccountChanged(e){this._activeAccountChangedDispatcher.fire(e)}getAccessTokenForScopes(...e){return this.getAccessToken({scopes:e})}}!function(e){e[e.Popup=0]="Popup",e[e.Redirect=1]="Redirect"}(a||(a={})),function(e){e[e.Loading=0]="Loading",e[e.SignedOut=1]="SignedOut",e[e.SignedIn=2]="SignedIn"}(i||(i={}))},"/w5G":function(e,t,n){"use strict";n.d(t,"b",function(){return i}),n.d(t,"c",function(){return r}),n.d(t,"a",function(){return o});var a=n("5ZAu");class i{constructor(){this.targetIndex=0}}class r extends i{constructor(){super(...arguments),this.createPlaceholder=a.a.createInterpolationPlaceholder}}class o extends i{constructor(e,t,n){super(),this.name=e,this.behavior=t,this.options=n}createPlaceholder(e){return a.a.createCustomAttributePlaceholder(this.name,e)}createBehavior(e){return new this.behavior(e,this.options)}}},"0Uyf":function(e,t,n){"use strict";n.d(t,"J",function(){return l}),n.d(t,"a",function(){return u}),n.d(t,"i",function(){return _}),n.d(t,"g",function(){return h}),n.d(t,"h",function(){return b}),n.d(t,"p",function(){return g}),n.d(t,"q",function(){return v}),n.d(t,"u",function(){return y}),n.d(t,"v",function(){return S}),n.d(t,"y",function(){return D}),n.d(t,"z",function(){return I}),n.d(t,"t",function(){return x}),n.d(t,"D",function(){return C}),n.d(t,"E",function(){return O}),n.d(t,"w",function(){return w}),n.d(t,"H",function(){return E}),n.d(t,"n",function(){return L}),n.d(t,"e",function(){return k}),n.d(t,"f",function(){return M}),n.d(t,"r",function(){return P}),n.d(t,"s",function(){return T}),n.d(t,"j",function(){return U}),n.d(t,"l",function(){return F}),n.d(t,"A",function(){return H}),n.d(t,"B",function(){return R}),n.d(t,"F",function(){return N}),n.d(t,"G",function(){return B}),n.d(t,"k",function(){return j}),n.d(t,"x",function(){return V}),n.d(t,"I",function(){return z}),n.d(t,"m",function(){return G}),n.d(t,"c",function(){return J}),n.d(t,"d",function(){return X}),n.d(t,"o",function(){return Z}),n.d(t,"C",function(){return $}),n.d(t,"K",function(){return ee}),n.d(t,"L",function(){return te}),n.d(t,"b",function(){return ne});var a=n("DNu6"),i=n("RIOo"),r=n("WHff"),o=n("imsm"),s=n("EYXE"),c=n("/4Fm"),d=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const l=e=>Array.isArray(e.nextExpectedRanges),u=()=>d(void 0,void 0,void 0,function*(){const e=a.a.getCache(o.a.fileLists,o.a.fileLists.stores.fileLists);yield e.clearStore()}),f=()=>a.a.config.files.invalidationPeriod||a.a.config.defaultInvalidationPeriod,p=()=>a.a.config.files.isEnabled&&a.a.config.isEnabled,m=()=>a.a.config.fileLists.isEnabled&&a.a.config.isEnabled,_=(e,t,n=o.a.files.stores.fileQueries,r="files.read")=>d(void 0,void 0,void 0,function*(){const s=a.a.getCache(o.a.files,n),c=yield Q(s,t);if(c)return c;let d;try{d=yield e.api(t).middlewareOptions(Object(i.a)(r)).get(),p()&&(yield s.putValue(t,{file:JSON.stringify(d)}))}catch(e){}return d||null}),h=(e,t,n)=>d(void 0,void 0,void 0,function*(){return _(e,`/drives/${t}/items/${n}`,o.a.files.stores.driveFiles)}),b=(e,t,n)=>d(void 0,void 0,void 0,function*(){return _(e,`/drives/${t}/root:/${n}`,o.a.files.stores.driveFiles)}),g=(e,t,n)=>d(void 0,void 0,void 0,function*(){return _(e,`/groups/${t}/drive/items/${n}`,o.a.files.stores.groupFiles)}),v=(e,t,n)=>d(void 0,void 0,void 0,function*(){return _(e,`/groups/${t}/drive/root:/${n}`,o.a.files.stores.groupFiles)}),y=(e,t)=>d(void 0,void 0,void 0,function*(){return _(e,`/me/drive/items/${t}`,o.a.files.stores.userFiles)}),S=(e,t)=>d(void 0,void 0,void 0,function*(){return _(e,`/me/drive/root:/${t}`,o.a.files.stores.userFiles)}),D=(e,t,n)=>d(void 0,void 0,void 0,function*(){return _(e,`/sites/${t}/drive/items/${n}`,o.a.files.stores.siteFiles)}),I=(e,t,n)=>d(void 0,void 0,void 0,function*(){return _(e,`/sites/${t}/drive/root:/${n}`,o.a.files.stores.siteFiles)}),x=(e,t,n,a)=>d(void 0,void 0,void 0,function*(){return _(e,`/sites/${t}/lists/${n}/items/${a}/driveItem`,o.a.files.stores.siteFiles)}),C=(e,t,n)=>d(void 0,void 0,void 0,function*(){return _(e,`/users/${t}/drive/items/${n}`,o.a.files.stores.userFiles)}),O=(e,t,n)=>d(void 0,void 0,void 0,function*(){return _(e,`/users/${t}/drive/root:/${n}`,o.a.files.stores.userFiles)}),w=(e,t,n)=>d(void 0,void 0,void 0,function*(){return _(e,`/me/insights/${t}/${n}/resource`,o.a.files.stores.insightFiles,"sites.read.all")}),E=(e,t,n,a)=>d(void 0,void 0,void 0,function*(){return _(e,`/users/${t}/insights/${n}/${a}/resource`,o.a.files.stores.insightFiles,"sites.read.all")}),A=(e,t,n,r,s)=>d(void 0,void 0,void 0,function*(){let c;const d=a.a.getCache(o.a.fileLists,n),l=yield Y(d,n,`${t}:${s}`);if(l)return c=q(e,l.files,l.nextLink),c;let u;try{if(u=e.api(t).middlewareOptions(Object(i.a)(...r)),s&&u.top(s),c=yield W(e,u),m()){const e=c.nextLink;yield d.putValue(t,{files:c.value.map(e=>JSON.stringify(e)),nextLink:e})}}catch(e){}return c||null}),L=(e,t)=>d(void 0,void 0,void 0,function*(){const n=o.a.fileLists.stores.fileLists;return A(e,"/me/drive/root/children",n,["files.read"],t)}),k=(e,t,n,a)=>d(void 0,void 0,void 0,function*(){const i=`/drives/${t}/items/${n}/children`,r=o.a.fileLists.stores.fileLists;return A(e,i,r,["files.read"],a)}),M=(e,t,n,a)=>d(void 0,void 0,void 0,function*(){const i=`/drives/${t}/root:/${n}:/children`,r=o.a.fileLists.stores.fileLists;return A(e,i,r,["files.read"],a)}),P=(e,t,n,a)=>d(void 0,void 0,void 0,function*(){const i=`/groups/${t}/drive/items/${n}/children`,r=o.a.fileLists.stores.fileLists;return A(e,i,r,["files.read"],a)}),T=(e,t,n,a)=>d(void 0,void 0,void 0,function*(){const i=`/groups/${t}/drive/root:/${n}:/children`,r=o.a.fileLists.stores.fileLists;return A(e,i,r,["files.read"],a)}),U=(e,t,n)=>d(void 0,void 0,void 0,function*(){const a=`/me/drive/items/${t}/children`,i=o.a.fileLists.stores.fileLists;return A(e,a,i,["files.read"],n)}),F=(e,t,n)=>d(void 0,void 0,void 0,function*(){const a=`/me/drive/root:/${t}:/children`,i=o.a.fileLists.stores.fileLists;return A(e,a,i,["files.read"],n)}),H=(e,t,n,a)=>d(void 0,void 0,void 0,function*(){const i=`/sites/${t}/drive/items/${n}/children`,r=o.a.fileLists.stores.fileLists;return A(e,i,r,["files.read"],a)}),R=(e,t,n,a)=>d(void 0,void 0,void 0,function*(){const i=`/sites/${t}/drive/root:/${n}:/children`,r=o.a.fileLists.stores.fileLists;return A(e,i,r,["files.read"],a)}),N=(e,t,n,a)=>d(void 0,void 0,void 0,function*(){const i=`/users/${t}/drive/items/${n}/children`,r=o.a.fileLists.stores.fileLists;return A(e,i,r,["files.read"],a)}),B=(e,t,n,a)=>d(void 0,void 0,void 0,function*(){const i=`/users/${t}/drive/root:/${n}:/children`,r=o.a.fileLists.stores.fileLists;return A(e,i,r,["files.read"],a)}),j=(e,t,n)=>d(void 0,void 0,void 0,function*(){const a=o.a.fileLists.stores.fileLists;return A(e,t,a,["files.read","sites.read.all"],n)}),V=(e,t)=>d(void 0,void 0,void 0,function*(){const n=`/me/insights/${t}`,r=o.a.fileLists.stores.insightfileLists,s=a.a.getCache(o.a.fileLists,r),c=yield Y(s,r,n);if(c)return c.files.map(e=>JSON.parse(e));const d=["sites.read.all"];let l;try{l=yield e.api(n).filter("resourceReference/type eq 'microsoft.graph.driveItem'").middlewareOptions(Object(i.a)(...d)).get()}catch(e){}const u=yield K(e,l,d);return m()&&(yield s.putValue(n,{files:u.map(e=>JSON.stringify(e))})),u||null}),z=(e,t,n)=>d(void 0,void 0,void 0,function*(){let r,s;"shared"===n?(r="/me/insights/shared",s=`((lastshared/sharedby/id eq '${t}') and (resourceReference/type eq 'microsoft.graph.driveItem'))`):(r=`/users/${t}/insights/${n}`,s="resourceReference/type eq 'microsoft.graph.driveItem'");const c=`${r}?$filter=${s}`,d=o.a.fileLists.stores.insightfileLists,l=a.a.getCache(o.a.fileLists,d),u=yield Y(l,d,c);if(u)return u.files.map(e=>JSON.parse(e));const f=["sites.read.all"];let p;try{p=yield e.api(r).filter(s).middlewareOptions(Object(i.a)(...f)).get()}catch(e){}const _=yield K(e,p,f);return m()&&(yield l.putValue(r,{files:_.map(e=>JSON.stringify(e))})),_||null}),G=(e,t)=>d(void 0,void 0,void 0,function*(){if(!t||0===t.length)return[];const n=e.createBatch(),i=[],r=["files.read"];let s,c;p()&&(s=a.a.getCache(o.a.files,o.a.files.stores.fileQueries));for(const e of t)p()&&(c=yield s.getValue(e)),p()&&c&&f()>Date.now()-c.timeCached?i.push(JSON.parse(c.file)):""!==e&&n.get(e,e,r);try{const e=yield n.executeAll();for(const n of t){const t=e.get(n);(null==t?void 0:t.content)&&(i.push(t.content),p()&&(yield s.putValue(n,{file:JSON.stringify(t.content)})))}return i}catch(n){try{return Promise.all(t.filter(e=>e&&""!==e).map(t=>d(void 0,void 0,void 0,function*(){const n=yield _(e,t);if(n)return p()&&(yield s.putValue(t,{file:JSON.stringify(n)})),n})))}catch(e){return[]}}}),K=(e,t,n)=>d(void 0,void 0,void 0,function*(){if(!t)return[];const a=t.value,r=e.createBatch(),o=[];for(const e of a){const t=e.resourceReference.id;""!==t&&r.get(t,t,n)}try{const e=yield r.executeAll();for(const t of a){const n=e.get(t.resourceReference.id);(null==n?void 0:n.content)&&o.push(n.content)}return o}catch(t){try{return Promise.all(a.filter(e=>Boolean(e.resourceReference.id)).map(t=>d(void 0,void 0,void 0,function*(){return yield e.api(t.resourceReference.id).middlewareOptions(Object(i.a)(...n)).get()})))}catch(e){return[]}}}),W=(e,t)=>d(void 0,void 0,void 0,function*(){return r.a.create(e,t)}),q=(e,t,n)=>r.a.createFromValue(e,t.map(e=>JSON.parse(e)),n),Q=(e,t)=>d(void 0,void 0,void 0,function*(){if(p()){const n=yield e.getValue(t);if(n&&f()>Date.now()-n.timeCached)return JSON.parse(n.file)}return null}),Y=(e,t,n)=>d(void 0,void 0,void 0,function*(){if(e||(e=a.a.getCache(o.a.fileLists,t)),m()){const t=yield e.getValue(n);if(t&&(a.a.config.fileLists.invalidationPeriod||a.a.config.defaultInvalidationPeriod)>Date.now()-t.timeCached)return t}return null}),J=e=>d(void 0,void 0,void 0,function*(){const t=e.nextLink;if(e.hasNext&&(yield e.next()),m()){const n=a.a.getCache(o.a.fileLists,o.a.fileLists.stores.fileLists),i=/(graph.microsoft.com\/(v1.0|beta))(.*?)(?=\?)/gi.exec(t)[3];yield n.putValue(i,{files:e.value.map(e=>JSON.stringify(e)),nextLink:t})}}),X=(e,t,n)=>d(void 0,void 0,void 0,function*(){try{const a=yield e.api(t).responseType(s.ResponseType.RAW).middlewareOptions(Object(i.a)(...n)).get();return 404===a.status?{eTag:null,thumbnail:null}:a.ok?{eTag:a.headers.get("eTag"),thumbnail:yield Object(c.a)(yield a.blob())}:null}catch(e){return null}}),Z=(e,t)=>d(void 0,void 0,void 0,function*(){try{return(yield e.api(t).middlewareOptions(Object(i.a)("files.read")).get())||null}catch(e){}return null}),$=(e,t,n)=>d(void 0,void 0,void 0,function*(){try{const a="files.readwrite",r={item:{"@microsoft.graph.conflictBehavior":0===n||null===n?"rename":"replace"}};let o;try{o=yield e.api(t).middlewareOptions(Object(i.a)(a)).post(JSON.stringify(r))}catch(e){}return o||null}catch(e){return null}}),ee=(e,t,n,a,r)=>d(void 0,void 0,void 0,function*(){try{const o="files.readwrite",s={"Content-Length":n,"Content-Range":a};let c;try{c=yield e.client.api(t).middlewareOptions(Object(i.a)(o)).headers(s).put(r)}catch(e){}return c||null}catch(e){return null}}),te=(e,t,n)=>d(void 0,void 0,void 0,function*(){try{const a="files.readwrite";let r;try{r=yield e.client.api(t).middlewareOptions(Object(i.a)(a)).put(n)}catch(e){}return r||null}catch(e){return null}}),ne=(e,t)=>d(void 0,void 0,void 0,function*(){try{yield e.client.api(t).middlewareOptions(Object(i.a)("files.readwrite")).delete()}catch(e){return null}})},"0mOt":function(e,t,n){"use strict";n.d(t,"a",function(){return i}),n.d(t,"b",function(){return r});var a=n("HN6m");const i=e=>`${a.a.prefix}-${e}`,r=(e,t,n)=>{const a=i(e);customElements.get(a)||customElements.define(a,t,n)}},"0q6d":function(e,t,n){"use strict";function a(e,t,n){return isNaN(e)||e<=t?t:e>=n?n:e}function i(e,t,n){return isNaN(e)||e<=t?0:e>=n?1:e/(n-t)}function r(e,t,n){return isNaN(e)?t:t+e*(n-t)}function o(e){return e*(Math.PI/180)}function s(e){return e*(180/Math.PI)}function c(e){const t=Math.round(a(e,0,255)).toString(16);return 1===t.length?"0"+t:t}function d(e,t,n){return isNaN(e)||e<=0?t:e>=1?n:t+e*(n-t)}function l(e,t,n){if(e<=0)return t%360;if(e>=1)return n%360;const a=(t-n+360)%360;return a<=(n-t+360)%360?(t-a*e+360)%360:(t+a*e+360)%360}function u(e,t){const n=Math.pow(10,t);return Math.round(e*n)/n}n.d(t,"a",function(){return a}),n.d(t,"g",function(){return i}),n.d(t,"c",function(){return r}),n.d(t,"b",function(){return o}),n.d(t,"h",function(){return s}),n.d(t,"d",function(){return c}),n.d(t,"e",function(){return d}),n.d(t,"f",function(){return l}),n.d(t,"i",function(){return u}),Math.PI},"1/1t":function(e,t,n){"use strict";n.d(t,"b",function(){return p}),n.d(t,"a",function(){return m});var a=n("qqMp"),i=n("R2TB"),r=n("Yn9m"),o=n("h2QR"),s=n("z0DP"),c=n("C001");const d=[a.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host .loading,:host .no-data{margin:0 20px;display:flex;justify-content:center}:host .no-data{font-style:normal;font-weight:600;font-size:14px;color:var(--font-color,#323130);line-height:19px}:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{position:relative;user-select:none}:host .root .part{display:grid;grid-template-columns:auto 1fr auto}:host .root .part .part__icon{display:flex;min-width:20px;width:20px;height:20px;align-items:center;justify-content:center;margin-left:20px;margin-top:10px;line-height:20px}:host .root .part .part__icon svg{fill:var(--contact-copy-icon-color,var(--neutral-foreground-hint))}:host .root .part .part__details{margin:10px 14px;overflow:hidden}:host .root .part .part__details .part__title{font-size:12px;color:var(--contact-title-color,var(--neutral-foreground-hint));line-height:16px}:host .root .part .part__details .part__value{grid-column:2;color:var(--contact-value-color,var(--neutral-foreground-rest));font-size:14px;font-weight:400;line-height:19px}:host .root .part .part__details .part__value .part__link{color:var(--contact-link-color,var(--accent-foreground-rest));font-size:14px;cursor:pointer;text-overflow:ellipsis;overflow:hidden;white-space:nowrap;width:100%;display:inline-block}:host .root .part .part__details .part__value .part__link:hover{color:var(--contact-link-hover-color,var(--accent-foreground-hover))}:host .root .part .part__copy{width:32px;height:100%;background-color:var(--contact-background-color,transparent);visibility:hidden;display:flex;align-items:center;justify-content:flex-start}:host .root .part .part__copy svg{fill:var(--contact-copy-icon-color,var(--neutral-foreground-hint));cursor:pointer}:host .root .part:hover .part__copy{visibility:visible}:host .root.compact{padding:0}:host .root.compact .part{height:30px;align-items:center}:host .root.compact .part__details{margin:0}:host .root.compact .part__title{display:none}:host .root.compact .part__icon{margin-top:0;margin-right:6px;margin-bottom:2px}[dir=rtl] .part__link.phone{text-align:right;direction:ltr}[dir=rtl] .part__icon{margin:10px 20px 0 0!important}[dir=rtl].compact .part__icon{margin-left:6px!important;margin-top:0!important}@media (forced-colors:active) and (prefers-color-scheme:dark){.root svg{fill:#fff!important;fill-rule:nonzero!important;clip-rule:nonzero!important}.root svg path,.root svg rect{fill:#fff!important;fill-rule:nonzero!important;clip-rule:nonzero!important}}@media (forced-colors:active) and (prefers-color-scheme:light){.root svg{fill:#000!important;fill-rule:nonzero!important;clip-rule:nonzero!important}.root svg path,.root svg rect{fill:#000!important;fill-rule:nonzero!important;clip-rule:nonzero!important}}
`];var l=n("7Cdu");const u={contactSectionTitle:"Contact",emailTitle:"Email",chatTitle:"Teams",businessPhoneTitle:"Business Phone",cellPhoneTitle:"Mobile Phone",departmentTitle:"Department",titleTitle:"Title",officeLocationTitle:"Office Location",copyToClipboardButton:"Copy to clipboard"};var f=n("0mOt");const p=()=>{Object(f.b)("contact",m)};class m extends c.a{static get styles(){return d}get strings(){return u}get hasData(){return!!this._contactParts&&!!Object.values(this._contactParts).filter(e=>!!e.value).length}constructor(e){var t;super(),this._contactParts={email:{icon:Object(l.b)(l.a.Email),onClick:()=>this.sendEmail(Object(s.e)(this._person)),showCompact:!0,title:this.strings.emailTitle},chat:{icon:Object(l.b)(l.a.Chat),onClick:()=>{var e;return this.sendChat(null===(e=this._person)||void 0===e?void 0:e.userPrincipalName)},showCompact:!1,title:this.strings.chatTitle},businessPhone:{icon:Object(l.b)(l.a.Phone),onClick:()=>{var e,t;return this.sendCall((null===(t=null===(e=this._person)||void 0===e?void 0:e.businessPhones)||void 0===t?void 0:t.length)>0?this._person.businessPhones[0]:null)},showCompact:!0,title:this.strings.businessPhoneTitle},cellPhone:{icon:Object(l.b)(l.a.CellPhone),onClick:()=>{var e;return this.sendCall(null===(e=this._person)||void 0===e?void 0:e.mobilePhone)},showCompact:!0,title:this.strings.cellPhoneTitle},department:{icon:Object(l.b)(l.a.Department),showCompact:!1,title:this.strings.departmentTitle},title:{icon:Object(l.b)(l.a.Person),showCompact:!1,title:this.strings.titleTitle},officeLocation:{icon:Object(l.b)(l.a.OfficeLocation),showCompact:!0,title:this.strings.officeLocationTitle}},this.sendCall=e=>{this.sendLink("tel:",e)},this._person=e,this._contactParts.email.value=Object(s.e)(this._person),this._contactParts.chat.value=this._person.userPrincipalName,this._contactParts.cellPhone.value=this._person.mobilePhone,this._contactParts.department.value=this._person.department,this._contactParts.title.value=this._person.jobTitle,this._contactParts.officeLocation.value=this._person.officeLocation,(null===(t=this._person.businessPhones)||void 0===t?void 0:t.length)&&(this._contactParts.businessPhone.value=this._person.businessPhones[0])}get displayName(){return this.strings.contactSectionTitle}get cardTitle(){return this.strings.contactSectionTitle}renderIcon(){return Object(l.b)(l.a.Contact)}clearState(){super.clearState();for(const e of Object.keys(this._contactParts))this._contactParts[e].value=null}renderCompactView(){if(!this.hasData)return null;const e=Object.values(this._contactParts).filter(e=>!!e.value);let t=Object.values(e).filter(e=>!!e.value&&e.showCompact);(null==t?void 0:t.length)||(t=Object.values(e).slice(0,2));const n=a.c`
      ${t.map(e=>this.renderContactPart(e))}
    `;return a.c`
      <div class="root compact" dir=${this.direction}>
        ${n}
      </div>
    `}renderFullView(){let e;if(this.hasData){const t=Object.values(this._contactParts).filter(e=>!!e.value);e=a.c`
        ${t.map(e=>this.renderContactPart(e))}
      `}return a.c`
      <div class="root" dir=${this.direction}>
        ${e}
      </div>
    `}renderContactPart(e){let t=!1;"Mobile Phone"!==e.title&&"Business Phone"!==e.title||(t=!0);const n={part__link:!0,phone:t},i=e.onClick?a.c`
          <span class=${Object(o.a)(n)} @click=${t=>e.onClick(t)}>${e.value}</span>
        `:a.c`
          ${e.value}
        `;return a.c`
      <div class="part" role="button" @click=${t=>this.handlePartClick(t,e.value)} tabindex="0">
        <div class="part__icon" aria-label=${e.title} title=${e.title}>${e.icon}</div>
        <div class="part__details">
          <div class="part__title">${e.title}</div>
          <div class="part__value" title=${e.title}>${i}</div>
        </div>
        <div
          class="part__copy"
          aria-label=${this.strings.copyToClipboardButton}
          title=${this.strings.copyToClipboardButton}
        >
          ${Object(l.b)(l.a.Copy)}
        </div>
      </div>
    `}handlePartClick(e,t){t&&navigator.clipboard.writeText(t)}sendLink(e,t){t?window.open(`${e}${t}`,"_blank","noreferrer"):Object(i.a)(`Target resource for ${e} link was not provided: resource: ${t}`)}sendChat(e){if(!e)return void Object(i.a)(" Can't send chat when upn is not provided");const t=`https://teams.microsoft.com/l/chat/0/0?users=${e}`,n=()=>window.open(t,"_blank","noreferrer");r.a.isAvailable?r.a.executeDeepLink(t,e=>{e||n()}):n()}sendEmail(e){this.sendLink("mailto:",e)}}},"2i1/":function(e,t,n){"use strict";n.d(t,"a",function(){return a});const a="not-allowed"},"2oY9":function(e,t,n){"use strict";n.d(t,"a",function(){return a});const a="3.1.3"},"3ATu":function(e,t,n){"use strict";n.d(t,"b",function(){return d}),n.d(t,"a",function(){return l});var a=n("qqMp"),i=n("C001"),r=n("7Cdu");const o=[a.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host .loading,:host .no-data{margin:0 20px;display:flex;justify-content:center}:host .no-data{font-style:normal;font-weight:600;font-size:14px;color:var(--font-color,#323130);line-height:19px}:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{position:relative;user-select:none;background-color:var(--profile-background-color,--neutral-layer-1)}:host .root{padding:20px 0}:host .root.compact{padding:0}:host .root .title{font-size:14px;font-weight:600;color:var(--profile-title-color,var(--neutral-foreground-rest));margin:0 20px 12px}:host .root section{margin-bottom:24px;padding:0 20px}:host .root section:last-child{margin:0}:host .root section .section__title{font-size:14px;color:var(--profile-section-title-color,var(--neutral-foreground-hint))}:host .root section .section__content{display:flex;flex-direction:column;margin-top:10px}:host .root .token-list{display:flex;flex-flow:row wrap;margin-top:-10px}:host .root .token-list .token-list__item{text-overflow:ellipsis;white-space:nowrap;display:inline-block;overflow:hidden;font-size:14px;align-items:center;background:var(--profile-token-item-background-color,var(--neutral-fill-secondary-rest));border-radius:2px;max-height:28px;padding:4px 8px;margin-right:10px;margin-top:10px;color:var(--profile-token-item-color,var(--neutral-foreground-rest))}:host .root .token-list .token-list__item:last-child{margin-right:initial}:host .root .token-list .token-list__item.overflow{display:none}:host .root .token-list .token-list__item.token-list__item--show-overflow{cursor:pointer;user-select:unset;background:0 0;color:var(--profile-token-overflow-color,var(--accent-foreground-rest))}:host .root .data-list__item{margin-bottom:20px}:host .root .data-list__item:last-child{margin-bottom:initial}:host .root .data-list__item .data-list__item__header{display:flex;justify-content:space-between;align-items:center}:host .root .data-list__item .data-list__item__content{font-size:12px;line-height:16px;color:var(--profile-section-title-color,var(--neutral-foreground-hint));margin-top:4px}:host .root .data-list__item .data-list__item__title{font-size:14px;line-height:16px;color:var(--profile-title-color,var(--neutral-foreground-rest))}:host .root .data-list__item .data-list__item__date-range{color:var(--profile-section-title-color,var(--neutral-foreground-hint));font-size:10px;line-height:12px}:host .root .language__proficiency{opacity:.77}:host .root .work-position .work-position__company{color:#023b8f}:host .root .work-position .work-position__location{font-size:10px;color:var(--profile-section-title-color,var(--neutral-foreground-hint));line-height:16px}:host .root .educational-activity .educational-activity__degree{font-size:12px;line-height:14px;color:var(--profile-section-title-color,var(--neutral-foreground-hint))}:host .root .birthday{display:flex;align-items:center;margin-top:-6px}:host .root .birthday .birthday__icon{margin-right:8px}:host .root .birthday .birthday__date{font-size:12px;color:var(--profile-title-color,var(--neutral-foreground-rest))}[dir=rtl] .token-list__item{margin-right:0!important}
`],s={SkillsAndExperienceSectionTitle:"Skills & Experience",AboutCompactSectionTitle:"About",SkillsSubSectionTitle:"Skills",LanguagesSubSectionTitle:"Languages",WorkExperienceSubSectionTitle:"Work Experience",EducationSubSectionTitle:"Education",professionalInterestsSubSectionTitle:"Professional Interests",personalInterestsSubSectionTitle:"Personal Interests",birthdaySubSectionTitle:"Birthday",currentYearSubtitle:"Current",socialMediaSubSectionTitle:"Social Media"};var c=n("0mOt");const d=()=>Object(c.b)("profile",l);class l extends i.a{static get styles(){return o}get strings(){return s}get displayName(){return this.strings.SkillsAndExperienceSectionTitle}get cardTitle(){return this.strings.AboutCompactSectionTitle}get hasData(){var e,t;if(!this.profile)return!1;const{languages:n,skills:a,positions:i,educationalActivities:r}=this.profile;return[this._birthdayAnniversary,null===(e=this._personalInterests)||void 0===e?void 0:e.length,null===(t=this._professionalInterests)||void 0===t?void 0:t.length,null==n?void 0:n.length,null==a?void 0:a.length,null==i?void 0:i.length,null==r?void 0:r.length].filter(e=>!!e).length>0}get profile(){return this._profile}set profile(e){e!==this._profile&&(this._profile=e,this._birthdayAnniversary=(null==e?void 0:e.anniversaries)?e.anniversaries.find(this.isBirthdayAnniversary):null,this._personalInterests=(null==e?void 0:e.interests)?e.interests.filter(this.isPersonalInterest):null,this._professionalInterests=(null==e?void 0:e.interests)?e.interests.filter(this.isProfessionalInterest):null)}constructor(e){super(),this.isPersonalInterest=e=>{var t;return null===(t=e.categories)||void 0===t?void 0:t.includes("personal")},this.isProfessionalInterest=e=>{var t;return null===(t=e.categories)||void 0===t?void 0:t.includes("professional")},this.isBirthdayAnniversary=e=>"birthday"===e.type,this.profile=e}renderIcon(){return Object(r.b)(r.a.Profile)}clearState(){super.clearState(),this.profile=null}renderCompactView(){return a.c`
       <div class="root compact" dir=${this.direction}>
         ${this.renderSubSections().slice(0,2)}
       </div>
     `}renderFullView(){return this.initPostRenderOperations(),a.c`
       <div class="root" dir=${this.direction}>
         ${this.renderSubSections()}
       </div>
     `}renderSubSections(){return[this.renderSkills(),this.renderBirthday(),this.renderLanguages(),this.renderWorkExperience(),this.renderEducation(),this.renderProfessionalInterests(),this.renderPersonalInterests()].filter(e=>!!e)}renderLanguages(){var e;const{languages:t}=this._profile;if(!(null==t?void 0:t.length))return null;const n=[];for(const i of t){let t=null;(null===(e=i.proficiency)||void 0===e?void 0:e.length)&&(t=a.c`
           <span class="language__proficiency" tabindex="0">
             &nbsp;(${i.proficiency})
           </span>
         `),n.push(a.c`
         <div class="token-list__item language">
           <span class="language__title" tabindex="0">${i.displayName}</span>
           ${t}
         </div>
       `)}const i=n.length?this.strings.LanguagesSubSectionTitle:"";return a.c`
       <section>
         <div class="section__title" tabindex="0">${i}</div>
         <div class="section__content">
           <div class="token-list">
             ${n}
           </div>
         </div>
       </section>
     `}renderSkills(){const{skills:e}=this._profile;if(!(null==e?void 0:e.length))return null;const t=[];for(const n of e)t.push(a.c`
         <div class="token-list__item skill" tabindex="0">
           ${n.displayName}
         </div>
       `);const n=t.length?this.strings.SkillsSubSectionTitle:"";return a.c`
       <section>
         <div class="section__title" tabindex="0">${n}</div>
         <div class="section__content">
           <div class="token-list">
             ${t}
           </div>
         </div>
       </section>
     `}renderWorkExperience(){var e,t,n,i,r;const{positions:o}=this._profile;if(!(null==o?void 0:o.length))return null;const s=[];for(const o of this._profile.positions)(o.detail.description||""!==o.detail.jobTitle)&&s.push(a.c`
           <div class="data-list__item work-position">
             <div class="data-list__item__header">
               <div class="data-list__item__title" tabindex="0">${null===(e=o.detail)||void 0===e?void 0:e.jobTitle}</div>
               <div class="data-list__item__date-range" tabindex="0">
                 ${this.getDisplayDateRange(o.detail)}
               </div>
             </div>
             <div class="data-list__item__content">
               <div class="work-position__company" tabindex="0">
                 ${null===(n=null===(t=null==o?void 0:o.detail)||void 0===t?void 0:t.company)||void 0===n?void 0:n.displayName}
               </div>
               <div class="work-position__location" tabindex="0">
                 ${this.displayLocation(null===(r=null===(i=null==o?void 0:o.detail)||void 0===i?void 0:i.company)||void 0===r?void 0:r.address)}
               </div>
             </div>
           </div>
         `);const c=s.length?this.strings.WorkExperienceSubSectionTitle:"";return a.c`
       <section>
         <div class="section__title" tabindex="0">${c}</div>
         <div class="section__content">
           <div class="data-list">
             ${s}
           </div>
         </div>
       </section>
     `}renderEducation(){const{educationalActivities:e}=this._profile;if(!(null==e?void 0:e.length))return null;const t=[];for(const n of e)t.push(a.c`
         <div class="data-list__item educational-activity">
           <div class="data-list__item__header">
             <div class="data-list__item__title" tabindex="0">${n.institution.displayName}</div>
             <div class="data-list__item__date-range" tabindex="0">
               ${this.getDisplayDateRange(n)}
             </div>
           </div>
           ${n.program.displayName?a.c`<div class="data-list__item__content">
                  <div class="educational-activity__degree" tabindex="0">
                  ${n.program.displayName}
                </div>`:a.d}
         </div>
       `);const n=t.length?this.strings.EducationSubSectionTitle:"";return a.c`
       <section>
         <div class="section__title" tabindex="0">${n}</div>
         <div class="section__content">
           <div class="data-list">
             ${t}
           </div>
         </div>
       </section>
     `}renderProfessionalInterests(){var e;if(!(null===(e=this._professionalInterests)||void 0===e?void 0:e.length))return null;const t=[];for(const e of this._professionalInterests)t.push(a.c`
         <div class="token-list__item interest interest--professional" tabindex="0">
           ${e.displayName}
         </div>
       `);const n=t.length?this.strings.professionalInterestsSubSectionTitle:"";return a.c`
       <section>
         <div class="section__title" tabindex="0">${n}</div>
         <div class="section__content">
           <div class="token-list">
             ${t}
           </div>
         </div>
       </section>
     `}renderPersonalInterests(){var e;if(!(null===(e=this._personalInterests)||void 0===e?void 0:e.length))return null;const t=[];for(const e of this._personalInterests)t.push(a.c`
         <div class="token-list__item interest interest--personal" tabindex="0">
           ${e.displayName}
         </div>
       `);const n=t.length?this.strings.personalInterestsSubSectionTitle:"";return a.c`
       <section>
         <div class="section__title" tabindex="0">${n}</div>
         <div class="section__content">
           <div class="token-list">
             ${t}
           </div>
         </div>
       </section>
     `}renderBirthday(){var e;return(null===(e=this._birthdayAnniversary)||void 0===e?void 0:e.date)?a.c`
       <section>
         <div class="section__title" tabindex="0">Birthday</div>
         <div class="section__content">
           <div class="birthday">
             <div class="birthday__icon">
               ${Object(r.b)(r.a.Birthday)}
             </div>
             <div class="birthday__date" tabindex="0">
               ${this.getDisplayDate(new Date(this._birthdayAnniversary.date))}
             </div>
           </div>
         </div>
       </section>
     `:null}getDisplayDate(e){return e.toLocaleString("default",{day:"numeric",month:"long"})}getDisplayDateRange(e){if(!e.startMonthYear)return a.d;const t=new Date(e.startMonthYear).getFullYear();return 0===t||1===t?a.d:`${t} — ${e.endMonthYear?new Date(e.endMonthYear).getFullYear():this.strings.currentYearSubtitle}`}displayLocation(e){return(null==e?void 0:e.city)?e.state?`${e.city}, ${e.state}`:e.city:a.d}initPostRenderOperations(){setTimeout(()=>{try{this.shadowRoot.querySelectorAll("section").forEach(e=>{this.handleTokenOverflow(e)})}catch(e){}},0)}handleTokenOverflow(e){const t=e.querySelectorAll(".token-list");if(null==t?void 0:t.length)for(const e of Array.from(t)){const t=e.querySelectorAll(".token-list__item");if(!(null==t?void 0:t.length))continue;let n=null,a=t[0].getBoundingClientRect();const i=e.getBoundingClientRect(),r=2*a.height+i.top;for(let e=0;e<t.length-1;e++)if(a=t[e].getBoundingClientRect(),a.top>r){n=Array.from(t).slice(e,t.length);break}if(n){n.forEach(e=>e.classList.add("overflow"));const t=document.createElement("div");t.classList.add("token-list__item"),t.classList.add("token-list__item--show-overflow"),t.tabIndex=0,t.innerText=`+ ${n.length} more`;const a=()=>{t.remove(),n.forEach(e=>e.classList.remove("overflow"))};t.addEventListener("click",()=>{a()}),t.addEventListener("keydown",e=>{"Enter"===e.code&&a()}),e.appendChild(t)}}}}},"3rEL":function(e,t,n){"use strict";function a(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o}n.d(t,"a",function(){return a}),Object.create,Object.create,"function"==typeof SuppressedError&&SuppressedError},"4Eu1":function(e,t,n){"use strict";n.d(t,"a",function(){return a});class a{}},"4X57":function(e,t,n){"use strict";n.d(t,"a",function(){return o}),n.d(t,"b",function(){return c});var a=n("olMv"),i=n("+yEz");function r(e,t){const n=[];let r="";const o=[];for(let s=0,c=e.length-1;s<c;++s){r+=e[s];let c=t[s];if(c instanceof a.a){const e=c.createBehavior();c=c.createCSS(),e&&o.push(e)}c instanceof i.a||c instanceof CSSStyleSheet?(""!==r.trim()&&(n.push(r),r=""),n.push(c)):r+=c}return r+=e[e.length-1],""!==r.trim()&&n.push(r),{styles:n,behaviors:o}}function o(e,...t){const{styles:n,behaviors:a}=r(e,t),o=i.a.create(n);return a.length&&o.withBehaviors(...a),o}class s extends a.a{constructor(e,t){super(),this.behaviors=t,this.css="";const n=e.reduce((e,t)=>("string"==typeof t?this.css+=t:e.push(t),e),[]);n.length&&(this.styles=i.a.create(n))}createBehavior(){return this}createCSS(){return this.css}bind(e){this.styles&&e.$fastController.addStyles(this.styles),this.behaviors.length&&e.$fastController.addBehaviors(this.behaviors)}unbind(e){this.styles&&e.$fastController.removeStyles(this.styles),this.behaviors.length&&e.$fastController.removeBehaviors(this.behaviors)}}function c(e,...t){const{styles:n,behaviors:a}=r(e,t);return new s(n,a)}},"4gKT":function(e,t,n){"use strict";n.d(t,"a",function(){return s}),n.d(t,"b",function(){return c});var a=n("+yEz"),i=n("q5bN");function r(e){return`${e.toLowerCase()}:presentation`}const o=new Map,s=Object.freeze({define(e,t,n){const a=r(e);void 0===o.get(a)?o.set(a,t):o.set(a,!1),n.register(i.b.instance(a,t))},forTag(e,t){const n=r(e),a=o.get(n);return!1===a?i.a.findResponsibleContainer(t).get(n):a||null}});class c{constructor(e,t){this.template=e||null,this.styles=void 0===t?null:Array.isArray(t)?a.a.create(t):t instanceof a.a?t:a.a.create([t])}applyTo(e){const t=e.$fastController;null===t.template&&(t.template=this.template),null===t.styles&&(t.styles=this.styles)}}},"57ob":function(e,t,n){"use strict";n.d(t,"a",function(){return r});var a=n("oePG");class i{constructor(e,t,n){this.propertyName=e,this.value=t,this.styles=n}bind(e){a.b.getNotifier(e).subscribe(this,this.propertyName),this.handleChange(e,this.propertyName)}unbind(e){a.b.getNotifier(e).unsubscribe(this,this.propertyName),e.$fastController.removeStyles(this.styles)}handleChange(e,t){e[t]===this.value?e.$fastController.addStyles(this.styles):e.$fastController.removeStyles(this.styles)}}function r(e,t){return new i("appearance",e,t)}},"5ZAu":function(e,t,n){"use strict";n.d(t,"c",function(){return c}),n.d(t,"b",function(){return d}),n.d(t,"a",function(){return l});var a=n("oZuh");const i=a.a.FAST.getById(1,()=>{const e=[],t=[];function n(){if(t.length)throw t.shift()}function i(e){try{e.call()}catch(e){t.push(e),setTimeout(n,0)}}function r(){let t=0;for(;t<e.length;)if(i(e[t]),t++,t>1024){for(let n=0,a=e.length-t;n<a;n++)e[n]=e[n+t];e.length-=t,t=0}e.length=0}return Object.freeze({enqueue:function(t){e.length<1&&a.a.requestAnimationFrame(r),e.push(t)},process:r})}),r=a.a.trustedTypes.createPolicy("fast-html",{createHTML:e=>e});let o=r;const s=`fast-${Math.random().toString(36).substring(2,8)}`,c=`${s}{`,d=`}${s}`,l=Object.freeze({supportsAdoptedStyleSheets:Array.isArray(document.adoptedStyleSheets)&&"replace"in CSSStyleSheet.prototype,setHTMLPolicy(e){if(o!==r)throw new Error("The HTML policy can only be set once.");o=e},createHTML:e=>o.createHTML(e),isMarker:e=>e&&8===e.nodeType&&e.data.startsWith(s),extractDirectiveIndexFromMarker:e=>parseInt(e.data.replace(`${s}:`,"")),createInterpolationPlaceholder:e=>`${c}${e}${d}`,createCustomAttributePlaceholder(e,t){return`${e}="${this.createInterpolationPlaceholder(t)}"`},createBlockPlaceholder:e=>`\x3c!--${s}:${e}--\x3e`,queueUpdate:i.enqueue,processUpdates:i.process,nextUpdate:()=>new Promise(i.enqueue),setAttribute(e,t,n){null==n?e.removeAttribute(t):e.setAttribute(t,n)},setBooleanAttribute(e,t,n){n?e.setAttribute(t,""):e.removeAttribute(t)},removeChildNodes(e){for(let t=e.firstChild;null!==t;t=e.firstChild)e.removeChild(t)},createTemplateWalker:e=>document.createTreeWalker(e,133,null,!1)})},"5rba":function(e,t,n){"use strict";var a;n.d(t,"a",function(){return j}),n.d(t,"b",function(){return C}),n.d(t,"c",function(){return O}),n.d(t,"d",function(){return w}),n.d(t,"e",function(){return z});const i=globalThis,r=i.trustedTypes,o=r?r.createPolicy("lit-html",{createHTML:e=>e}):void 0,s="$lit$",c=`lit$${(Math.random()+"").slice(9)}$`,d="?"+c,l=`<${d}>`,u=document,f=()=>u.createComment(""),p=e=>null===e||"object"!=typeof e&&"function"!=typeof e,m=Array.isArray,_=e=>m(e)||"function"==typeof(null==e?void 0:e[Symbol.iterator]),h="[ \t\n\f\r]",b=/<(?:(!--|\/[^a-zA-Z])|(\/?[a-zA-Z][^>\s]*)|(\/?$))/g,g=/-->/g,v=/>/g,y=RegExp(`>|${h}(?:([^\\s"'>=/]+)(${h}*=${h}*(?:[^ \t\n\f\r"'\`<>=]|("|')|))|$)`,"g"),S=/'/g,D=/"/g,I=/^(?:script|style|textarea|title)$/i,x=e=>(t,...n)=>({_$litType$:e,strings:t,values:n}),C=x(1),O=(x(2),Symbol.for("lit-noChange")),w=Symbol.for("lit-nothing"),E=new WeakMap,A=u.createTreeWalker(u,129);function L(e,t){if(!Array.isArray(e)||!e.hasOwnProperty("raw"))throw Error("invalid template strings array");return void 0!==o?o.createHTML(t):t}const k=(e,t)=>{const n=e.length-1,a=[];let i,r=2===t?"<svg>":"",o=b;for(let t=0;t<n;t++){const n=e[t];let u,f,p=-1,m=0;for(;m<n.length&&(o.lastIndex=m,f=o.exec(n),null!==f);){var d;m=o.lastIndex,o===b?"!--"===f[1]?o=g:void 0!==f[1]?o=v:void 0!==f[2]?(I.test(f[2])&&(i=RegExp("</"+f[2],"g")),o=y):void 0!==f[3]&&(o=y):o===y?">"===f[0]?(o=null!==(d=i)&&void 0!==d?d:b,p=-1):void 0===f[1]?p=-2:(p=o.lastIndex-f[2].length,u=f[1],o=void 0===f[3]?y:'"'===f[3]?D:S):o===D||o===S?o=y:o===g||o===v?o=b:(o=y,i=void 0)}const _=o===y&&e[t+1].startsWith("/>")?" ":"";r+=o===b?n+l:p>=0?(a.push(u),n.slice(0,p)+s+n.slice(p)+c+_):n+c+(-2===p?t:_)}return[L(e,r+(e[n]||"<?>")+(2===t?"</svg>":"")),a]};class M{constructor({strings:e,_$litType$:t},n){let a;this.parts=[];let i=0,o=0;const l=e.length-1,u=this.parts,[p,m]=k(e,t);if(this.el=M.createElement(p,n),A.currentNode=this.el.content,2===t){const e=this.el.content.firstChild;e.replaceWith(...e.childNodes)}for(;null!==(a=A.nextNode())&&u.length<l;){if(1===a.nodeType){if(a.hasAttributes())for(const e of a.getAttributeNames())if(e.endsWith(s)){const t=m[o++],n=a.getAttribute(e).split(c),r=/([.?@])?(.*)/.exec(t);u.push({type:1,index:i,name:r[2],strings:n,ctor:"."===r[1]?H:"?"===r[1]?R:"@"===r[1]?N:F}),a.removeAttribute(e)}else e.startsWith(c)&&(u.push({type:6,index:i}),a.removeAttribute(e));if(I.test(a.tagName)){const e=a.textContent.split(c),t=e.length-1;if(t>0){a.textContent=r?r.emptyScript:"";for(let n=0;n<t;n++)a.append(e[n],f()),A.nextNode(),u.push({type:2,index:++i});a.append(e[t],f())}}}else if(8===a.nodeType)if(a.data===d)u.push({type:2,index:i});else{let e=-1;for(;-1!==(e=a.data.indexOf(c,e+1));)u.push({type:7,index:i}),e+=c.length-1}i++}}static createElement(e,t){const n=u.createElement("template");return n.innerHTML=e,n}}function P(e,t,n=e,a){var i,r,o,s,c;if(t===O)return t;let d=void 0!==a?null===(i=n._$Co)||void 0===i?void 0:i[a]:n._$Cl;const l=p(t)?void 0:t._$litDirective$;return(null===(r=d)||void 0===r?void 0:r.constructor)!==l&&(null!==(o=d)&&void 0!==o&&null!==(s=o._$AO)&&void 0!==s&&s.call(o,!1),void 0===l?d=void 0:(d=new l(e),d._$AT(e,n,a)),void 0!==a?(null!==(c=n._$Co)&&void 0!==c?c:n._$Co=[])[a]=d:n._$Cl=d),void 0!==d&&(t=P(e,d._$AS(e,t.values),d,a)),t}class T{constructor(e,t){this._$AV=[],this._$AN=void 0,this._$AD=e,this._$AM=t}get parentNode(){return this._$AM.parentNode}get _$AU(){return this._$AM._$AU}u(e){var t;const{el:{content:n},parts:a}=this._$AD,i=(null!==(t=null==e?void 0:e.creationScope)&&void 0!==t?t:u).importNode(n,!0);A.currentNode=i;let r=A.nextNode(),o=0,s=0,c=a[0];for(;void 0!==c;){var d;if(o===c.index){let t;2===c.type?t=new U(r,r.nextSibling,this,e):1===c.type?t=new c.ctor(r,c.name,c.strings,this,e):6===c.type&&(t=new B(r,this,e)),this._$AV.push(t),c=a[++s]}o!==(null===(d=c)||void 0===d?void 0:d.index)&&(r=A.nextNode(),o++)}return A.currentNode=u,i}p(e){let t=0;for(const n of this._$AV)void 0!==n&&(void 0!==n.strings?(n._$AI(e,n,t),t+=n.strings.length-2):n._$AI(e[t])),t++}}class U{get _$AU(){var e,t;return null!==(e=null===(t=this._$AM)||void 0===t?void 0:t._$AU)&&void 0!==e?e:this._$Cv}constructor(e,t,n,a){var i;this.type=2,this._$AH=w,this._$AN=void 0,this._$AA=e,this._$AB=t,this._$AM=n,this.options=a,this._$Cv=null===(i=null==a?void 0:a.isConnected)||void 0===i||i}get parentNode(){var e;let t=this._$AA.parentNode;const n=this._$AM;return void 0!==n&&11===(null===(e=t)||void 0===e?void 0:e.nodeType)&&(t=n.parentNode),t}get startNode(){return this._$AA}get endNode(){return this._$AB}_$AI(e,t=this){e=P(this,e,t),p(e)?e===w||null==e||""===e?(this._$AH!==w&&this._$AR(),this._$AH=w):e!==this._$AH&&e!==O&&this._(e):void 0!==e._$litType$?this.g(e):void 0!==e.nodeType?this.$(e):_(e)?this.T(e):this._(e)}k(e){return this._$AA.parentNode.insertBefore(e,this._$AB)}$(e){this._$AH!==e&&(this._$AR(),this._$AH=this.k(e))}_(e){this._$AH!==w&&p(this._$AH)?this._$AA.nextSibling.data=e:this.$(u.createTextNode(e)),this._$AH=e}g(e){var t;const{values:n,_$litType$:a}=e,i="number"==typeof a?this._$AC(e):(void 0===a.el&&(a.el=M.createElement(L(a.h,a.h[0]),this.options)),a);if((null===(t=this._$AH)||void 0===t?void 0:t._$AD)===i)this._$AH.p(n);else{const e=new T(i,this),t=e.u(this.options);e.p(n),this.$(t),this._$AH=e}}_$AC(e){let t=E.get(e.strings);return void 0===t&&E.set(e.strings,t=new M(e)),t}T(e){m(this._$AH)||(this._$AH=[],this._$AR());const t=this._$AH;let n,a=0;for(const i of e)a===t.length?t.push(n=new U(this.k(f()),this.k(f()),this,this.options)):n=t[a],n._$AI(i),a++;a<t.length&&(this._$AR(n&&n._$AB.nextSibling,a),t.length=a)}_$AR(e=this._$AA.nextSibling,t){for(null===(n=this._$AP)||void 0===n||n.call(this,!1,!0,t);e&&e!==this._$AB;){var n;const t=e.nextSibling;e.remove(),e=t}}setConnected(e){var t;void 0===this._$AM&&(this._$Cv=e,null===(t=this._$AP)||void 0===t||t.call(this,e))}}class F{get tagName(){return this.element.tagName}get _$AU(){return this._$AM._$AU}constructor(e,t,n,a,i){this.type=1,this._$AH=w,this._$AN=void 0,this.element=e,this.name=t,this._$AM=a,this.options=i,n.length>2||""!==n[0]||""!==n[1]?(this._$AH=Array(n.length-1).fill(new String),this.strings=n):this._$AH=w}_$AI(e,t=this,n,a){const i=this.strings;let r=!1;if(void 0===i)e=P(this,e,t,0),r=!p(e)||e!==this._$AH&&e!==O,r&&(this._$AH=e);else{const a=e;let s,c;for(e=i[0],s=0;s<i.length-1;s++){var o;c=P(this,a[n+s],t,s),c===O&&(c=this._$AH[s]),r||(r=!p(c)||c!==this._$AH[s]),c===w?e=w:e!==w&&(e+=(null!==(o=c)&&void 0!==o?o:"")+i[s+1]),this._$AH[s]=c}}r&&!a&&this.j(e)}j(e){e===w?this.element.removeAttribute(this.name):this.element.setAttribute(this.name,null!=e?e:"")}}class H extends F{constructor(){super(...arguments),this.type=3}j(e){this.element[this.name]=e===w?void 0:e}}class R extends F{constructor(){super(...arguments),this.type=4}j(e){this.element.toggleAttribute(this.name,!!e&&e!==w)}}class N extends F{constructor(e,t,n,a,i){super(e,t,n,a,i),this.type=5}_$AI(e,t=this){var n;if((e=null!==(n=P(this,e,t,0))&&void 0!==n?n:w)===O)return;const a=this._$AH,i=e===w&&a!==w||e.capture!==a.capture||e.once!==a.once||e.passive!==a.passive,r=e!==w&&(a===w||i);i&&this.element.removeEventListener(this.name,this,a),r&&this.element.addEventListener(this.name,this,e),this._$AH=e}handleEvent(e){var t,n;"function"==typeof this._$AH?this._$AH.call(null!==(t=null===(n=this.options)||void 0===n?void 0:n.host)&&void 0!==t?t:this.element,e):this._$AH.handleEvent(e)}}class B{constructor(e,t,n){this.element=e,this.type=6,this._$AN=void 0,this._$AM=t,this.options=n}get _$AU(){return this._$AM._$AU}_$AI(e){P(this,e)}}const j={S:s,A:c,P:d,C:1,M:k,L:T,R:_,V:P,D:U,I:F,H:R,N:N,U:H,B:B},V=i.litHtmlPolyfillSupport;null!=V&&V(M,U),(null!==(a=i.litHtmlVersions)&&void 0!==a?a:i.litHtmlVersions=[]).push("3.0.0");const z=(e,t,n)=>{var a;const i=null!==(a=null==n?void 0:n.renderBefore)&&void 0!==a?a:t;let r=i._$litPart$;if(void 0===r){var o;const e=null!==(o=null==n?void 0:n.renderBefore)&&void 0!==o?o:null;i._$litPart$=r=new U(t.insertBefore(f(),e),e,void 0,null!=n?n:{})}return r._$AI(e),r}},"6BDD":function(e,t,n){"use strict";n.d(t,"a",function(){return E});var a=n("5ZAu"),i=n("oePG"),r=n("/w5G");function o(e,t){this.source=e,this.context=t,null===this.bindingObserver&&(this.bindingObserver=i.b.binding(this.binding,this,this.isBindingVolatile)),this.updateTarget(this.bindingObserver.observe(e,t))}function s(e,t){this.source=e,this.context=t,this.target.addEventListener(this.targetName,this)}function c(){this.bindingObserver.disconnect(),this.source=null,this.context=null}function d(){this.bindingObserver.disconnect(),this.source=null,this.context=null;const e=this.target.$fastView;void 0!==e&&e.isComposed&&(e.unbind(),e.needsBindOnly=!0)}function l(){this.target.removeEventListener(this.targetName,this),this.source=null,this.context=null}function u(e){a.a.setAttribute(this.target,this.targetName,e)}function f(e){a.a.setBooleanAttribute(this.target,this.targetName,e)}function p(e){if(null==e&&(e=""),e.create){this.target.textContent="";let t=this.target.$fastView;void 0===t?t=e.create():this.target.$fastTemplate!==e&&(t.isComposed&&(t.remove(),t.unbind()),t=e.create()),t.isComposed?t.needsBindOnly&&(t.needsBindOnly=!1,t.bind(this.source,this.context)):(t.isComposed=!0,t.bind(this.source,this.context),t.insertBefore(this.target),this.target.$fastView=t,this.target.$fastTemplate=e)}else{const t=this.target.$fastView;void 0!==t&&t.isComposed&&(t.isComposed=!1,t.remove(),t.needsBindOnly?t.needsBindOnly=!1:t.unbind()),this.target.textContent=e}}function m(e){this.target[this.targetName]=e}function _(e){const t=this.classVersions||Object.create(null),n=this.target;let a=this.version||0;if(null!=e&&e.length){const i=e.split(/\s+/);for(let e=0,r=i.length;e<r;++e){const r=i[e];""!==r&&(t[r]=a,n.classList.add(r))}}if(this.classVersions=t,this.version=a+1,0!==a){a-=1;for(const e in t)t[e]===a&&n.classList.remove(e)}}class h extends r.c{constructor(e){super(),this.binding=e,this.bind=o,this.unbind=c,this.updateTarget=u,this.isBindingVolatile=i.b.isVolatileBinding(this.binding)}get targetName(){return this.originalTargetName}set targetName(e){if(this.originalTargetName=e,void 0!==e)switch(e[0]){case":":if(this.cleanedTargetName=e.substr(1),this.updateTarget=m,"innerHTML"===this.cleanedTargetName){const e=this.binding;this.binding=(t,n)=>a.a.createHTML(e(t,n))}break;case"?":this.cleanedTargetName=e.substr(1),this.updateTarget=f;break;case"@":this.cleanedTargetName=e.substr(1),this.bind=s,this.unbind=l;break;default:this.cleanedTargetName=e,"class"===e&&(this.updateTarget=_)}}targetAtContent(){this.updateTarget=p,this.unbind=d}createBehavior(e){return new b(e,this.binding,this.isBindingVolatile,this.bind,this.unbind,this.updateTarget,this.cleanedTargetName)}}class b{constructor(e,t,n,a,i,r,o){this.source=null,this.context=null,this.bindingObserver=null,this.target=e,this.binding=t,this.isBindingVolatile=n,this.bind=a,this.unbind=i,this.updateTarget=r,this.targetName=o}handleChange(){this.updateTarget(this.bindingObserver.observe(this.source,this.context))}handleEvent(e){i.a.setEvent(e);const t=this.binding(this.source,this.context);i.a.setEvent(null),!0!==t&&e.preventDefault()}}let g=null;class v{addFactory(e){e.targetIndex=this.targetIndex,this.behaviorFactories.push(e)}captureContentBinding(e){e.targetAtContent(),this.addFactory(e)}reset(){this.behaviorFactories=[],this.targetIndex=-1}release(){g=this}static borrow(e){const t=g||new v;return t.directives=e,t.reset(),g=null,t}}function y(e){if(1===e.length)return e[0];let t;const n=e.length,a=e.map(e=>"string"==typeof e?()=>e:(t=e.targetName||t,e.binding)),i=new h((e,t)=>{let i="";for(let r=0;r<n;++r)i+=a[r](e,t);return i});return i.targetName=t,i}const S=a.b.length;function D(e,t){const n=t.split(a.c);if(1===n.length)return null;const i=[];for(let t=0,r=n.length;t<r;++t){const r=n[t],o=r.indexOf(a.b);let s;if(-1===o)s=r;else{const t=parseInt(r.substring(0,o));i.push(e.directives[t]),s=r.substring(o+S)}""!==s&&i.push(s)}return i}function I(e,t,n=!1){const a=t.attributes;for(let i=0,r=a.length;i<r;++i){const o=a[i],s=o.value,c=D(e,s);let d=null;null===c?n&&(d=new h(()=>s),d.targetName=o.name):d=y(c),null!==d&&(t.removeAttributeNode(o),i--,r--,e.addFactory(d))}}function x(e,t,n){const a=D(e,t.textContent);if(null!==a){let i=t;for(let r=0,o=a.length;r<o;++r){const o=a[r],s=0===r?t:i.parentNode.insertBefore(document.createTextNode(""),i.nextSibling);"string"==typeof o?s.textContent=o:(s.textContent=" ",e.captureContentBinding(o)),i=s,e.targetIndex++,s!==t&&n.nextNode()}e.targetIndex--}}var C=n("Ne8q");class O{constructor(e,t){this.behaviorCount=0,this.hasHostBehaviors=!1,this.fragment=null,this.targetOffset=0,this.viewBehaviorFactories=null,this.hostBehaviorFactories=null,this.html=e,this.directives=t}create(e){if(null===this.fragment){let e;const t=this.html;if("string"==typeof t){e=document.createElement("template"),e.innerHTML=a.a.createHTML(t);const n=e.content.firstElementChild;null!==n&&"TEMPLATE"===n.tagName&&(e=n)}else e=t;const n=function(e,t){const n=e.content;document.adoptNode(n);const i=v.borrow(t);I(i,e,!0);const r=i.behaviorFactories;i.reset();const o=a.a.createTemplateWalker(n);let s;for(;s=o.nextNode();)switch(i.targetIndex++,s.nodeType){case 1:I(i,s);break;case 3:x(i,s,o);break;case 8:a.a.isMarker(s)&&i.addFactory(t[a.a.extractDirectiveIndexFromMarker(s)])}let c=0;(a.a.isMarker(n.firstChild)||1===n.childNodes.length&&t.length)&&(n.insertBefore(document.createComment(""),n.firstChild),c=-1);const d=i.behaviorFactories;return i.release(),{fragment:n,viewBehaviorFactories:d,hostBehaviorFactories:r,targetOffset:c}}(e,this.directives);this.fragment=n.fragment,this.viewBehaviorFactories=n.viewBehaviorFactories,this.hostBehaviorFactories=n.hostBehaviorFactories,this.targetOffset=n.targetOffset,this.behaviorCount=this.viewBehaviorFactories.length+this.hostBehaviorFactories.length,this.hasHostBehaviors=this.hostBehaviorFactories.length>0}const t=this.fragment.cloneNode(!0),n=this.viewBehaviorFactories,i=new Array(this.behaviorCount),r=a.a.createTemplateWalker(t);let o=0,s=this.targetOffset,c=r.nextNode();for(let e=n.length;o<e;++o){const e=n[o],t=e.targetIndex;for(;null!==c;){if(s===t){i[o]=e.createBehavior(c);break}c=r.nextNode(),s++}}if(this.hasHostBehaviors){const t=this.hostBehaviorFactories;for(let n=0,a=t.length;n<a;++n,++o)i[o]=t[n].createBehavior(e)}return new C.a(t,i)}render(e,t,n){"string"==typeof t&&(t=document.getElementById(t)),void 0===n&&(n=t);const a=this.create(n);return a.bind(e,i.c),a.appendTo(t),a}}const w=/([ \x09\x0a\x0c\x0d])([^\0-\x1F\x7F-\x9F "'>=/]+)([ \x09\x0a\x0c\x0d]*=[ \x09\x0a\x0c\x0d]*(?:[^ \x09\x0a\x0c\x0d"'`<>=]*|"[^"]*|'[^']*))$/;function E(e,...t){const n=[];let a="";for(let i=0,o=e.length-1;i<o;++i){const o=e[i];let s=t[i];if(a+=o,s instanceof O){const e=s;s=()=>e}if("function"==typeof s&&(s=new h(s)),s instanceof r.c){const e=w.exec(o);null!==e&&(s.targetName=e[2])}s instanceof r.b?(a+=s.createPlaceholder(n.length),n.push(s)):a+=s}return a+=e[e.length-1],new O(a,n)}},"6eT7":function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("0q6d");class i{constructor(e,t,n,a){this.r=e,this.g=t,this.b=n,this.a="number"!=typeof a||isNaN(a)?1:a}static fromObject(e){return!e||isNaN(e.r)||isNaN(e.g)||isNaN(e.b)?null:new i(e.r,e.g,e.b,e.a)}equalValue(e){return this.r===e.r&&this.g===e.g&&this.b===e.b&&this.a===e.a}toStringHexRGB(){return"#"+[this.r,this.g,this.b].map(this.formatHexValue).join("")}toStringHexRGBA(){return this.toStringHexRGB()+this.formatHexValue(this.a)}toStringHexARGB(){return"#"+[this.a,this.r,this.g,this.b].map(this.formatHexValue).join("")}toStringWebRGB(){return`rgb(${Math.round(Object(a.c)(this.r,0,255))},${Math.round(Object(a.c)(this.g,0,255))},${Math.round(Object(a.c)(this.b,0,255))})`}toStringWebRGBA(){return`rgba(${Math.round(Object(a.c)(this.r,0,255))},${Math.round(Object(a.c)(this.g,0,255))},${Math.round(Object(a.c)(this.b,0,255))},${Object(a.a)(this.a,0,1)})`}roundToPrecision(e){return new i(Object(a.i)(this.r,e),Object(a.i)(this.g,e),Object(a.i)(this.b,e),Object(a.i)(this.a,e))}clamp(){return new i(Object(a.a)(this.r,0,1),Object(a.a)(this.g,0,1),Object(a.a)(this.b,0,1),Object(a.a)(this.a,0,1))}toObject(){return{r:this.r,g:this.g,b:this.b,a:this.a}}formatHexValue(e){return Object(a.d)(Object(a.c)(e,0,255))}}},"6fxl":function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("QBS5");function i(e,...t){const n=a.a.locate(e);t.forEach(t=>{Object.getOwnPropertyNames(t.prototype).forEach(n=>{"constructor"!==n&&Object.defineProperty(e.prototype,n,Object.getOwnPropertyDescriptor(t.prototype,n))}),a.a.locate(t).forEach(e=>n.push(e))})}},"6gCt":function(e,t,n){"use strict";function a(e,t,n){return e.nodeType!==Node.TEXT_NODE||"string"==typeof e.nodeValue&&!!e.nodeValue.trim().length}n.d(t,"a",function(){return a})},"6vBc":function(e,t,n){"use strict";n.d(t,"a",function(){return o});const a=e=>"function"==typeof e,i=()=>null;function r(e){return void 0===e?i:a(e)?e:()=>e}function o(e,t,n){const i=a(e)?e:()=>e,o=r(t),s=r(n);return(e,t)=>i(e,t)?o(e,t):s(e,t)}},"7Cdu":function(e,t,n){"use strict";n.d(t,"a",function(){return a}),n.d(t,"b",function(){return r});var a,i=n("qqMp");!function(e){e[e.ArrowDown=0]="ArrowDown",e[e.TeamSeparator=1]="TeamSeparator",e[e.Search=2]="Search",e[e.SkypeArrow=3]="SkypeArrow",e[e.SmallEmail=4]="SmallEmail",e[e.SmallEmailHovered=5]="SmallEmailHovered",e[e.SmallChat=6]="SmallChat",e[e.SmallChatHovered=7]="SmallChatHovered",e[e.Video=8]="Video",e[e.VideoHovered=9]="VideoHovered",e[e.ExpandDown=10]="ExpandDown",e[e.Overview=11]="Overview",e[e.Send=12]="Send",e[e.Contact=13]="Contact",e[e.Copy=14]="Copy",e[e.Phone=15]="Phone",e[e.CellPhone=16]="CellPhone",e[e.Chat=17]="Chat",e[e.Call=18]="Call",e[e.CallHovered=19]="CallHovered",e[e.Confirmation=20]="Confirmation",e[e.Department=21]="Department",e[e.Email=22]="Email",e[e.OfficeLocation=23]="OfficeLocation",e[e.Person=24]="Person",e[e.Messages=25]="Messages",e[e.Organization=26]="Organization",e[e.ExpandRight=27]="ExpandRight",e[e.Profile=28]="Profile",e[e.Birthday=29]="Birthday",e[e.File=30]="File",e[e.Files=31]="Files",e[e.Back=32]="Back",e[e.Close=33]="Close",e[e.Upload=34]="Upload",e[e.FileCloud=35]="FileCloud",e[e.DragFile=36]="DragFile",e[e.Cancel=37]="Cancel",e[e.CheckMark=38]="CheckMark",e[e.Success=39]="Success",e[e.Fail=40]="Fail",e[e.SelectAccount=41]="SelectAccount",e[e.News=42]="News",e[e.DoubleBookmark=43]="DoubleBookmark",e[e.ChevronLeft=44]="ChevronLeft",e[e.ChevronRight=45]="ChevronRight",e[e.Event=46]="Event",e[e.BookOpen=47]="BookOpen",e[e.FileOuter=48]="FileOuter",e[e.BookQuestion=49]="BookQuestion",e[e.Globe=50]="Globe",e[e.Delete=51]="Delete",e[e.Add=52]="Add",e[e.Calendar=53]="Calendar",e[e.Planner=54]="Planner",e[e.Milestone=55]="Milestone",e[e.PersonAdd=56]="PersonAdd",e[e.PresenceAvailable=57]="PresenceAvailable",e[e.PresenceOofAvailable=58]="PresenceOofAvailable",e[e.PresenceBusy=59]="PresenceBusy",e[e.PresenceOofBusy=60]="PresenceOofBusy",e[e.PresenceDnd=61]="PresenceDnd",e[e.PresenceOofDnd=62]="PresenceOofDnd",e[e.PresenceAway=63]="PresenceAway",e[e.PresenceOofAway=64]="PresenceOofAway",e[e.PresenceOffline=65]="PresenceOffline",e[e.PresenceStatusUnknown=66]="PresenceStatusUnknown",e[e.Loading=67]="Loading"}(a||(a={}));const r=(e,t)=>{switch(e){case a.ArrowDown:return i.c`
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M6 9L2.46447 5.46447H6H9.53553L6 9Z" />
        </svg>
      `;case a.TeamSeparator:return i.c`
        <svg width="6" height="10" viewBox="0 0 6 10" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path
            fill-rule="evenodd"
            clip-rule="evenodd"
            d="M5.70711 5L1.49999 9.20711L0.792886 8.50001L4.29289 5L0.792887 1.49999L1.49999 0.792885L5.70711 5Z"
            fill=${t}
          />
        </svg>
      `;case a.Search:return i.c`
        <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
          <path d="M7.30887 8.01596C6.53903 8.63179 5.56252 9.00004 4.5 9.00004C2.01472 9.00004 0 6.98531 0 4.50002C0 2.01473 2.01472 0 4.5 0C6.98528 0 9 2.01473 9 4.50002C9 5.56252 8.63177 6.53901 8.01597 7.30885L11.8536 11.1464C12.0488 11.3417 12.0488 11.6583 11.8536 11.8536C11.6583 12.0488 11.3417 12.0488 11.1464 11.8536L7.30887 8.01596ZM8 4.50002C8 2.56701 6.433 1 4.5 1C2.567 1 1 2.56701 1 4.50002C1 6.43302 2.567 8.00003 4.5 8.00003C6.433 8.00003 8 6.43302 8 4.50002Z" fill="currentColor"/>
        </svg>`;case a.SkypeArrow:return i.c`
        <svg viewBox="0 0 12 10" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path
            d="M3.95184 0.480534C4.23385 0.10452 4.70926 -0.0724722 5.1685 0.0275755C5.62775 0.127623 5.98645 0.486329 6.0865 0.945575C6.18655 1.40482 6.00956 1.88023 5.63354 2.16224L4.07196 3.72623H10.7988C11.4622 3.72623 12 4.26403 12 4.92744C12 5.59086 11.4622 6.12866 10.7988 6.12866H4.07196L5.63114 7.68784C6.0955 8.15225 6.0955 8.90515 5.63114 9.36955C5.51655 9.48381 5.38119 9.57514 5.23234 9.63862C5.09341 9.69857 4.94399 9.73042 4.79269 9.73232C4.63498 9.73233 4.4789 9.70046 4.33382 9.63862C4.18765 9.57669 4.05593 9.48507 3.94703 9.36955L0.343377 5.7659C-0.114459 5.29881 -0.114459 4.55128 0.343377 4.08419L3.95184 0.480534Z"
            fill="#B4009E"
          />
        </svg>
      `;case a.SmallEmail:return i.c`
        <svg width="17" height="14" viewBox="0 0 17 14" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path class="filled" d="M16 3.373V10.5C16 11.8807 14.8807 13 13.5 13H2.5C1.11929 13 -7.62939e-05 11.8807 -7.62939e-05 10.5V3.373L7.74649 7.93097C7.90297 8.02301 8.09704 8.02301 8.25351 7.93097L16 3.373ZM13.5 0C14.7871 0 15.847 0.972604 15.9848 2.22293L8 6.91991L0.0151832 2.22293C0.15304 0.972604 1.21294 0 2.5 0H13.5Z" fill="currentColor"/>
          <path class="regular" d="M13.608 0.833496C14.9887 0.833496 16.108 1.95278 16.108 3.3335V11.3335C16.108 12.7142 14.9887 13.8335 13.608 13.8335H2.60803C1.22732 13.8335 0.108032 12.7142 0.108032 11.3335V3.3335C0.108032 1.95278 1.22732 0.833496 2.60803 0.833496H13.608ZM15.108 4.7945L8.36154 8.76446C8.23115 8.84117 8.07464 8.85395 7.93554 8.80281L7.85452 8.76446L1.10803 4.7965V11.3335C1.10803 12.1619 1.77961 12.8335 2.60803 12.8335H13.608C14.4365 12.8335 15.108 12.1619 15.108 11.3335V4.7945ZM13.608 1.8335H2.60803C1.77961 1.8335 1.10803 2.50507 1.10803 3.3335V3.6355L8.10803 7.75341L15.108 3.6345V3.3335C15.108 2.50507 14.4365 1.8335 13.608 1.8335Z" fill="currentColor"/>
        </svg>
      `;case a.SmallChat:return i.c`
        <svg width="17" height="17" viewBox="0 0 17 17" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path class="filled" d="M8 0C12.4183 0 16 3.58172 16 8C16 12.4183 12.4183 16 8 16C6.72679 16 5.49591 15.7018 4.38669 15.1393L4.266 15.075L0.621091 15.9851C0.311269 16.0625 0.0262241 15.8369 0.00130987 15.5438L0.00114131 15.4624L0.0149329 15.3787L0.925 11.735L0.86169 11.6153C0.406602 10.7186 0.124331 9.74223 0.0327466 8.72826L0.00737596 8.34634L0 8C0 3.58172 3.58172 0 8 0ZM8.5 9H5.5L5.41012 9.00806C5.17688 9.05039 5 9.25454 5 9.5C5 9.74546 5.17688 9.94961 5.41012 9.99194L5.5 10H8.5L8.58988 9.99194C8.82312 9.94961 9 9.74546 9 9.5C9 9.25454 8.82312 9.05039 8.58988 9.00806L8.5 9ZM10.5 6H5.5L5.41012 6.00806C5.17688 6.05039 5 6.25454 5 6.5C5 6.74546 5.17688 6.94961 5.41012 6.99194L5.5 7H10.5L10.5899 6.99194C10.8231 6.94961 11 6.74546 11 6.5C11 6.25454 10.8231 6.05039 10.5899 6.00806L10.5 6Z" fill="currentColor"/>
          <path class="regular" d="M8.38599 0.833496C12.8043 0.833496 16.386 4.41522 16.386 8.8335C16.386 13.2518 12.8043 16.8335 8.38599 16.8335C7.11277 16.8335 5.8819 16.5353 4.77267 15.9728L4.65199 15.9085L1.00708 16.8186C0.697255 16.8959 0.41221 16.6704 0.387296 16.3773L0.387128 16.2959L0.400919 16.2122L1.31099 12.5685L1.24768 12.4488C0.792589 11.5521 0.510317 10.5757 0.418733 9.56176L0.393362 9.17984L0.385986 8.8335C0.385986 4.41522 3.96771 0.833496 8.38599 0.833496ZM8.38599 1.8335C4.51999 1.8335 1.38599 4.9675 1.38599 8.8335C1.38599 10.0505 1.69653 11.2213 2.27951 12.2584C2.32645 12.3419 2.34809 12.4365 2.34291 12.5308L2.32873 12.6247L1.57299 15.6455L4.59703 14.8918C4.65892 14.8764 4.72261 14.8731 4.78472 14.8814L4.87629 14.9026L4.963 14.941C5.9996 15.5233 7.16969 15.8335 8.38599 15.8335C12.252 15.8335 15.386 12.6995 15.386 8.8335C15.386 4.9675 12.252 1.8335 8.38599 1.8335ZM8.88599 9.8335C9.16213 9.8335 9.38599 10.0574 9.38599 10.3335C9.38599 10.579 9.20911 10.7831 8.97586 10.8254L8.88599 10.8335H5.88599C5.60984 10.8335 5.38599 10.6096 5.38599 10.3335C5.38599 10.088 5.56286 9.88389 5.79611 9.84155L5.88599 9.8335H8.88599ZM10.886 6.8335C11.1621 6.8335 11.386 7.05735 11.386 7.3335C11.386 7.57896 11.2091 7.7831 10.9759 7.82544L10.886 7.8335L5.88599 7.8335C5.60984 7.8335 5.38599 7.60964 5.38599 7.3335C5.38599 7.08804 5.56286 6.88389 5.79611 6.84155L5.88599 6.8335L10.886 6.8335Z" fill="currentColor"/>
        </svg>
      `;case a.Video:return i.c`
        <svg width="17" height="13" viewBox="0 0 17 13" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path class="filled" d="M11 2.5C11 1.11929 9.88071 0 8.5 0H2.5C1.11929 0 0 1.11929 0 2.5V9.5C0 10.8807 1.11929 12 2.5 12H8.5C9.88071 12 11 10.8807 11 9.5V2.5ZM12 3.93082V8.08149L14.7642 10.4319C15.2512 10.8461 16 10.4999 16 9.86055V2.19315C16 1.55685 15.2575 1.20962 14.7692 1.61756L12 3.93082Z" fill="currentColor"/>
          <path class="regular" d="M2.72095 0.833496C1.34024 0.833496 0.220947 1.95278 0.220947 3.3335V10.3335C0.220947 11.7142 1.34024 12.8335 2.72095 12.8335H9.72095C11.1017 12.8335 12.2209 11.7142 12.2209 10.3335V9.33348L14.6209 11.1335C15.2802 11.6279 16.2209 11.1575 16.2209 10.3335V3.33348C16.2209 2.50944 15.2802 2.03905 14.6209 2.53348L12.2209 4.33348V3.3335C12.2209 1.95278 11.1017 0.833496 9.72095 0.833496H2.72095ZM12.2209 5.58348L15.2209 3.33348V10.3335L12.2209 8.08348V5.58348ZM11.2209 3.3335V10.3335C11.2209 11.1619 10.5494 11.8335 9.72095 11.8335H2.72095C1.89252 11.8335 1.22095 11.1619 1.22095 10.3335V3.3335C1.22095 2.50507 1.89252 1.8335 2.72095 1.8335H9.72095C10.5494 1.8335 11.2209 2.50507 11.2209 3.3335Z" fill="currentColor"/>
        </svg>
      `;case a.ExpandDown:return i.c`
        <svg width="15" height="8" viewBox="0 0 15 8" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M14 1L7.5 7L1 1" stroke="currentColor" />
        </svg>
      `;case a.Overview:return i.c`
        <svg width="16" height="13" viewBox="0 0 16 13" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M3.5 10.5C2.67157 10.5 2 9.82843 2 9V4C2 3.17157 2.67157 2.5 3.5 2.5H12.5C13.3284 2.5 14 3.17157 14 4V9C14 9.82843 13.3284 10.5 12.5 10.5H3.5ZM3.5 3.5C3.22386 3.5 3 3.72386 3 4V5.5H13V4C13 3.72386 12.7761 3.5 12.5 3.5H3.5ZM5 6.5H3V9C3 9.27614 3.22386 9.5 3.5 9.5H5V6.5ZM6 9.5H12.5C12.7761 9.5 13 9.27614 13 9V6.5H6V9.5ZM0 3C0 1.61929 1.11929 0.5 2.5 0.5H13.5C14.8807 0.5 16 1.61929 16 3V10C16 11.3807 14.8807 12.5 13.5 12.5H2.5C1.11929 12.5 0 11.3807 0 10V3ZM2.5 1.5C1.67157 1.5 1 2.17157 1 3V10C1 10.8284 1.67157 11.5 2.5 11.5H13.5C14.3284 11.5 15 10.8284 15 10V3C15 2.17157 14.3284 1.5 13.5 1.5H2.5Z" fill="#717171"/>
        </svg>
      `;case a.Send:return i.c`
        <svg xmlns="http://www.w3.org/2000/svg">
          <path
            d="M4.27144 8.99999L1.72572 2.45387C1.54854 1.99826 1.9928 1.56256 2.43227 1.71743L2.50153 1.74688L16.0015 8.49688C16.3902 8.69122 16.4145 9.22336 16.0744 9.45992L16.0015 9.50311L2.50153 16.2531C2.0643 16.4717 1.58932 16.0697 1.70282 15.6178L1.72572 15.5461L4.27144 8.99999L1.72572 2.45387L4.27144 8.99999ZM3.3028 3.4053L5.25954 8.43705L10.2302 8.43749C10.515 8.43749 10.7503 8.64911 10.7876 8.92367L10.7927 8.99999C10.7927 9.28476 10.5811 9.52011 10.3065 9.55736L10.2302 9.56249L5.25954 9.56205L3.3028 14.5947L14.4922 8.99999L3.3028 3.4053Z"
          />
        </svg>
      `;case a.Contact:return i.c`
        <svg width="16" height="12" viewBox="0 0 16 12" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M6 4.5C6 5.32843 5.32843 6 4.5 6C3.67157 6 3 5.32843 3 4.5C3 3.67157 3.67157 3 4.5 3C5.32843 3 6 3.67157 6 4.5ZM2 7.69879C2 7.17479 2.42479 6.75 2.94879 6.75H6.05121C6.57521 6.75 7 7.17479 7 7.69879C7 8.54603 6.42338 9.28454 5.60144 9.49003L5.54243 9.50478C4.85801 9.67589 4.14199 9.67589 3.45757 9.50478L3.39856 9.49003C2.57661 9.28454 2 8.54603 2 7.69879ZM9.5 4C9.22386 4 9 4.22386 9 4.5C9 4.77614 9.22386 5 9.5 5H12.5C12.7761 5 13 4.77614 13 4.5C13 4.22386 12.7761 4 12.5 4H9.5ZM9.5 7C9.22386 7 9 7.22386 9 7.5C9 7.77614 9.22386 8 9.5 8H12.5C12.7761 8 13 7.77614 13 7.5C13 7.22386 12.7761 7 12.5 7H9.5ZM0 1.75C0 0.783502 0.783502 0 1.75 0H14.25C15.2165 0 16 0.783502 16 1.75V10.25C16 11.2165 15.2165 12 14.25 12H1.75C0.783501 12 0 11.2165 0 10.25V1.75ZM1.75 1C1.33579 1 1 1.33579 1 1.75V10.25C1 10.6642 1.33579 11 1.75 11H14.25C14.6642 11 15 10.6642 15 10.25V1.75C15 1.33579 14.6642 1 14.25 1H1.75Z" fill="#717171"/>
        </svg>
      `;case a.Call:return i.c`
        <svg width="13" height="17" viewBox="0 0 13 17" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path class="filled" d="M3.98706 1.06589C4.89545 0.792081 5.86254 1.19479 6.31418 2.01224L6.38841 2.16075L7.04987 3.63213C7.46246 4.54992 7.28209 5.61908 6.60754 6.3496L6.47529 6.48248L5.43194 7.45541C5.24417 7.63298 5.38512 8.32181 6.06527 9.49986C6.67716 10.5597 7.17487 11.0552 7.41986 11.0823L7.4628 11.082L7.5158 11.0716L9.56651 10.4446C10.1332 10.2713 10.7438 10.4487 11.1298 10.8865L11.2215 11.0014L12.5781 12.8815C13.1299 13.6462 13.0689 14.6842 12.4533 15.378L12.3314 15.5039L11.7886 16.018C10.7948 16.9592 9.34348 17.2346 8.07389 16.7231C6.13867 15.9433 4.38077 14.1607 2.78368 11.3945C1.18323 8.62242 0.519004 6.20438 0.815977 4.13565C0.99977 2.85539 1.87301 1.78674 3.07748 1.3462L3.27036 1.28192L3.98706 1.06589Z" fill="currentColor"/>
          <path class="regular" d="M3.9292 0.399388L3.2125 0.615419C1.90898 1.00834 0.951582 2.1215 0.758116 3.46915C0.461142 5.53787 1.12537 7.95591 2.72582 10.728C4.32291 13.4942 6.0808 15.2768 8.01603 16.0565C9.28562 16.5681 10.7369 16.2927 11.7308 15.3515L12.2736 14.8374C13.0011 14.1485 13.1065 13.0275 12.5202 12.215L11.1636 10.3349C10.788 9.81423 10.1226 9.59039 9.50865 9.7781L7.45794 10.4051L7.40494 10.4154C7.17877 10.4485 6.65754 9.95942 6.00741 8.83335C5.32726 7.65531 5.1863 6.96648 5.37408 6.7889L6.41743 5.81598C7.19937 5.08681 7.43039 3.94078 6.99201 2.96562L6.33055 1.49424C5.91866 0.578005 4.89102 0.109471 3.9292 0.399388ZM5.41847 1.90427L6.07993 3.37564C6.34277 3.96031 6.20426 4.64744 5.73543 5.08463L4.68953 6.05994C4.02044 6.69268 4.24206 7.77567 5.14138 9.33335C5.98759 10.799 6.75958 11.5233 7.58908 11.3977L7.71341 11.3711L9.80102 10.7344C10.0057 10.6718 10.2275 10.7464 10.3527 10.92L11.7093 12.8001C12.0024 13.2064 11.9497 13.7669 11.586 14.1113L11.0432 14.6254C10.3333 15.2977 9.29663 15.4944 8.38977 15.129C6.6917 14.4448 5.08689 12.8175 3.59185 10.228C2.09375 7.63319 1.48745 5.42604 1.74797 3.61125C1.88616 2.64864 2.57001 1.85352 3.5011 1.57287L4.2178 1.35684C4.69871 1.21188 5.21253 1.44615 5.41847 1.90427Z" fill="currentColor"/>
        </svg>
      `;case a.Confirmation:return i.c`
      <svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M7.03212 13.9072L3.56056 10.0017C3.28538 9.69214 2.81132 9.66425 2.50174 9.93944C2.19215 10.2146 2.16426 10.6887 2.43945 10.9983L6.43945 15.4983C6.72614 15.8208 7.2252 15.8355 7.53034 15.5303L18.0303 5.03033C18.3232 4.73744 18.3232 4.26256 18.0303 3.96967C17.7374 3.67678 17.2626 3.67678 16.9697 3.96967L7.03212 13.9072Z" fill="#009E00"/>
      </svg>
      `;case a.Copy:return i.c`
        <svg width="13" height="14" viewBox="0 0 13 14" xmlns="http://www.w3.org/2000/svg">
          <path
            d="M12.625 5.50293V14H3.875V11.375H0.375V0H6.24707L8.87207 2.625H9.74707L12.625 5.50293ZM10 5.25H11.1279L10 4.12207V5.25ZM3.875 2.625H7.62793L5.87793 0.875H1.25V10.5H3.875V2.625ZM11.75 6.125H9.125V3.5H4.75V13.125H11.75V6.125Z"
          />
        </svg>
      `;case a.Phone:return i.c`
        <svg width="15" height="15" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048" fill="none">
            <path d="M1607 1213q44 0 84 16t72 48l220 220q31 31 47 71t17 85q0 44-16 84t-48 72l-14 14q-54 54-99 96t-94 70-109 44-143 15q-125 0-257-39t-262-108-256-164-237-207-206-238-162-256T38 775 0 523q0-83 14-142t43-108 70-93 96-99l16-16q31-31 71-48t85-17q44 0 84 17t72 48l220 220q31 31 47 71t17 85q0 44-15 78t-37 63-48 51-49 45-37 44-15 49q0 38 27 65l551 551q27 27 65 27 26 0 48-15t45-37 45-48 51-49 62-37 79-15zm-83 707q72 0 120-13t88-39 76-64 85-86q27-27 27-65 0-18-14-42t-38-52-51-55-56-54-51-47-37-35q-27-27-66-27-26 0-48 15t-44 37-45 48-52 49-62 37-79 15q-44 0-84-16t-72-48L570 927q-31-31-47-71t-17-85q0-44 15-78t37-63 48-51 49-46 37-44 15-48q0-39-27-66-13-13-34-36t-47-51-54-56-56-52-51-37-43-15q-38 0-65 27l-85 85q-37 37-64 76t-40 87-14 120q0 112 36 231t101 238 153 234 192 219 219 190 234 150 236 99 226 36z" fill="${t}"></path>
        </svg>
      `;case a.CellPhone:return i.c`
        <svg width="10" height="16" viewBox="0 0 10 16" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M4 12C3.72386 12 3.5 12.2239 3.5 12.5C3.5 12.7761 3.72386 13 4 13H6C6.27614 13 6.5 12.7761 6.5 12.5C6.5 12.2239 6.27614 12 6 12H4ZM2 0C0.89543 0 0 0.895431 0 2V14C0 15.1046 0.895431 16 2 16H8C9.10457 16 10 15.1046 10 14V2C10 0.89543 9.10457 0 8 0H2ZM1 2C1 1.44772 1.44772 1 2 1H8C8.55228 1 9 1.44772 9 2V14C9 14.5523 8.55228 15 8 15H2C1.44772 15 1 14.5523 1 14V2Z" fill="#717171"/>
        </svg>
      `;case a.Chat:return i.c`
        <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M8 0C12.4183 0 16 3.58172 16 8C16 12.4183 12.4183 16 8 16C6.72679 16 5.49591 15.7018 4.38669 15.1393L4.266 15.075L0.621091 15.9851C0.311269 16.0625 0.0262241 15.8369 0.00130987 15.5438L0.00114131 15.4624L0.0149329 15.3787L0.925 11.735L0.86169 11.6153C0.406602 10.7186 0.124331 9.74223 0.0327466 8.72826L0.00737596 8.34634L0 8C0 3.58172 3.58172 0 8 0ZM8 1C4.13401 1 1 4.13401 1 8C1 9.21704 1.31054 10.3878 1.89352 11.4249C1.94046 11.5084 1.9621 11.603 1.95692 11.6973L1.94274 11.7912L1.187 14.812L4.21104 14.0583C4.27294 14.0429 4.33662 14.0396 4.39873 14.0479L4.4903 14.0691L4.57701 14.1075C5.61362 14.6898 6.7837 15 8 15C11.866 15 15 11.866 15 8C15 4.13401 11.866 1 8 1ZM8.5 9C8.77614 9 9 9.22386 9 9.5C9 9.74546 8.82312 9.94961 8.58988 9.99194L8.5 10H5.5C5.22386 10 5 9.77614 5 9.5C5 9.25454 5.17688 9.05039 5.41012 9.00806L5.5 9H8.5ZM10.5 6C10.7761 6 11 6.22386 11 6.5C11 6.74546 10.8231 6.94961 10.5899 6.99194L10.5 7L5.5 7C5.22386 7 5 6.77614 5 6.5C5 6.25454 5.17688 6.05039 5.41012 6.00806L5.5 6L10.5 6Z" fill="#717171"/>
        </svg>
      `;case a.Department:return i.c`
        <svg width="17" height="14" viewBox="0 0 17 14" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M4 1.5V3H2C0.895431 3 0 3.89543 0 5V12C0 13.1046 0.895432 14 2 14H14.0026C15.1071 14 16.0026 13.1046 16.0026 12V5C16.0026 3.89543 15.1071 3 14.0026 3H12V1.5C12 0.671573 11.3284 0 10.5 0H5.5C4.67157 0 4 0.671573 4 1.5ZM5.5 1H10.5C10.7761 1 11 1.22386 11 1.5V3H5V1.5C5 1.22386 5.22386 1 5.5 1ZM2 4H14.0026C14.5549 4 15.0026 4.44772 15.0026 5V7H12L12 6.50073C12 6.22459 11.7761 6.00073 11.5 6.00073C11.2239 6.00073 11 6.22459 11 6.50074L11 7H5V6.50073C5 6.22459 4.77614 6.00073 4.5 6.00073C4.22386 6.00073 4 6.22459 4 6.50073V7H1V5C1 4.44772 1.44772 4 2 4ZM11 8L11 8.5C11 8.77615 11.2239 9 11.5 9C11.7762 9 12 8.77614 12 8.5L12 8H15.0026V12C15.0026 12.5523 14.5549 13 14.0026 13H2C1.44772 13 1 12.5523 1 12V8H4V8.5C4 8.77614 4.22386 9 4.5 9C4.77614 9 5 8.77614 5 8.5V8H11Z" fill="#717171"/>
        </svg>
      `;case a.Email:return i.c`
        <svg width="16" height="13" viewBox="0 0 16 13" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M13.5 0C14.8807 0 16 1.11929 16 2.5V10.5C16 11.8807 14.8807 13 13.5 13H2.5C1.11929 13 0 11.8807 0 10.5V2.5C0 1.11929 1.11929 0 2.5 0H13.5ZM15 3.961L8.25351 7.93097C8.12311 8.00767 7.96661 8.02045 7.82751 7.96932L7.74649 7.93097L1 3.963V10.5C1 11.3284 1.67157 12 2.5 12H13.5C14.3284 12 15 11.3284 15 10.5V3.961ZM13.5 1H2.5C1.67157 1 1 1.67157 1 2.5V2.802L8 6.91991L15 2.801V2.5C15 1.67157 14.3284 1 13.5 1Z" fill="#717171"/>
        </svg>
      `;case a.OfficeLocation:return i.c`
        <svg width="14" height="16" viewBox="0 0 14 16" fill="none" xmlns="http://www.w3.org/2000/svg">
          <title>Location icon</title>
          <path d="M10 7C10 8.65685 8.65685 10 7 10C5.34315 10 4 8.65685 4 7C4 5.34315 5.34315 4 7 4C8.65685 4 10 5.34315 10 7ZM9 7C9 5.89543 8.10457 5 7 5C5.89543 5 5 5.89543 5 7C5 8.10457 5.89543 9 7 9C8.10457 9 9 8.10457 9 7ZM11.9497 11.955C14.6834 9.2201 14.6834 4.78601 11.9497 2.05115C9.21608 -0.683716 4.78392 -0.683716 2.05025 2.05115C-0.683418 4.78601 -0.683418 9.2201 2.05025 11.955L3.57128 13.4538L5.61408 15.4389L5.74691 15.5567C6.52168 16.1847 7.65623 16.1455 8.38611 15.4391L10.8223 13.0691L11.9497 11.955ZM2.75499 2.75619C5.09944 0.410715 8.90055 0.410715 11.245 2.75619C13.5294 5.04153 13.5879 8.71039 11.4207 11.0667L11.245 11.2499L9.92371 12.5539L7.69315 14.7225L7.60016 14.8021C7.24594 15.0699 6.7543 15.0698 6.40012 14.802L6.30713 14.7224L3.3263 11.817L2.75499 11.2499L2.57927 11.0667C0.412077 8.71039 0.47065 5.04153 2.75499 2.75619Z" fill="#717171"/>
        </svg>
      `;case a.Birthday:return i.c`
        <svg width="14" height="13" viewBox="0 0 14 13" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M5.81357 0.667333C6.07581 0.320242 6.48151 0 7 0C7.51849 0 7.92419 0.320242 8.18643 0.667333C8.45471 1.0224 8.63508 1.47991 8.69561 1.93392C8.75552 2.3832 8.70532 2.89341 8.42852 3.3086C8.134 3.75039 7.63965 4 7 4C6.36035 4 5.866 3.75039 5.57148 3.3086C5.29468 2.89341 5.24448 2.3832 5.30439 1.93392C5.36492 1.47991 5.54529 1.0224 5.81357 0.667333ZM6.40353 2.7539C6.484 2.87461 6.63965 3 7 3C7.36035 3 7.516 2.87461 7.59647 2.7539C7.69468 2.60659 7.74448 2.3668 7.70439 2.06608C7.66492 1.77009 7.54529 1.4776 7.38857 1.27017C7.22581 1.05476 7.08151 1 7 1C6.91849 1 6.77419 1.05476 6.61143 1.27017C6.45471 1.4776 6.33508 1.77009 6.29561 2.06608C6.25552 2.3668 6.30532 2.60659 6.40353 2.7539ZM1 7C1 5.89543 1.89543 5 3 5H11C12.1046 5 13 5.89543 13 7V12H13.5C13.7761 12 14 12.2239 14 12.5C14 12.7761 13.7761 13 13.5 13H0.5C0.223858 13 0 12.7761 0 12.5C0 12.2239 0.223858 12 0.5 12H1V7ZM2 12H12V9.01931C11.9109 9.10285 11.8174 9.18842 11.7208 9.27412C11.4006 9.55804 11.0346 9.85319 10.6715 10.0802C10.3274 10.2953 9.90815 10.5 9.5 10.5C8.77182 10.5 8.30806 9.98986 8.00471 9.65617C7.9834 9.63273 7.96289 9.61016 7.94312 9.58869C7.5876 9.20261 7.35769 9 7 9C6.64231 9 6.4124 9.20261 6.05688 9.58869C6.03711 9.61016 6.0166 9.63273 5.99529 9.65617C5.69194 9.98986 5.22818 10.5 4.5 10.5C4.10587 10.5 3.69263 10.2897 3.35907 10.0789C3.00198 9.85313 2.63516 9.55951 2.31117 9.27666C2.20329 9.18247 2.09896 9.08844 2 8.9971V12ZM2 7.59993C2.05039 7.65198 2.11363 7.71652 2.1873 7.7902C2.38843 7.99132 2.6649 8.25801 2.96883 8.52334C3.27484 8.79049 3.59802 9.04687 3.89343 9.23362C4.21237 9.43525 4.41413 9.5 4.5 9.5C4.75817 9.5 4.93171 9.33433 5.32125 8.91131L5.32447 8.90781C5.65956 8.5439 6.16039 8 7 8C7.83961 8 8.34044 8.5439 8.67553 8.90781L8.67875 8.91131C9.06828 9.33433 9.24183 9.5 9.5 9.5C9.6106 9.5 9.82569 9.42967 10.1414 9.23229C10.4381 9.04681 10.7573 8.79196 11.0573 8.52588C11.3554 8.26163 11.6238 7.99594 11.8184 7.7955C11.89 7.72165 11.9513 7.65703 12 7.60506V7C12 6.44772 11.5523 6 11 6H3C2.44772 6 2 6.44772 2 7V7.59993Z" fill="#717171"/>
        </svg>
      `;case a.Person:return i.c`
        <svg width="14" height="16" viewBox="0 0 14 16" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M7 0C4.79086 0 3 1.79086 3 4C3 6.20914 4.79086 8 7 8C9.20914 8 11 6.20914 11 4C11 1.79086 9.20914 0 7 0ZM4 4C4 2.34315 5.34315 1 7 1C8.65685 1 10 2.34315 10 4C10 5.65685 8.65685 7 7 7C5.34315 7 4 5.65685 4 4ZM2.00873 9C0.903151 9 0 9.88687 0 11C0 12.6912 0.83281 13.9663 2.13499 14.7966C3.41697 15.614 5.14526 16 7 16C8.85474 16 10.583 15.614 11.865 14.7966C13.1672 13.9663 14 12.6912 14 11C14 9.89557 13.1045 9.00001 12 9.00001L2.00873 9ZM1 11C1 10.4467 1.44786 10 2.00873 10L12 10C12.5522 10 13 10.4478 13 11C13 12.3088 12.3777 13.2837 11.3274 13.9534C10.2568 14.636 8.73511 15 7 15C5.26489 15 3.74318 14.636 2.67262 13.9534C1.62226 13.2837 1 12.3088 1 11Z" fill="#717171"/>
        </svg>
      `;case a.Messages:return i.c`
        <svg width="16" height="13" viewBox="0 0 16 13" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M13.5 0C14.8807 0 16 1.11929 16 2.5V10.5C16 11.8807 14.8807 13 13.5 13H2.5C1.11929 13 0 11.8807 0 10.5V2.5C0 1.11929 1.11929 0 2.5 0H13.5ZM15 3.961L8.25351 7.93097C8.12311 8.00767 7.96661 8.02045 7.82751 7.96932L7.74649 7.93097L1 3.963V10.5C1 11.3284 1.67157 12 2.5 12H13.5C14.3284 12 15 11.3284 15 10.5V3.961ZM13.5 1H2.5C1.67157 1 1 1.67157 1 2.5V2.802L8 6.91991L15 2.801V2.5C15 1.67157 14.3284 1 13.5 1Z" fill="#717171"/>
        </svg>
      `;case a.Organization:return i.c`
        <svg width="16" height="17" viewBox="0 0 16 17" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M7.99992 0C6.34271 0 4.99927 1.34344 4.99927 3.00066C4.99927 4.48754 6.08073 5.72183 7.49999 5.95985V7.5H4.5C3.67157 7.5 3 8.17157 3 9V10.0416C1.5811 10.2799 0.5 11.514 0.5 13.0007C0.5 14.6579 1.84344 16.0013 3.50066 16.0013C5.15787 16.0013 6.50131 14.6579 6.50131 13.0007C6.50131 11.5136 5.41954 10.2791 4 10.0414V9C4 8.72386 4.22386 8.5 4.5 8.5H11.5C11.7761 8.5 12 8.72386 12 9V10.0416C10.5811 10.2799 9.5 11.514 9.5 13.0007C9.5 14.6579 10.8434 16.0013 12.5007 16.0013C14.1579 16.0013 15.5013 14.6579 15.5013 13.0007C15.5013 11.5136 14.4195 10.2791 13 10.0414V9C13 8.17157 12.3284 7.5 11.5 7.5H8.49999V5.95983C9.91918 5.72176 11.0006 4.48749 11.0006 3.00066C11.0006 1.34344 9.65714 0 7.99992 0ZM5.99927 3.00066C5.99927 1.89572 6.89499 1 7.99992 1C9.10485 1 10.0006 1.89572 10.0006 3.00066C10.0006 4.10559 9.10485 5.00131 7.99992 5.00131C6.89499 5.00131 5.99927 4.10559 5.99927 3.00066ZM1.5 13.0007C1.5 11.8957 2.39572 11 3.50066 11C4.60559 11 5.50131 11.8957 5.50131 13.0007C5.50131 14.1056 4.60559 15.0013 3.50066 15.0013C2.39572 15.0013 1.5 14.1056 1.5 13.0007ZM12.5007 11C13.6056 11 14.5013 11.8957 14.5013 13.0007C14.5013 14.1056 13.6056 15.0013 12.5007 15.0013C11.3957 15.0013 10.5 14.1056 10.5 13.0007C10.5 11.8957 11.3957 11 12.5007 11Z" fill="#717171"/>
        </svg>
      `;case a.ExpandRight:return i.c`
        <svg width="8" height="13" viewBox="0 0 8 13" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M1 12L6.5 6.5L1 1" stroke="#B8B8B8" stroke-width="2" />
        </svg>
      `;case a.Profile:return i.c`
        <svg width="14" height="16" viewBox="0 0 14 16" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M7 0C4.79086 0 3 1.79086 3 4C3 6.20914 4.79086 8 7 8C9.20914 8 11 6.20914 11 4C11 1.79086 9.20914 0 7 0ZM4 4C4 2.34315 5.34315 1 7 1C8.65685 1 10 2.34315 10 4C10 5.65685 8.65685 7 7 7C5.34315 7 4 5.65685 4 4ZM2.00873 9C0.903151 9 0 9.88687 0 11C0 12.6912 0.83281 13.9663 2.13499 14.7966C3.41697 15.614 5.14526 16 7 16C8.85474 16 10.583 15.614 11.865 14.7966C13.1672 13.9663 14 12.6912 14 11C14 9.89557 13.1045 9.00001 12 9.00001L2.00873 9ZM1 11C1 10.4467 1.44786 10 2.00873 10L12 10C12.5522 10 13 10.4478 13 11C13 12.3088 12.3777 13.2837 11.3274 13.9534C10.2568 14.636 8.73511 15 7 15C5.26489 15 3.74318 14.636 2.67262 13.9534C1.62226 13.2837 1 12.3088 1 11Z" fill="#717171"/>
        </svg>
      `;case a.File:return i.c`
        <svg width="28" height="28" viewBox="0 0 20 28" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path
            opacity="0.64"
            d="M19.613 6.993L13.018 0.421C12.7473 0.15221 12.3815 0.000947397 12 0H1.4C0.628 0 0 0.615 0 1.372V24.642C0 25.391 0.622 26 1.386 26H18.6C19.372 26 20 25.385 20 24.628V7.923C20 7.577 19.86 7.237 19.613 6.993Z"
            fill="#605E5C"
          />
          <path
            d="M19 24.628C19 24.83 18.816 25 18.6 25H1.386C1.173 25 1 24.84 1 24.642V1.372C1 1.17 1.184 1 1.4 1H12V6.6C12 7.372 12.628 8 13.4 8H19V24.628Z"
            fill="white"
          />
          <path d="M18.204 6.99994L13 1.81494V6.59994C13 6.81994 13.18 6.99994 13.4 6.99994H18.204Z" fill="white" />
        </svg>
      `;case a.Files:return i.c`
        <svg width="17" height="15" viewBox="0 0 17 15" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M14.9956 4.07309V4C14.9956 2.61929 13.8763 1.5 12.4956 1.5H7.66418L6.06391 0.299946C5.80428 0.105247 5.48851 0 5.16399 0H2.5C1.11916 0 -0.000181445 1.11949 2.20615e-08 2.50033L0.0011832 11.4978C0.00135142 12.7772 0.962583 13.8321 2.2023 13.9798C2.2867 13.9945 2.37375 14.0021 2.46289 14.0021H13.1798C13.8981 14.0021 14.5156 13.4929 14.6524 12.7877L16.0097 5.78769C16.1587 5.01967 15.696 4.29703 14.9956 4.07309ZM2.5 1H5.16399C5.27216 1 5.37742 1.03508 5.46396 1.09998L7.19756 2.40002C7.2841 2.46492 7.38936 2.5 7.49753 2.5H12.4956C13.324 2.5 13.9956 3.17157 13.9956 4V4.00214H3.824C3.10596 4.00214 2.48863 4.511 2.35158 5.21583L1.05351 11.8916C1.01941 11.7661 1.0012 11.634 1.00118 11.4976L1 2.5002C0.999891 1.67169 1.6715 1 2.5 1ZM3.33319 5.4067C3.37888 5.17176 3.58465 5.00214 3.824 5.00214H14.5372C14.8515 5.00214 15.0879 5.28874 15.028 5.59732L13.6706 12.5973C13.6251 12.8324 13.4192 13.0021 13.1798 13.0021H2.46289C2.14845 13.0021 1.91206 12.7154 1.97208 12.4067L3.33319 5.4067Z" fill="#717171"/>
        </svg>
      `;case a.Back:return i.c`
        <svg width="16" height="16" viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg">
          <path
            d="M16 8.5H1.95312L8.10156 14.6484L7.39844 15.3516L0.046875 8L7.39844 0.648438L8.10156 1.35156L1.95312 7.5H16V8.5Z"
          />
        </svg>
      `;case a.Close:return i.c`
        <svg width="10" height="10" viewBox="0 0 10 10" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
          <path d="M5.73838 5.032L9.70337 1.067L8.99638 0.360001L5.03137 4.325L1.06637 0.360001L0.359375 1.067L4.32438 5.032L0.359375 8.997L1.06637 9.704L5.03137 5.739L8.99638 9.704L9.70337 8.997L5.73838 5.032Z" fill="currentColor"/>
        </svg>
     `;case a.Upload:return i.c`
        <svg class="upload-icon" width="21" height="12" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M4.5 2C4.08579 2 3.75 2.33579 3.75 2.75C3.75 3.16421 4.08579 3.5 4.5 3.5H15C15.4142 3.5 15.75 3.16421 15.75 2.75C15.75 2.33579 15.4142 2 15 2H4.5ZM10.4963 17.3493C10.4466 17.7154 10.1328 17.9976 9.75311 17.9976C9.3389 17.9976 9.00311 17.6618 9.00311 17.2476L9.00249 7.05856L6.02995 10.026L5.94578 10.0986C5.65202 10.3162 5.23537 10.2917 4.96929 10.0253C4.67661 9.73215 4.67695 9.25728 4.97005 8.96459L9.25962 4.67989C9.33377 4.61512 9.42089 4.56485 9.5169 4.53385L9.59777 4.51072C9.64749 4.50019 9.69837 4.4947 9.74849 4.4947L9.80855 4.49661L9.87781 4.50451L9.99828 4.53462L10.0895 4.57254L10.1259 4.59371L10.2169 4.6523L10.2875 4.71481L14.5303 8.96546L14.6029 9.04964C14.8205 9.34345 14.7959 9.7601 14.5294 10.0261L14.4452 10.0987C14.1514 10.3162 13.7347 10.2917 13.4687 10.0251L10.5025 7.05456L10.5031 17.2476L10.4963 17.3493Z" fill="#ffffff"/>
        </svg>
      `;case a.FileCloud:return i.c`
        <svg width="16" height="16"  xmlns="http://www.w3.org/2000/svg">
          <path
            d="m8,0c2.8166,0 4.4145,1.9233 4.6469,4.246l0.071,0c1.8127,0 3.2821,1.5119 3.2821,3.377c0,0.0953 -0.0038,0.1897 -0.0114,0.283c-0.322,-0.4017 -0.6967,-0.7591 -1.1138,-1.062c-0.3104,-0.9329 -1.1627,-1.598 -2.1568,-1.598l-0.0711,0c-0.5137,0 -0.9439,-0.3893 -0.9951,-0.9005c-0.2021,-2.0206 -1.5433,-3.3455 -3.6518,-3.3455c-2.1139,0 -3.4489,1.3159 -3.6518,3.3455c-0.0511,0.5112 -0.4813,0.9005 -0.9951,0.9005l-0.071,0c-1.2539,0 -2.2821,1.0579 -2.2821,2.377c0,1.3191 1.0282,2.377 2.2821,2.377l2.6655,0c-0.087,0.323 -0.1466,0.6572 -0.1762,1l-2.4893,0c-1.8127,0 -3.2821,-1.5119 -3.2821,-3.377c0,-1.8029 1.3731,-3.2758 3.102,-3.372l0.2511,-0.005c0.2338,-2.338 1.8303,-4.246 4.6469,-4.246zm3.5,16c2.4853,0 4.5,-2.0147 4.5,-4.5c0,-2.4853 -2.0147,-4.5 -4.5,-4.5c-2.4853,0 -4.5,2.0147 -4.5,4.5c0,2.4853 2.0147,4.5 4.5,4.5zm0,-7c0.2761,0 0.5,0.2239 0.5,0.5l0,1.5l1.5,0c0.2761,0 0.5,0.2239 0.5,0.5c0,0.2761 -0.2239,0.5 -0.5,0.5l-1.5,0l0,1.5c0,0.2761 -0.2239,0.5 -0.5,0.5c-0.2761,0 -0.5,-0.2239 -0.5,-0.5l0,-1.5l-1.5,0c-0.2761,0 -0.5,-0.2239 -0.5,-0.5c0,-0.2761 0.2239,-0.5 0.5,-0.5l1.5,0l0,-1.5c0,-0.2761 0.2239,-0.5 0.5,-0.5z" fill="#0078D4"
          />
        </svg>
      `;case a.DragFile:return i.c`
        <svg width="13" height="16" xmlns="http://www.w3.org/2000/svg">
          <path
            d="m0,1.00189c0,-0.8451 0.983,-1.3091 1.636,-0.772l11.006,9.0622c0.724,0.5964 0.302,1.772 -0.636,1.772l-5.592,0c-0.435,0 -0.849,0.1892 -1.134,0.5185l-3.524,4.0725c-0.606,0.7003 -1.756,0.2717 -1.756,-0.6544l0,-13.9988zm12.006,9.0622l-11.006,-9.0622l0,13.9988l3.524,-4.0724c0.475,-0.5488 1.164,-0.8642 1.89,-0.8642l5.592,0z"
          />
        </svg>
      `;case a.Cancel:return i.c`
        <svg width="11" height="11" viewBox="0 0 11 11" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M0.588591 0.715694L0.646447 0.646447C0.820013 0.47288 1.08944 0.453595 1.28431 0.588591L1.35355 0.646447L6 5.293L10.6464 0.646447C10.8417 0.451185 11.1583 0.451185 11.3536 0.646447C11.5488 0.841709 11.5488 1.15829 11.3536 1.35355L6.707 6L11.3536 10.6464C11.5271 10.82 11.5464 11.0894 11.4114 11.2843L11.3536 11.3536C11.18 11.5271 10.9106 11.5464 10.7157 11.4114L10.6464 11.3536L6 6.707L1.35355 11.3536C1.15829 11.5488 0.841709 11.5488 0.646447 11.3536C0.451185 11.1583 0.451185 10.8417 0.646447 10.6464L5.293 6L0.646447 1.35355C0.47288 1.17999 0.453595 0.910563 0.588591 0.715694L0.646447 0.646447L0.588591 0.715694Z" fill="currentColor"/>
        </svg>
     `;case a.Success:return i.c`
        <svg width="12" height="12" xmlns="http://www.w3.org/2000/svg">
          <path fill="#009E00" d="m6.322,12c3.492,0 6.323,-2.6863 6.323,-6c0,-3.3137 -2.831,-6 -6.323,-6c-3.491,0 -6.322,2.6863 -6.322,6c0,3.3137 2.831,6 6.322,6z"/>
          <path fill="white" d="m9.629,3.7509c-0.131,-0.125 -0.31,-0.1952 -0.496,-0.1952c-0.187,0 -0.365,0.0702 -0.497,0.1952l-3.267,3.1l-1.393,-1.3222c-0.177,-0.1695 -0.436,-0.2361 -0.68,-0.1746c-0.243,0.0615 -0.433,0.2418 -0.497,0.4725c-0.065,0.2307 0.005,0.4767 0.184,0.6449l1.807,1.7154c0.019,0.0331 0.041,0.0646 0.066,0.094c0.289,0.2562 0.738,0.2562 1.027,0c0.024,-0.0294 0.047,-0.0609 0.065,-0.0941l3.681,-3.4931c0.275,-0.2603 0.275,-0.6824 0,-0.9428z"/>
        </svg>
      `;case a.CheckMark:return i.c`
      <svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path class="filled" d="M10 0C15.5228 0 20 4.47715 20 10C20 15.5228 15.5228 20 10 20C4.47715 20 0 15.5228 0 10C0 4.47715 4.47715 0 10 0ZM13.2197 6.96967L8.75 11.4393L6.78033 9.46967C6.48744 9.17678 6.01256 9.17678 5.71967 9.46967C5.42678 9.76256 5.42678 10.2374 5.71967 10.5303L8.21967 13.0303C8.51256 13.3232 8.98744 13.3232 9.28033 13.0303L14.2803 8.03033C14.5732 7.73744 14.5732 7.26256 14.2803 6.96967C13.9874 6.67678 13.5126 6.67678 13.2197 6.96967Z" fill="currentColor"/>
        <path class="regular" d="M10 1.5C5.30558 1.5 1.5 5.30558 1.5 10C1.5 14.6944 5.30558 18.5 10 18.5C14.6944 18.5 18.5 14.6944 18.5 10C18.5 5.30558 14.6944 1.5 10 1.5ZM0 10C0 4.47715 4.47715 0 10 0C15.5228 0 20 4.47715 20 10C20 15.5228 15.5228 20 10 20C4.47715 20 0 15.5228 0 10Z" fill="currentColor"/>
      </svg>
      `;case a.Fail:return i.c`
        <svg width="12" height="12" xmlns="http://www.w3.org/2000/svg">
          <path fill="#EF355D" d="m6,12c3.314,0 6,-2.6863 6,-6c0,-3.3137 -2.686,-6 -6,-6c-3.314,0 -6,2.6863 -6,6c0,3.3137 2.686,6 6,6z"/>
          <path fill="white" d="m6.943,6.0004l1.544,-1.5444c0.169,-0.1683 0.236,-0.4142 0.174,-0.6448c-0.061,-0.2306 -0.241,-0.4107 -0.472,-0.4722c-0.231,-0.0616 -0.477,0.0049 -0.645,0.1742l-1.544,1.5443l-1.545,-1.5443c-0.26,-0.259 -0.681,-0.2583 -0.941,0.0014c-0.26,0.2598 -0.26,0.6808 -0.001,0.9414l1.544,1.5444l-1.544,1.5444c-0.259,0.2606 -0.259,0.6815 0.001,0.9413c0.26,0.2598 0.681,0.2604 0.941,0.0015l1.545,-1.5444l1.544,1.5444c0.261,0.2589 0.682,0.2583 0.942,-0.0015c0.259,-0.2598 0.26,-0.6807 0.001,-0.9413l-1.544,-1.5444z" />
        </svg>
      `;case a.SelectAccount:return i.c`
    <svg xmlns="http://www.w3.org/2000/svg" width="17" height="17" viewBox="0 0 17 17" fill="none">
      <path fill=${t} d="M6.22176 13.9567C3.55468 13.653 2 11.8026 2 10V9.5C2 8.67157 2.67157 8 3.5 8H5.59971C5.43777 8.31679 5.30564 8.65136 5.20703 9H3.5C3.22386 9 3 9.22386 3 9.5V10C3 11.1281 3.88187 12.333 5.50235 12.7996C5.69426 13.216 5.93668 13.6043 6.22176 13.9567ZM9.62596 5.06907C9.70657 4.81036 9.75 4.53525 9.75 4.25C9.75 2.73122 8.51878 1.5 7 1.5C5.48122 1.5 4.25 2.73122 4.25 4.25C4.25 5.53662 5.13357 6.61687 6.32704 6.91706C6.64202 6.55055 7.00446 6.226 7.40482 5.95294C7.27488 5.98371 7.13934 6 7 6C6.0335 6 5.25 5.2165 5.25 4.25C5.25 3.2835 6.0335 2.5 7 2.5C7.9665 2.5 8.75 3.2835 8.75 4.25C8.75 4.73141 8.55561 5.16743 8.24104 5.48382C8.67558 5.28783 9.14016 5.14664 9.62596 5.06907ZM10.5 15C12.9853 15 15 12.9853 15 10.5C15 8.01472 12.9853 6 10.5 6C8.01472 6 6 8.01472 6 10.5C6 12.9853 8.01472 15 10.5 15ZM10.5 8C10.7761 8 11 8.22386 11 8.5V10H12.5C12.7761 10 13 10.2239 13 10.5C13 10.7761 12.7761 11 12.5 11H11V12.5C11 12.7761 10.7761 13 10.5 13C10.2239 13 10 12.7761 10 12.5V11H8.5C8.22386 11 8 10.7761 8 10.5C8 10.2239 8.22386 10 8.5 10H10V8.5C10 8.22386 10.2239 8 10.5 8Z"/>
    </svg>
  `;case a.News:return i.c`
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
        <path d="M18.75 20H5.25C3.51697 20 2.10075 18.6435 2.00514 16.9344L2 16.75V6.25C2 5.05914 2.92516 4.08436 4.09595 4.00519L4.25 4H16.75C17.9409 4 18.9156 4.92516 18.9948 6.09595L19 6.25V7H19.75C20.9409 7 21.9156 7.92516 21.9948 9.09595L22 9.25V16.75C22 18.483 20.6435 19.8992 18.9344 19.9949L18.75 20H5.25H18.75ZM5.25 18.5H18.75C19.6682 18.5 20.4212 17.7929 20.4942 16.8935L20.5 16.75V9.25C20.5 8.8703 20.2178 8.55651 19.8518 8.50685L19.75 8.5H19V16.25C19 16.6297 18.7178 16.9435 18.3518 16.9932L18.25 17C17.8703 17 17.5565 16.7178 17.5068 16.3518L17.5 16.25V6.25C17.5 5.8703 17.2178 5.55651 16.8518 5.50685L16.75 5.5H4.25C3.8703 5.5 3.55651 5.78215 3.50685 6.14823L3.5 6.25V16.75C3.5 17.6682 4.20711 18.4212 5.10647 18.4942L5.25 18.5H18.75H5.25ZM12.246 14.5H15.2522C15.6665 14.5 16.0022 14.8358 16.0022 15.25C16.0022 15.6297 15.7201 15.9435 15.354 15.9932L15.2522 16H12.246C11.8318 16 11.496 15.6642 11.496 15.25C11.496 14.8703 11.7782 14.5565 12.1442 14.5068L12.246 14.5H15.2522H12.246ZM9.24328 11.0045C9.6575 11.0045 9.99328 11.3403 9.99328 11.7545V15.25C9.99328 15.6642 9.6575 16 9.24328 16H5.74776C5.33355 16 4.99776 15.6642 4.99776 15.25V11.7545C4.99776 11.3403 5.33355 11.0045 5.74776 11.0045H9.24328ZM8.49328 12.5045H6.49776V14.5H8.49328V12.5045ZM12.246 11.0045H15.2522C15.6665 11.0045 16.0022 11.3403 16.0022 11.7545C16.0022 12.1342 15.7201 12.448 15.354 12.4976L15.2522 12.5045H12.246C11.8318 12.5045 11.496 12.1687 11.496 11.7545C11.496 11.3748 11.7782 11.061 12.1442 11.0113L12.246 11.0045H15.2522H12.246ZM5.74776 7.50247H15.2522C15.6665 7.50247 16.0022 7.83826 16.0022 8.25247C16.0022 8.63217 15.7201 8.94596 15.354 8.99563L15.2522 9.00247H5.74776C5.33355 9.00247 4.99776 8.66669 4.99776 8.25247C4.99776 7.87278 5.27991 7.55898 5.64599 7.50932L5.74776 7.50247H15.2522H5.74776Z" fill="none" />
      </svg>
      `;case a.DoubleBookmark:return i.c`
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
          <path d="M4 6.74814C4 5.50535 5.00742 4.49786 6.25013 4.49786H15.2506C16.4933 4.49786 17.5008 5.50535 17.5008 6.74814V21.2499C17.5008 21.5185 17.3572 21.7666 17.1243 21.9003C16.8914 22.0341 16.6048 22.0332 16.3728 21.8978L10.7504 18.6179L5.12797 21.8978C4.89599 22.0332 4.60936 22.0341 4.37648 21.9003C4.14359 21.7666 4 21.5185 4 21.2499V6.74814ZM6.25013 5.99805C5.83589 5.99805 5.50008 6.33387 5.50008 6.74814V19.944L10.3725 17.1016C10.606 16.9653 10.8948 16.9653 11.1283 17.1016L16.0007 19.944V6.74814C16.0007 6.33387 15.6649 5.99805 15.2506 5.99805H6.25013ZM15.2497 2C17.8732 2 20 4.12691 20 6.75058V18.6232C20 19.0374 19.6642 19.3733 19.25 19.3733C18.8357 19.3733 18.4999 19.0374 18.4999 18.6232V6.75058C18.4999 4.95543 17.0448 3.50018 15.2497 3.50018H6.63687C6.63687 3.50018 6.75016 2.94339 7.43379 2.41948C8.00023 2 8.60182 2 8.60182 2H15.2497Z" fill="none" />
        </svg>
      `;case a.ChevronLeft:return i.c`
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
          <path d="M15.5303 4.21967C15.8232 4.51256 15.8232 4.98744 15.5303 5.28033L8.81066 12L15.5303 18.7197C15.8232 19.0126 15.8232 19.4874 15.5303 19.7803C15.2374 20.0732 14.7626 20.0732 14.4697 19.7803L7.21967 12.5303C6.92678 12.2374 6.92678 11.7626 7.21967 11.4697L14.4697 4.21967C14.7626 3.92678 15.2374 3.92678 15.5303 4.21967Z" fill="none" />
        </svg>`;case a.ChevronRight:return i.c`
        <svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
          <path d="M8.46967 4.21967C8.17678 4.51256 8.17678 4.98744 8.46967 5.28033L15.1893 12L8.46967 18.7197C8.17678 19.0126 8.17678 19.4874 8.46967 19.7803C8.76256 20.0732 9.23744 20.0732 9.53033 19.7803L16.7803 12.5303C17.0732 12.2374 17.0732 11.7626 16.7803 11.4697L9.53033 4.21967C9.23744 3.92678 8.76256 3.92678 8.46967 4.21967Z" fill="currentColor" />
        </svg>`;case a.Delete:return i.c`
        <svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M8.5 4H11.5C11.5 3.17157 10.8284 2.5 10 2.5C9.17157 2.5 8.5 3.17157 8.5 4ZM7.5 4C7.5 2.61929 8.61929 1.5 10 1.5C11.3807 1.5 12.5 2.61929 12.5 4H17.5C17.7761 4 18 4.22386 18 4.5C18 4.77614 17.7761 5 17.5 5H16.4456L15.2521 15.3439C15.0774 16.8576 13.7957 18 12.2719 18H7.72813C6.20431 18 4.92256 16.8576 4.7479 15.3439L3.55437 5H2.5C2.22386 5 2 4.77614 2 4.5C2 4.22386 2.22386 4 2.5 4H7.5ZM5.74131 15.2292C5.85775 16.2384 6.71225 17 7.72813 17H12.2719C13.2878 17 14.1422 16.2384 14.2587 15.2292L15.439 5H4.56101L5.74131 15.2292ZM8.5 7.5C8.77614 7.5 9 7.72386 9 8V14C9 14.2761 8.77614 14.5 8.5 14.5C8.22386 14.5 8 14.2761 8 14V8C8 7.72386 8.22386 7.5 8.5 7.5ZM12 8C12 7.72386 11.7761 7.5 11.5 7.5C11.2239 7.5 11 7.72386 11 8V14C11 14.2761 11.2239 14.5 11.5 14.5C11.7761 14.5 12 14.2761 12 14V8Z" fill="currentColor"/>
        </svg>
    `;case a.Add:return i.c`
        <svg width="16" height="15" viewBox="0 0 16 15" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
          <path d="M8.39563 0.5C8.39563 0.223858 8.17177 0 7.89563 0C7.61949 0 7.39563 0.223858 7.39563 0.5V7H0.89563C0.619487 7 0.39563 7.22386 0.39563 7.5C0.39563 7.77614 0.619487 8 0.89563 8H7.39563V14.5C7.39563 14.7761 7.61949 15 7.89563 15C8.17177 15 8.39563 14.7761 8.39563 14.5V8H14.8956C15.1718 8 15.3956 7.77614 15.3956 7.5C15.3956 7.22386 15.1718 7 14.8956 7H8.39563V0.5Z" fill="${t}"/>
        </svg>`;case a.Calendar:return i.c`
        <svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M7 11C7.55228 11 8 10.5523 8 10C8 9.44771 7.55228 9 7 9C6.44772 9 6 9.44771 6 10C6 10.5523 6.44772 11 7 11ZM8 13C8 13.5523 7.55228 14 7 14C6.44772 14 6 13.5523 6 13C6 12.4477 6.44772 12 7 12C7.55228 12 8 12.4477 8 13ZM10 11C10.5523 11 11 10.5523 11 10C11 9.44771 10.5523 9 10 9C9.44771 9 9 9.44771 9 10C9 10.5523 9.44771 11 10 11ZM11 13C11 13.5523 10.5523 14 10 14C9.44771 14 9 13.5523 9 13C9 12.4477 9.44771 12 10 12C10.5523 12 11 12.4477 11 13ZM13 11C13.5523 11 14 10.5523 14 10C14 9.44771 13.5523 9 13 9C12.4477 9 12 9.44771 12 10C12 10.5523 12.4477 11 13 11ZM17 5.5C17 4.11929 15.8807 3 14.5 3H5.5C4.11929 3 3 4.11929 3 5.5V14.5C3 15.8807 4.11929 17 5.5 17H14.5C15.8807 17 17 15.8807 17 14.5V5.5ZM4 7H16V14.5C16 15.3284 15.3284 16 14.5 16H5.5C4.67157 16 4 15.3284 4 14.5V7ZM5.5 4H14.5C15.3284 4 16 4.67157 16 5.5V6H4V5.5C4 4.67157 4.67157 4 5.5 4Z" fill=${t}/>
        </svg>
    `;case a.Planner:return i.c`
        <svg width="24" height="24" fill="none" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
          <path d="M17.75 3A3.25 3.25 0 0 1 21 6.25v11.5A3.25 3.25 0 0 1 17.75 21H6.25A3.25 3.25 0 0 1 3 17.75V6.25A3.25 3.25 0 0 1 6.25 3h11.5Zm1.75 5.5h-15v9.25c0 .966.784 1.75 1.75 1.75h11.5a1.75 1.75 0 0 0 1.75-1.75V8.5Zm-11.75 6a1.25 1.25 0 1 1 0 2.5 1.25 1.25 0 0 1 0-2.5Zm4.25 0a1.25 1.25 0 1 1 0 2.5 1.25 1.25 0 0 1 0-2.5Zm-4.25-4a1.25 1.25 0 1 1 0 2.5 1.25 1.25 0 0 1 0-2.5Zm4.25 0a1.25 1.25 0 1 1 0 2.5 1.25 1.25 0 0 1 0-2.5Zm4.25 0a1.25 1.25 0 1 1 0 2.5 1.25 1.25 0 0 1 0-2.5Zm1.5-6H6.25A1.75 1.75 0 0 0 4.5 6.25V7h15v-.75a1.75 1.75 0 0 0-1.75-1.75Z" fill="${t}"/>
        </svg>`;case a.Milestone:return i.c`
        <svg width="24" height="24" fill="${t}" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
          <path d="M15.25 13c.966 0 1.75.784 1.75 1.75v4.5A1.75 1.75 0 0 1 15.25 21H3.75A1.75 1.75 0 0 1 2 19.25v-4.5c0-.966.783-1.75 1.75-1.75h11.5ZM21 14.899v5.351a.75.75 0 0 1-1.494.102l-.006-.102v-5.338a3.006 3.006 0 0 0 1.5-.013Zm-5.75-.399H3.75a.25.25 0 0 0-.25.25v4.5c0 .138.111.25.25.25h11.5a.25.25 0 0 0 .25-.25v-4.5a.25.25 0 0 0-.25-.25Zm5-4.408a1.908 1.908 0 1 1 0 3.816 1.908 1.908 0 0 1 0-3.816ZM15.244 3c.967 0 1.75.784 1.75 1.75v4.5a1.75 1.75 0 0 1-1.75 1.75h-11.5a1.75 1.75 0 0 1-1.75-1.75v-4.5a1.75 1.75 0 0 1 1.607-1.744L3.745 3h11.5Zm0 1.5h-11.5l-.057.007a.25.25 0 0 0-.193.243v4.5c0 .138.112.25.25.25h11.5a.25.25 0 0 0 .25-.25v-4.5a.25.25 0 0 0-.25-.25ZM20.25 3a.75.75 0 0 1 .743.648L21 3.75v5.351a3.004 3.004 0 0 0-1.5-.013V3.75a.75.75 0 0 1 .75-.75Z" fill="${t}"/>
        </svg>`;case a.PersonAdd:return i.c`
        <svg width="20" height="20" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg" class="svg" fill="currentColor">
          <path d="M9 2a4 4 0 100 8 4 4 0 000-8zM6 6a3 3 0 116 0 3 3 0 01-6 0z"></path>
          <path d="M4 11a2 2 0 00-2 2c0 1.7.83 2.97 2.13 3.8A9.14 9.14 0 009 18c.41 0 .82-.02 1.21-.06A5.5 5.5 0 019.6 17 12 12 0 019 17a8.16 8.16 0 01-4.33-1.05A3.36 3.36 0 013 13a1 1 0 011-1h5.6c.18-.36.4-.7.66-1H4z"></path>
          <path d="M14.5 19a4.5 4.5 0 100-9 4.5 4.5 0 000 9zm0-7c.28 0 .5.22.5.5V14h1.5a.5.5 0 010 1H15v1.5a.5.5 0 01-1 0V15h-1.5a.5.5 0 010-1H14v-1.5c0-.28.22-.5.5-.5z"></path>
        </svg>`;case a.Event:return i.c`
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
          <path d="M17.75 3C19.5449 3 21 4.45507 21 6.25V17.75C21 19.5449 19.5449 21 17.75 21H6.25C4.45507 21 3 19.5449 3 17.75V6.25C3 4.45507 4.45507 3 6.25 3H17.75ZM19.5 8.5H4.5V17.75C4.5 18.7165 5.2835 19.5 6.25 19.5H17.75C18.7165 19.5 19.5 18.7165 19.5 17.75V8.5ZM7.75 14.5C8.44036 14.5 9 15.0596 9 15.75C9 16.4404 8.44036 17 7.75 17C7.05964 17 6.5 16.4404 6.5 15.75C6.5 15.0596 7.05964 14.5 7.75 14.5ZM12 14.5C12.6904 14.5 13.25 15.0596 13.25 15.75C13.25 16.4404 12.6904 17 12 17C11.3096 17 10.75 16.4404 10.75 15.75C10.75 15.0596 11.3096 14.5 12 14.5ZM7.75 10.5C8.44036 10.5 9 11.0596 9 11.75C9 12.4404 8.44036 13 7.75 13C7.05964 13 6.5 12.4404 6.5 11.75C6.5 11.0596 7.05964 10.5 7.75 10.5ZM12 10.5C12.6904 10.5 13.25 11.0596 13.25 11.75C13.25 12.4404 12.6904 13 12 13C11.3096 13 10.75 12.4404 10.75 11.75C10.75 11.0596 11.3096 10.5 12 10.5ZM16.25 10.5C16.9404 10.5 17.5 11.0596 17.5 11.75C17.5 12.4404 16.9404 13 16.25 13C15.5596 13 15 12.4404 15 11.75C15 11.0596 15.5596 10.5 16.25 10.5ZM17.75 4.5H6.25C5.2835 4.5 4.5 5.2835 4.5 6.25V7H19.5V6.25C19.5 5.2835 18.7165 4.5 17.75 4.5Z" fill="none" />
        </svg>
      `;case a.BookOpen:return i.c`
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
          <path d="M12 19.1375C11.4986 19.6686 10.788 20 10 20H3.75C2.7835 20 2 19.2165 2 18.25V5.75C2 4.7835 2.7835 4 3.75 4H10C10.788 4 11.4986 4.33145 12 4.86253C12.5014 4.33145 13.212 4 14 4H20.25C21.2165 4 22 4.7835 22 5.75V18.25C22 19.2165 21.2165 20 20.25 20H14C13.212 20 12.5014 19.6686 12 19.1375ZM3.5 5.75V18.25C3.5 18.3881 3.61193 18.5 3.75 18.5H10C10.6904 18.5 11.25 17.9404 11.25 17.25V6.75C11.25 6.05964 10.6904 5.5 10 5.5H3.75C3.61193 5.5 3.5 5.61193 3.5 5.75ZM12.75 17.25C12.75 17.9404 13.3096 18.5 14 18.5H20.25C20.3881 18.5 20.5 18.3881 20.5 18.25V5.75C20.5 5.61193 20.3881 5.5 20.25 5.5H14C13.3096 5.5 12.75 6.05964 12.75 6.75V17.25Z" fill="none" />
        </svg>
      `;case a.FileOuter:return i.c`
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
          <path d="M6 2C4.89543 2 4 2.89543 4 4V20C4 21.1046 4.89543 22 6 22H18C19.1046 22 20 21.1046 20 20V9.82777C20 9.29733 19.7893 8.78863 19.4142 8.41355L13.5864 2.58579C13.2114 2.21071 12.7027 2 12.1722 2H6ZM5.5 4C5.5 3.72386 5.72386 3.5 6 3.5H12V8C12 9.10457 12.8954 10 14 10H18.5V20C18.5 20.2761 18.2761 20.5 18 20.5H6C5.72386 20.5 5.5 20.2761 5.5 20V4ZM17.3793 8.5H14C13.7239 8.5 13.5 8.27614 13.5 8V4.62066L17.3793 8.5Z" fill="none" />
        </svg>
      `;case a.BookQuestion:return i.c`
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
          <path d="M10.9998 8.01752C10.9905 8.42363 10.6584 8.74999 10.25 8.74999C9.5 8.74999 9.5 7.9989 9.5 7.9989L9.5 7.99777L9.50001 7.99539L9.50006 7.99017C9.50032 7.9755 9.50072 7.96084 9.50144 7.94618C9.50262 7.92198 9.50473 7.89159 9.50846 7.8559C9.51591 7.78477 9.52996 7.69092 9.55665 7.58186C9.60973 7.36492 9.71565 7.07652 9.92848 6.78906C10.3825 6.17582 11.1982 5.72727 12.513 5.7501C13.4627 5.76659 14.3059 6.16497 14.834 6.82047C15.371 7.48704 15.5517 8.3902 15.1964 9.27853C14.8342 10.1839 14.0149 10.5437 13.5442 10.7503L13.4932 10.7728C13.2147 10.8957 13.0813 10.9599 13.0013 11.024L13 11.0251L13 11.7492C13.0001 12.1634 12.6643 12.4999 12.2501 12.5C11.8359 12.5 11.5001 12.1643 11.5 11.7501L11.5 11C11.5 10.4769 11.752 10.1029 12.0633 9.85345C12.3134 9.65303 12.6276 9.51483 12.8491 9.4174L12.8875 9.40049C13.4292 9.16137 13.6868 9.01346 13.8036 8.72145C13.9483 8.35977 13.8789 8.02596 13.6659 7.76153C13.4439 7.48604 13.0371 7.25943 12.487 7.24988C11.5518 7.23364 11.2425 7.53509 11.134 7.68162C11.0656 7.77404 11.0309 7.86797 11.0137 7.93838C11.0052 7.973 11.0017 7.99908 11.0003 8.01197L10.9998 8.01752ZM12.25 15.5C12.8023 15.5 13.25 15.0523 13.25 14.5C13.25 13.9477 12.8023 13.5 12.25 13.5C11.6977 13.5 11.25 13.9477 11.25 14.5C11.25 15.0523 11.6977 15.5 12.25 15.5ZM4 4.5C4 3.11929 5.11929 2 6.5 2H18C19.3807 2 20.5 3.11929 20.5 4.5V18.75C20.5 19.1642 20.1642 19.5 19.75 19.5H5.5C5.5 20.0523 5.94772 20.5 6.5 20.5H19.75C20.1642 20.5 20.5 20.8358 20.5 21.25C20.5 21.6642 20.1642 22 19.75 22H6.5C5.11929 22 4 20.8807 4 19.5V4.5ZM5.5 4.5V18H19V4.5C19 3.94772 18.5523 3.5 18 3.5H6.5C5.94772 3.5 5.5 3.94772 5.5 4.5Z" fill="none" />
        </svg>
      `;case a.Globe:return i.c`
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
          <path d="M12.0001 1.99805C17.5238 1.99805 22.0016 6.47589 22.0016 11.9996C22.0016 17.5233 17.5238 22.0011 12.0001 22.0011C6.47638 22.0011 1.99854 17.5233 1.99854 11.9996C1.99854 6.47589 6.47638 1.99805 12.0001 1.99805ZM14.939 16.4993H9.06118C9.71322 18.9135 10.8466 20.5011 12.0001 20.5011C13.1536 20.5011 14.2869 18.9135 14.939 16.4993ZM7.5084 16.4999L4.78591 16.4998C5.74425 18.0328 7.1777 19.2384 8.88008 19.9104C8.3578 19.0906 7.92681 18.0643 7.60981 16.8949L7.5084 16.4999ZM19.2143 16.4998L16.4918 16.4999C16.168 17.8337 15.7004 18.9995 15.119 19.9104C16.716 19.2804 18.0757 18.1814 19.0291 16.7833L19.2143 16.4998ZM7.09351 9.99895H3.7359L3.73115 10.0162C3.57906 10.6525 3.49854 11.3166 3.49854 11.9996C3.49854 13.0558 3.69112 14.0669 4.0431 14.9999L7.21626 14.9995C7.07396 14.0504 6.99854 13.0422 6.99854 11.9996C6.99854 11.3156 7.031 10.6464 7.09351 9.99895ZM15.397 9.99901H8.60316C8.53514 10.6393 8.49853 11.309 8.49853 11.9996C8.49853 13.0591 8.58468 14.0694 8.73827 14.9997H15.2619C15.4155 14.0694 15.5016 13.0591 15.5016 11.9996C15.5016 11.309 15.465 10.6393 15.397 9.99901ZM20.2647 9.99811L16.9067 9.99897C16.9692 10.6464 17.0016 11.3156 17.0016 11.9996C17.0016 13.0422 16.9262 14.0504 16.7839 14.9995L19.9571 14.9999C20.309 14.0669 20.5016 13.0558 20.5016 11.9996C20.5016 11.3102 20.4196 10.64 20.2647 9.99811ZM8.88114 4.08875L8.85823 4.09747C6.81092 4.91218 5.1549 6.49949 4.25023 8.49935L7.29835 8.49972C7.61171 6.74693 8.15855 5.221 8.88114 4.08875ZM12.0001 3.49805L11.8844 3.50335C10.619 3.6191 9.39651 5.62107 8.8288 8.4993H15.1714C14.6052 5.62914 13.388 3.63033 12.1264 3.50436L12.0001 3.49805ZM15.1201 4.08881L15.2269 4.2629C15.8961 5.37537 16.4043 6.83525 16.7018 8.49972L19.7499 8.49935C18.8853 6.58795 17.3343 5.05341 15.4113 4.21008L15.1201 4.08881Z" />
        </svg>
      `;case a.PresenceAvailable:return i.c`
        <svg fill="#13a10e" aria-hidden="true" width="10" height="10" viewBox="0 0 10 10" xmlns="http://www.w3.org/2000/svg">
          <path d="M5 10A5 5 0 1 0 5 0a5 5 0 0 0 0 10Zm2.1-5.9L4.85 6.35a.5.5 0 0 1-.7 0l-1-1a.5.5 0 0 1 .7-.7l.65.64 1.9-1.9a.5.5 0 0 1 .7.71Z" fill="#13a10e"></path>
        </svg>`;case a.PresenceOofAvailable:return i.c`
        <svg fill="#13a10e" aria-hidden="true" width="10" height="10" viewBox="0 0 10 10" xmlns="http://www.w3.org/2000/svg">
            <path d="M5 0a5 5 0 1 0 0 10A5 5 0 0 0 5 0ZM1 5a4 4 0 1 1 8 0 4 4 0 0 1-8 0Zm6.1-1.6c.2.2.2.5 0 .7L4.85 6.35a.5.5 0 0 1-.7 0l-1-1a.5.5 0 1 1 .7-.7l.65.64 1.9-1.9c.2-.19.5-.19.7 0Z" fill="#13a10e"></path>
        </svg>`;case a.PresenceBusy:return i.c`
        <svg fill="#d13438" aria-hidden="true" width="10" height="10" viewBox="0 0 10 10" xmlns="http://www.w3.org/2000/svg">
          <path d="M10 5A5 5 0 1 1 0 5a5 5 0 0 1 10 0Z" fill="#d13438"></path>
        </svg>`;case a.PresenceOofBusy:return i.c`
        <svg fill="#d13438" aria-hidden="true" width="10" height="10" viewBox="0 0 10 10" xmlns="http://www.w3.org/2000/svg">
          <path d="M5 1a4 4 0 1 0 0 8 4 4 0 0 0 0-8ZM0 5a5 5 0 1 1 10 0A5 5 0 0 1 0 5Z" fill="#d13438"></path>
        </svg>
      `;case a.PresenceDnd:return i.c`
        <svg fill="#d13438" aria-hidden="true" width="10" height="10" viewBox="0 0 10 10" xmlns="http://www.w3.org/2000/svg">
          <path d="M5 10A5 5 0 1 0 5 0a5 5 0 0 0 0 10ZM3.5 4.5h3a.5.5 0 0 1 0 1h-3a.5.5 0 0 1 0-1Z" fill="#d13438"></path>
        </svg>`;case a.PresenceOofDnd:return i.c`
        <svg fill="#d13438" aria-hidden="true" width="10" height="10" viewBox="0 0 10 10" xmlns="http://www.w3.org/2000/svg">
          <path d="M5 0a5 5 0 1 0 0 10A5 5 0 0 0 5 0ZM1 5a4 4 0 1 1 8 0 4 4 0 0 1-8 0Zm2 0c0-.28.22-.5.5-.5h3a.5.5 0 0 1 0 1h-3A.5.5 0 0 1 3 5Z" fill="#d13438"></path>
        </svg>`;case a.PresenceAway:return i.c`
        <svg fill="#eaa300" aria-hidden="true" width="10" height="10" viewBox="0 0 10 10" xmlns="http://www.w3.org/2000/svg">
          <path d="M5 10A5 5 0 1 0 5 0a5 5 0 0 0 0 10Zm0-7v1.8l1.35 1.35a.5.5 0 1 1-.7.7l-1.5-1.5A.5.5 0 0 1 4 5V3a.5.5 0 0 1 1 0Z" fill="#eaa300"></path>
        </svg>`;case a.PresenceOofAway:return i.c`
        <svg fill="#c239b3" aria-hidden="true" width="10" height="10" viewBox="0 0 10 10" xmlns="http://www.w3.org/2000/svg">
          <path d="M5.35 3.85a.5.5 0 1 0-.7-.7l-1.5 1.5a.5.5 0 0 0 0 .7l1.5 1.5a.5.5 0 1 0 .7-.7L4.7 5.5h1.8a.5.5 0 1 0 0-1H4.7l.65-.65ZM5 0a5 5 0 1 0 0 10A5 5 0 0 0 5 0ZM1 5a4 4 0 1 1 8 0 4 4 0 0 1-8 0Z" fill="#c239b3"></path>
        </svg>`;case a.PresenceOffline:return i.c`
        <svg fill="#929292" aria-hidden="true" width="10" height="10" viewBox="0 0 10 10" xmlns="http://www.w3.org/2000/svg">
          <path d="M6.85 3.15c.2.2.2.5 0 .7L5.71 5l1.14 1.15a.5.5 0 1 1-.7.7L5 5.71 3.85 6.85a.5.5 0 1 1-.7-.7L4.29 5 3.15 3.85a.5.5 0 1 1 .7-.7L5 4.29l1.15-1.14c.2-.2.5-.2.7 0ZM0 5a5 5 0 1 1 10 0A5 5 0 0 1 0 5Zm5-4a4 4 0 1 0 0 8 4 4 0 0 0 0-8Z" fill="#929292"></path>
        </svg>`;case a.PresenceStatusUnknown:return i.c`
        <svg fill="#d13438" aria-hidden="true" width="10" height="10" viewBox="0 0 10 10" xmlns="http://www.w3.org/2000/svg">
          <path d="M5 1a4 4 0 1 0 0 8 4 4 0 0 0 0-8ZM0 5a5 5 0 1 1 10 0A5 5 0 0 1 0 5Z" fill="#d13438"></path>
        </svg>`;case a.Loading:return i.c`
        <svg fill="#d13438" aria-hidden="true" width="10" height="10" viewBox="0 0 10 10" xmlns="http://www.w3.org/2000/svg">
          <path d="M12 4.75a7.25 7.25 0 1 0 7.201 6.406c-.068-.588.358-1.156.95-1.156.515 0 .968.358 1.03.87a9.25 9.25 0 1 1-3.432-6.116V4.25a1 1 0 1 1 2.001 0v2.698l.034.052h-.034v.25a1 1 0 0 1-1 1h-3a1 1 0 1 1 0-2h.666A7.219 7.219 0 0 0 12 4.75Z" fill="#212121"/>
        </svg>`}}},"7SX6":function(e,t,n){"use strict";n.d(t,"b",function(){return M}),n.d(t,"c",function(){return P}),n.d(t,"a",function(){return T});var a=n("Y1A4"),i=n("HN6m"),r=n("cBsD"),o=n("zFbe"),s=n("/i08"),c=n("qqMp"),d=n("ZgG/"),l=n("h2QR"),u=n("z0DP"),f=n("glp4"),p=n("zlIh"),m=n("s9K1"),_=n("hgjj"),h=n("ZzBS"),b=n("7Cdu"),g=n("c1DA"),v=n("zCAG");const y=[c.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{font-size:var(--default-font-size)}:host .flyout [slot=anchor]{display:flex;cursor:pointer;align-items:var(--person-alignment,normal)}:host .flyout [slot=anchor].vertical{flex-direction:column;justify-content:center;align-items:center;margin-inline-start:0;--person-avatar-size:72px}:host .person-root{display:flex;flex-direction:row;align-items:center;background-color:var(--person-background-color,transparent);border-radius:var(--person-border-radius,4px)}:host .person-root.small .avatar-wrapper{min-width:var(--person-avatar-size,24px);width:var(--person-avatar-size,24px);height:var(--person-avatar-size,24px)}:host .person-root.noline .presence-basic,:host .person-root.oneline .presence-basic{border-width:1px;position:relative;bottom:calc(var(--person-avatar-size,24px) * .12 - 4px)}:host .person-root.twolines .avatar-wrapper{min-width:var(--person-avatar-size,40px);width:var(--person-avatar-size,40px);height:var(--person-avatar-size,40px)}:host .person-root.twolines .avatar-wrapper .contact-icon,:host .person-root.twolines .avatar-wrapper .initials{font-size:calc(var(--person-avatar-size,40px) * .4)}:host .person-root.twolines .avatar-wrapper .presence-wrapper svg{width:calc(var(--person-avatar-size,40px) * .28);height:calc(var(--person-avatar-size,40px) * .28)}:host .person-root.large .avatar-wrapper,:host .person-root.threelines .avatar-wrapper{min-width:var(--person-avatar-size,56px);width:var(--person-avatar-size,56px);height:var(--person-avatar-size,56px)}:host .person-root.large .avatar-wrapper .contact-icon,:host .person-root.large .avatar-wrapper .initials,:host .person-root.threelines .avatar-wrapper .contact-icon,:host .person-root.threelines .avatar-wrapper .initials{font-size:calc(var(--person-avatar-size,56px) * .4)}:host .person-root.large .avatar-wrapper .presence-wrapper svg,:host .person-root.threelines .avatar-wrapper .presence-wrapper svg{width:calc(var(--person-avatar-size,56px) * .28);height:calc(var(--person-avatar-size,56px) * .28)}:host .person-root.fourlines .avatar-wrapper{min-width:var(--person-avatar-size,72px);width:var(--person-avatar-size,72px);height:var(--person-avatar-size,72px)}:host .person-root.fourlines .avatar-wrapper .contact-icon,:host .person-root.fourlines .avatar-wrapper .initials{font-size:calc(var(--person-avatar-size,72px) * .4)}:host .person-root.fourlines .avatar-wrapper .presence-wrapper svg{width:calc(var(--person-avatar-size,72px) * .28);height:calc(var(--person-avatar-size,72px) * .28)}:host .person-root.vertical{flex-direction:column;justify-content:center;align-items:center}:host .person-root.vertical .avatar-wrapper{min-width:var(--person-avatar-size,72px);width:var(--person-avatar-size,72px);height:var(--person-avatar-size,72px)}:host .person-root.vertical .avatar-wrapper .contact-icon,:host .person-root.vertical .avatar-wrapper .initials{font-size:calc(var(--person-avatar-size,72px) * .4)}:host .person-root.vertical .avatar-wrapper .presence-wrapper svg{width:calc(var(--person-avatar-size,72px) * .28);height:calc(var(--person-avatar-size,72px) * .28)}:host .person-root .avatar-wrapper{min-width:var(--person-avatar-size,24px);width:var(--person-avatar-size,24px);height:var(--person-avatar-size,24px);position:relative;box-sizing:border-box}:host .person-root .avatar-wrapper .contact-icon,:host .person-root .avatar-wrapper .initials,:host .person-root .avatar-wrapper img{height:100%;width:100%;border:var(--person-avatar-border,none);border-radius:var(--person-avatar-border-radius,50%);margin-block-start:var(--person-avatar-top-spacing,0)}:host .person-root .avatar-wrapper .contact-icon,:host .person-root .avatar-wrapper .initials{display:flex;justify-content:center;align-items:center;font-size:calc(var(--person-avatar-size,24px) * .4);font-weight:400;background:var(--person-initials-background-color,var(--neutral-fill-secondary-rest));color:var(--person-initials-text-color,var(--neutral-fill-strong-hover))}:host .person-root .avatar-wrapper .presence-wrapper{bottom:var(--person-presence-wrapper-bottom,0);right:0;position:absolute;border-radius:50%;background-color:var(--neutral-layer-1);border:1px solid var(--neutral-layer-1)}:host .person-root .avatar-wrapper .presence-wrapper svg{display:flex;width:calc(var(--person-avatar-size,24px) * .28);height:calc(var(--person-avatar-size,24px) * .28)}:host .person-root .details-wrapper{display:flex;flex-direction:column;align-items:flex-start;min-width:var(--person-details-wrapper-width,168px);margin-inline-start:var(--person-details-left-spacing,12px);margin-block-end:var(--person-details-bottom-spacing,0)}:host .person-root .details-wrapper.vertical{display:inline-flex;flex-direction:column;justify-content:center;align-items:center;margin-inline-start:0}:host .person-root .details-wrapper .line1{font-size:var(--person-line1-font-size,ms-font-size-14);font-weight:var(--person-line1-font-weight,600);color:var(--person-line1-text-color,var(--neutral-foreground-rest));text-transform:var(--person-line1-text-transform,inherit);line-height:var(--person-line1-text-line-height,20px)}:host .person-root .details-wrapper .line2{font-size:var(--person-line2-font-size,var(--email-font-size,ms-font-size-12));font-weight:var(--person-line2-font-weight,400);color:var(--person-line2-text-color,var(--secondary-text-color));text-transform:var(--person-line2-text-transform,inherit);line-height:var(--person-line2-text-line-height,16px)}:host .person-root .details-wrapper .line3{font-size:var(--person-line3-font-size,var(--email-font-size,ms-font-size-12));font-weight:var(--person-line3-font-weight,400);color:var(--person-line3-text-color,var(--secondary-text-color));text-transform:var(--person-line3-text-transform,inherit);line-height:var(--person-line3-text-line-height,16px)}:host .person-root .details-wrapper .line4{font-size:var(--person-line4-font-size,var(--email-font-size,ms-font-size-12));font-weight:var(--person-line4-font-weight,400);color:var(--person-line4-text-color,var(--secondary-text-color));text-transform:var(--person-line4-text-transform,inherit);line-height:var(--person-line4-text-line-height,16px)}@media (forced-colors:active) and (prefers-color-scheme:dark){:host svg,:host svg>path{fill:#fff;fill-rule:nonzero;clip-rule:nonzero}}[dir=rtl] .presence-wrapper{right:unset!important;left:0}
`];var S=n("dP6N");const D=Object.assign(Object.assign({},{Available:"Available",Away:"Away",BeRightBack:"Be right back",Busy:"Busy",DoNotDisturb:"Do not disturb",InACall:"In a call",InAConferenceCall:"In a conference call",Inactive:"Inactive",InAMeeting:"In a meeting",Offline:"Offline",OffWork:"Off work",OutOfOffice:"Out of office",PresenceUnknown:"Presence unknown",Presenting:"Presenting",UrgentInterruptionsOnly:"Urgent interruptions only"}),{photoFor:"Photo for",emailAddress:"Email address",initials:"Initials"});var I=n("fU9z"),x=n("ox5k"),C=n("0mOt"),O=n("R2TB"),w=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};class E{static get config(){return this.presenceConfig}static register(e){null!==this.initPromise&&void 0!==this.initPromise||(this.initPromise=this.init()),this.initPromise.then(()=>{setTimeout(()=>{this.registerPromise&&this.registerPromise(null)},this.config.initial),this.components.add(e)},e=>Object(O.a)("the PresenceService could not be initialized; presence will not be updated",e))}static unregister(e){this.components.delete(e)}static stop(){this.isStopped=!0}static init(){return w(this,void 0,void 0,function*(){return(()=>{w(this,void 0,void 0,function*(){for(;!this.isStopped;){const e=new Promise(e=>this.registerPromise=e),t=new Promise(e=>setTimeout(e,this.config.refresh));yield Promise.race([e,t]);const n=o.a.globalProvider;if(!n||n.state===s.c.Loading||n.state===s.c.SignedOut)continue;Object(O.b)(`updating presence for ${this.components.size} components.`);const a=new Map;for(const e of this.components){const t=e.presenceId;t&&!a.has(t)&&a.set(t,void 0)}const i=[];try{const e=yield Object(p.b)(n.graph,Array.from(a.keys()).map(e=>({id:e})),!0);if(e)for(const t of Object.values(e))t.id&&(a.set(t.id,t),i.push(`${t.id}=${t.availability}/${t.activity}`))}catch(e){Object(O.a)("could not update presence",e);continue}Object(O.b)(`updated presence for ${i.length} users.`,i);for(const e of this.components){const t=e.presenceId;if(t){const n=a.get(t);n&&e.onPresenceChange(n)}}}})})(),Promise.resolve()})}}E.components=new Set,E.initPromise=null,E.isStopped=!1,E.presenceConfig={initial:1e3,refresh:3e4};var A=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},L=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)},k=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const M=["businessPhones","displayName","givenName","jobTitle","department","mail","mobilePhone","officeLocation","preferredLanguage","surname","userPrincipalName","id","userType"],P=()=>{Object(C.b)("person",T),Object(g.a)()};class T extends a.a{static get styles(){return y}get strings(){return D}get personQuery(){return this._personQuery}set personQuery(e){e!==this._personQuery&&(this._personQuery=e,this.personDetailsInternal=null,this.requestStateUpdate())}get fallbackDetails(){return this._fallbackDetails}set fallbackDetails(e){e!==this._fallbackDetails&&(this._fallbackDetails=e,this.personDetailsInternal||this.requestStateUpdate())}get userId(){return this._userId}set userId(e){e!==this._userId&&(this._userId=e,this.personDetailsInternal=null,this.requestStateUpdate())}get usage(){return this._usage}set usage(e){e!==this._usage&&(this._usage=e,this.requestStateUpdate())}get personDetailsInternal(){return this._personDetailsInternal}set personDetailsInternal(e){this._personDetailsInternal!==e&&(this._personDetailsInternal=e,this._fetchedImage=null,this._fetchedPresence=null,this.requestStateUpdate(),this.requestUpdate("personDetailsInternal"))}get personDetails(){return this._personDetails}set personDetails(e){this._personDetails!==e&&(this._personDetails=e,this._fetchedImage=null,this._fetchedPresence=null,this.requestStateUpdate(),this.requestUpdate("personDetails"))}get personImage(){return this._personImage||this._fetchedImage}set personImage(e){if(e===this._personImage)return;this._isInvalidImageSrc=!e;const t=this._personImage;this._personImage=e,this.requestUpdate("personImage",t)}get avatarType(){return this._avatarType}set avatarType(e){e!==this._avatarType&&(this._avatarType=e,this.requestStateUpdate())}get personPresence(){return this._personPresence||this._fetchedPresence}set personPresence(e){if(e===this._personPresence)return;const t=this._personPresence;this._personPresence=e,this.requestUpdate("personPresence",t)}static get requiredScopes(){const e=["user.readbasic.all","user.read","people.read","presence.read.all","presence.read"];return T.config.useContactApis&&e.push("contacts.read"),e}get flyout(){return this.renderRoot.querySelector(".flyout")}constructor(){super(),this._hasLoadedPersonCard=!1,this._mouseLeaveTimeout=-1,this._mouseEnterTimeout=-1,this.handleMouseClick=e=>{const t=e.target;this.personCardInteraction===v.a.click&&t.tagName!==`${i.a.prefix}-PERSON-CARD`.toUpperCase()&&this.showPersonCard()},this.handleKeyDown=e=>{e&&"Enter"===e.key&&this.showPersonCard()},this.handleMouseEnter=()=>{clearTimeout(this._mouseEnterTimeout),clearTimeout(this._mouseLeaveTimeout),this.personCardInteraction===v.a.hover&&(this._mouseEnterTimeout=window.setTimeout(this.showPersonCard,500))},this.handleMouseLeave=()=>{clearTimeout(this._mouseEnterTimeout),clearTimeout(this._mouseLeaveTimeout),this._mouseLeaveTimeout=window.setTimeout(this.hidePersonCard,500)},this.hidePersonCard=()=>{const e=this.flyout;e&&e.close();const t=this.querySelector(".mgt-person-card")||this.renderRoot.querySelector(".mgt-person-card");t&&(t.isExpanded=!1,t.clearHistory())},this.loadPersonCardResources=()=>k(this,void 0,void 0,function*(){if(this.personCardInteraction!==v.a.none&&!this._hasLoadedPersonCard){const{registerMgtPersonCardComponent:e}=yield Promise.resolve().then(n.bind(null,"uVfA"));customElements.get(Object(C.a)("person-card"))||e(),this._hasLoadedPersonCard=!0}}),this.showPersonCard=()=>{this._personCardShouldRender||(this._personCardShouldRender=!0,this.loadPersonCardResources());const e=this.flyout;e&&e.open()},this.personCardInteraction=v.a.none,this.line1Property="displayName",this.line2Property="jobTitle",this.line3Property="department",this.line4Property="email",this.view=h.a.image,this.avatarSize="auto",this.disableImageFetch=!1,this._isInvalidImageSrc=!1,this._avatarType=S.b.photo,this.verticalLayout=!1}disconnectedCallback(){this.showPresence&&!this.disablePresenceRefresh&&E.unregister(this),super.disconnectedCallback()}render(){if(this.isLoadingState&&!this.personDetails&&!this.personDetailsInternal&&!this.fallbackDetails)return this.renderLoading();const e=this.personDetails||this.personDetailsInternal||this.fallbackDetails,t=this.getImage(),n=this.personPresence||this._fetchedPresence;if(!e&&!t)return this.renderNoData();!(null==e?void 0:e.personImage)&&t&&(e.personImage=t);let a=this.renderTemplate("default",{person:e,personImage:t,personPresence:n});if(!a){const i=this.renderDetails(e,n),r=this.renderAvatar(e,t,n);a=c.c`
        ${r}
        ${i}
      `}this.personCardInteraction!==v.a.none&&(a=this.renderFlyout(a,e,t,n));const i=Object(l.a)({"person-root":!0,small:!this.isThreeLines()&&!this.isFourLines()&&!this.isLargeAvatar(),large:"auto"!==this.avatarSize&&this.isLargeAvatar(),noline:this.isNoLine(),oneline:this.isOneLine(),twolines:this.isTwoLines(),threelines:this.isThreeLines(),fourlines:this.isFourLines(),vertical:this.isVertical()});return c.c`
      <div
        class="${i}"
        dir=${this.direction}
        @click=${this.handleMouseClick}
        @mouseenter=${this.handleMouseEnter}
        @mouseleave=${this.handleMouseLeave}
        @keydown=${this.handleKeyDown}
        tabindex="${Object(x.a)(this.personCardInteraction!==v.a.none?"0":void 0)}"
      >
        ${a}
      </div>
    `}renderLoading(){const e=this.renderTemplate("loading",null);if(e)return e;const t={"avatar-icon":!0,vertical:this.isVertical(),small:!this.isLargeAvatar(),threeLines:this.isThreeLines(),fourLines:this.isFourLines()};return c.c`
      <i class=${Object(l.a)(t)} icon='loading'>${this.renderLoadingIcon()}</i>
    `}clearState(){this._personImage="",this._personDetailsInternal=null,this._fetchedImage=null,this._fetchedPresence=null}renderNoData(){const e=this.renderTemplate("no-data",null);if(e)return e;const t={"avatar-icon":!0,vertical:this.isVertical(),small:!this.isLargeAvatar(),threeLines:this.isThreeLines(),fourLines:this.isFourLines()};return c.c`
      <i class=${Object(l.a)(t)} icon='no-data'>${this.renderPersonIcon()}</i>
    `}renderLoadingIcon(){return Object(b.b)(b.a.Loading)}renderPersonIcon(){return Object(b.b)(b.a.Person)}renderImage(e,t){var n;const a=`${this.strings.photoFor} ${e.displayName}`,i=t&&!this._isInvalidImageSrc&&this.avatarType===S.b.photo,r=this.avatarType===S.b.photo&&this.view===h.a.image,o=null!==(n=(null==e?void 0:e.displayName)||Object(u.e)(e))&&void 0!==n?n:void 0,s=c.c`<img
      title="${Object(x.a)(r?o:void 0)}"
      alt=${a}
      src=${t}
      @error=${()=>this._isInvalidImageSrc=!0} />`,d=e?this.getInitials(e):"",f=null==d?void 0:d.length,p=Object(l.a)({initials:f&&!i,"contact-icon":!f}),m=c.c`<i>${this.renderPersonIcon()}</i>`,_=c.c`
      <span 
        title="${Object(x.a)(this.view===h.a.image?o:void 0)}"
        role="${Object(x.a)(this.view===h.a.image?void 0:"presentation")}"
        class="${p}"
      >
        ${f?d:m}
      </span>
`;return i?s:_}renderPresence(e){var t;if(!this.showPresence||!e)return c.c``;let n;const{activity:a,availability:i}=e;switch(i){case"Available":case"AvailableIdle":n="OutOfOffice"===a?Object(b.b)(b.a.PresenceOofAvailable):Object(b.b)(b.a.PresenceAvailable);break;case"Busy":case"BusyIdle":switch(a){case"OutOfOffice":case"OnACall":n=Object(b.b)(b.a.PresenceOofBusy);break;default:n=Object(b.b)(b.a.PresenceBusy)}break;case"DoNotDisturb":n="OutOfOffice"===a?Object(b.b)(b.a.PresenceOofDnd):Object(b.b)(b.a.PresenceDnd);break;case"Away":n="OutOfOffice"===a?Object(b.b)(b.a.PresenceOofAway):Object(b.b)(b.a.PresenceAway);break;case"BeRightBack":n=Object(b.b)(b.a.PresenceAway);break;case"Offline":switch(a){case"Offline":n=Object(b.b)(b.a.PresenceOffline);break;case"OutOfOffice":case"OffWork":n=Object(b.b)(b.a.PresenceOofAway);break;default:n=Object(b.b)(b.a.PresenceStatusUnknown)}break;default:n=Object(b.b)(b.a.PresenceStatusUnknown)}const r=Object(l.a)({"presence-wrapper":!0,noline:this.isNoLine(),oneline:this.isOneLine()}),o=null!==(t=this.strings[a])&&void 0!==t?t:c.d;return c.c`
      <span
        class="${r}"
        title="${o}"
        aria-label="${o}"
        role="img">
          ${n}
      </span>
    `}renderAvatar(e,t,n){let a="";if(t&&!this._isInvalidImageSrc&&this._avatarType!==S.b.initials||!e?(a=e&&e.displayName||"",""!==a&&(a=`${this.strings.photoFor} ${a}`)):a=`${this.strings.initials} ${this.getInitials(e)}`,""===a){const t=Object(u.e)(e);null!==t&&(a=`${this.strings.emailAddress} ${t}`)}const i=this.renderImage(e,t),r=this.renderPresence(n);return c.c`
      <div class="avatar-wrapper">
        ${i}
        ${r}
      </div>
    `}handleLine1Clicked(){this.fireCustomEvent("line1clicked",this.personDetailsInternal)}handleLine2Clicked(){this.fireCustomEvent("line2clicked",this.personDetailsInternal)}handleLine3Clicked(){this.fireCustomEvent("line3clicked",this.personDetailsInternal)}handleLine4Clicked(){this.fireCustomEvent("line4clicked",this.personDetailsInternal)}renderDetails(e,t){if(!e||this.view===h.a.image||this.view===S.a.avatar)return c.c``;const n=e;t&&(n.presenceActivity=null==t?void 0:t.activity,n.presenceAvailability=null==t?void 0:t.availability);const a=[];if(this.view>h.a.image){const e=this.getTextFromProperty(n,this.line1Property);if(this.hasTemplate("line1")){const t=this.renderTemplate("line1",{person:n});a.push(c.c`
           <div class="line1" @click=${()=>this.handleLine1Clicked()} role="presentation" aria-label="${e}">${t}</div>
         `)}else e&&a.push(c.c`
             <div class="line1" @click=${()=>this.handleLine1Clicked()} role="presentation" aria-label="${e}">${e}</div>
           `)}if(this.view>h.a.oneline){const e=this.getTextFromProperty(n,this.line2Property);if(this.hasTemplate("line2")){const t=this.renderTemplate("line2",{person:n});a.push(c.c`
           <div class="line2" @click=${()=>this.handleLine2Clicked()} role="presentation" aria-label="${e}">${t}</div>
         `)}else e&&a.push(c.c`
             <div class="line2" @click=${()=>this.handleLine2Clicked()} role="presentation" aria-label="${e}">${e}</div>
           `)}if(this.view>h.a.twolines){const e=this.getTextFromProperty(n,this.line3Property);if(this.hasTemplate("line3")){const t=this.renderTemplate("line3",{person:n});a.push(c.c`
           <div class="line3" @click=${()=>this.handleLine3Clicked()} role="presentation" aria-label="${e}">${t}</div>
         `)}else e&&a.push(c.c`
             <div class="line3" @click=${()=>this.handleLine3Clicked()} role="presentation" aria-label="${e}">${e}</div>
           `)}if(this.view>h.a.threelines){const e=this.getTextFromProperty(n,this.line4Property);if(this.hasTemplate("line4")){const t=this.renderTemplate("line4",{person:n});a.push(c.c`
          <div class="line4" @click=${()=>this.handleLine4Clicked()} role="presentation" aria-label="${e}">${t}</div>
        `)}else e&&a.push(c.c`
            <div class="line4" @click=${()=>this.handleLine4Clicked()} role="presentation" aria-label="${e}">${e}</div>
          `)}const i=Object(l.a)({"details-wrapper":!0,vertical:this.isVertical()});return c.c`
      <div class="${i}">
        ${a}
      </div>
    `}renderFlyout(e,t,n,a){const i=this._personCardShouldRender&&this._hasLoadedPersonCard?c.c`
           <div slot="flyout" data-testid="flyout-slot">
             ${this.renderFlyoutContent(t,n,a)}
           </div>`:c.c``,o=Object(l.a)({vertical:this.isVertical()});return r.a`
      <mgt-flyout light-dismiss class="flyout" .avoidHidingAnchor=${!1}>
        <div slot="anchor" class="${o}">${e}</div>
        ${i}
      </mgt-flyout>`}renderFlyoutContent(e,t,n){return this.renderTemplate("person-card",{person:e,personImage:t})||r.a`
        <mgt-person-card
          class="mgt-person-card"
          lock-tab-navigation
          .personDetails=${e}
          .personImage=${t}
          .personPresence=${n}
          .showPresence=${this.showPresence}>
        </mgt-person-card>`}loadState(){return k(this,void 0,void 0,function*(){const e=o.a.globalProvider;if(!e||e.state===s.c.Loading)return;if(e&&e.state===s.c.SignedOut)return void(this.personDetailsInternal=null);const t=e.graph.forComponent(this);(this.verticalLayout&&this.view<h.a.fourlines||this.fallbackDetails)&&(this.line2Property="email");let n=[...M,this.line1Property,this.line2Property,this.line3Property,this.line4Property];n=n.filter(e=>"email"!==e);let a=this.personDetailsInternal||this.personDetails;if(a){if(!a.personImage&&this.fetchImage&&this._avatarType===S.b.photo&&!this.personImage&&!this._fetchedImage){let e;e="groupTypes"in a?yield Object(f.b)(t,a):yield Object(f.d)(t,a,T.config.useContactApis),e&&(a.personImage=e,this._fetchedImage=e)}}else if(this.userId||"me"===this.personQuery){let e;e=this._avatarType!==S.b.photo||this.disableImageFetch?"me"===this.personQuery?yield Object(m.e)(t,n):yield Object(m.f)(t,this.userId,n):yield Object(_.a)(t,this.userId,n),this.personDetailsInternal=e,this.personDetails=e,this._fetchedImage=this.getImage()}else if(this.personQuery){let e=yield Object(u.d)(t,this.personQuery,1);if(e&&0!==e.length||(e=(yield Object(m.b)(t,this.personQuery,1))||[]),(null==e?void 0:e.length)&&(this.personDetailsInternal=e[0],this.personDetails=e[0],this._avatarType===S.b.photo&&!this.disableImageFetch)){const n=yield Object(f.d)(t,e[0],T.config.useContactApis);n&&(this.personDetailsInternal.personImage=n,this.personDetails.personImage=n,this._fetchedImage=n)}}a=this.personDetailsInternal||this.personDetails||this.fallbackDetails;const i={activity:"Offline",availability:"Offline",id:null};if(this.showPresence&&!this.disablePresenceRefresh)this._fetchedPresence=i,E.register(this);else if(this.showPresence&&!this.personPresence&&!this._fetchedPresence)try{if(a){const e="me"!==this.personQuery?null==a?void 0:a.id:null;this._fetchedPresence=yield Object(p.a)(t,e)}else this._fetchedPresence=i}catch(e){this._fetchedPresence=i}})}getInitials(e){var t,n,a,i,r,o;if(e||(e=this.personDetailsInternal),Object(I.a)(e))return e.initials;let s="";if(Object(I.c)(e)&&(s+=null!==(a=null===(n=null===(t=e.givenName)||void 0===t?void 0:t[0])||void 0===n?void 0:n.toUpperCase())&&void 0!==a?a:"",s+=null!==(o=null===(r=null===(i=e.surname)||void 0===i?void 0:i[0])||void 0===r?void 0:r.toUpperCase())&&void 0!==o?o:""),!s&&e.displayName){const t=e.displayName.split(/\s+/);for(let e=0;e<2&&e<t.length;e++)t[e][0]&&this.isLetter(t[e][0])&&(s+=t[e][0].toUpperCase())}return s}getImage(){if(this.personImage)return this.personImage;if(this._fetchedImage)return this._fetchedImage;const e=this.personDetailsInternal||this.personDetails;return(null==e?void 0:e.personImage)?e.personImage:null}isLetter(e){try{return e.match(new RegExp("\\p{L}","u"))}catch(t){return e.toLowerCase()!==e.toUpperCase()}}getTextFromProperty(e,t){if(!t||0===t.length)return null;const n=t.trim().split(",");let a,i=0;for(;!a&&i<n.length;){const t=n[i].trim();switch(t){case"mail":case"email":a=Object(u.e)(e);break;default:a=e[t]}i++}return a}isLargeAvatar(){return"large"===this.avatarSize||"auto"===this.avatarSize&&this.view>h.a.oneline}isNoLine(){return this.view<h.a.oneline}isOneLine(){return this.view===h.a.oneline}isTwoLines(){return this.view===h.a.twolines}isThreeLines(){return this.view===h.a.threelines}isFourLines(){return this.view===h.a.fourlines}isVertical(){return this.verticalLayout}get presenceId(){var e,t,n;return(null===(e=this.personDetailsInternal)||void 0===e?void 0:e.id)||(null===(t=this.personDetails)||void 0===t?void 0:t.id)||(null===(n=this.fallbackDetails)||void 0===n?void 0:n.id)}onPresenceChange(e){this.personPresence=e}}T.config={useContactApis:!0},A([Object(d.b)({attribute:"person-query"}),L("design:type",String),L("design:paramtypes",[String])],T.prototype,"personQuery",null),A([Object(d.b)({attribute:"fallback-details",type:Object}),L("design:type",Object),L("design:paramtypes",[Object])],T.prototype,"fallbackDetails",null),A([Object(d.b)({attribute:"user-id"}),L("design:type",String),L("design:paramtypes",[String])],T.prototype,"userId",null),A([Object(d.b)({attribute:"usage"}),L("design:type",String),L("design:paramtypes",[String])],T.prototype,"usage",null),A([Object(d.b)({attribute:"show-presence",type:Boolean}),L("design:type",Boolean)],T.prototype,"showPresence",void 0),A([Object(d.b)({attribute:"disable-presence-refresh",type:Boolean}),L("design:type",Object)],T.prototype,"disablePresenceRefresh",void 0),A([Object(d.b)({attribute:"avatar-size",type:String}),L("design:type",String)],T.prototype,"avatarSize",void 0),A([Object(d.c)(),L("design:type",Object),L("design:paramtypes",[Object])],T.prototype,"personDetailsInternal",null),A([Object(d.b)({attribute:"person-details",type:Object}),L("design:type",Object),L("design:paramtypes",[Object])],T.prototype,"personDetails",null),A([Object(d.b)({attribute:"person-image",type:String}),L("design:type",String),L("design:paramtypes",[String])],T.prototype,"personImage",null),A([Object(d.b)({attribute:"fetch-image",type:Boolean}),L("design:type",Boolean)],T.prototype,"fetchImage",void 0),A([Object(d.b)({attribute:"disable-image-fetch",type:Boolean}),L("design:type",Boolean)],T.prototype,"disableImageFetch",void 0),A([Object(d.b)({attribute:"vertical-layout",type:Boolean}),L("design:type",Boolean)],T.prototype,"verticalLayout",void 0),A([Object(d.b)({attribute:"avatar-type",converter:e=>"initials"===(e=e.toLowerCase())?S.b.initials:S.b.photo}),L("design:type",String),L("design:paramtypes",[String])],T.prototype,"avatarType",null),A([Object(d.b)({attribute:"person-presence",type:Object}),L("design:type",Object),L("design:paramtypes",[Object])],T.prototype,"personPresence",null),A([Object(d.b)({attribute:"person-card",converter:e=>(e=e.toLowerCase(),void 0===v.a[e]?v.a.none:v.a[e])}),L("design:type",Number)],T.prototype,"personCardInteraction",void 0),A([Object(d.b)({attribute:"line1-property"}),L("design:type",String)],T.prototype,"line1Property",void 0),A([Object(d.b)({attribute:"line2-property"}),L("design:type",String)],T.prototype,"line2Property",void 0),A([Object(d.b)({attribute:"line3-property"}),L("design:type",String)],T.prototype,"line3Property",void 0),A([Object(d.b)({attribute:"line4-property"}),L("design:type",String)],T.prototype,"line4Property",void 0),A([Object(d.b)({converter:e=>e&&0!==e.length?(e=e.toLowerCase(),void 0===h.a[e]?h.a.image:h.a[e]):h.a.image}),L("design:type",Number)],T.prototype,"view",void 0),A([Object(d.c)(),L("design:type",String)],T.prototype,"_fetchedImage",void 0),A([Object(d.c)(),L("design:type",Object)],T.prototype,"_fetchedPresence",void 0),A([Object(d.c)(),L("design:type",Boolean)],T.prototype,"_isInvalidImageSrc",void 0),A([Object(d.c)(),L("design:type",Boolean)],T.prototype,"_personCardShouldRender",void 0),A([Object(d.c)(),L("design:type",Object)],T.prototype,"_hasLoadedPersonCard",void 0)},"81Gc":function(e,t,n){"use strict";n.d(t,"f",function(){return f}),n.d(t,"c",function(){return p}),n.d(t,"a",function(){return m}),n.d(t,"b",function(){return _}),n.d(t,"d",function(){return h}),n.d(t,"e",function(){return b});var a=n("4X57"),i=n("j9Xn"),r=n("xY0q"),o=n("wHpb"),s=n("oqLQ"),c=n("QkLF"),d=n("8hiW"),l=n("NLdk"),u=n("FVZ7");const f=(e,t,n,i="[disabled]")=>a.a`
    ${Object(r.a)("inline-flex")}
    
    :host {
      position: relative;
      box-sizing: border-box;
      ${l.a}
      height: calc(${c.a} * 1px);
      min-width: calc(${c.a} * 1px);
      color: ${d.fb};
      border-radius: calc(${d.q} * 1px);
      fill: currentcolor;
    }

    .control {
      border: calc(${d.vb} * 1px) solid transparent;
      flex-grow: 1;
      box-sizing: border-box;
      display: inline-flex;
      justify-content: center;
      align-items: center;
      padding: 0 calc((10 + (${d.s} * 2 * ${d.r})) * 1px);
      white-space: nowrap;
      outline: none;
      text-decoration: none;
      color: inherit;
      border-radius: inherit;
      fill: inherit;
      font-family: inherit;
    }

    .control,
    .end,
    .start {
      font: inherit;
    }

    .control.icon-only {
      padding: 0;
      line-height: 0;
    }

    .control:${o.a} {
      ${u.a}
    }

    .control::-moz-focus-inner {
      border: 0;
    }

    .content {
      pointer-events: none;
    }

    .start,
    .end {
      display: flex;
      pointer-events: none;
    }

    .start {
      margin-inline-end: 11px;
    }

    .end {
      margin-inline-start: 11px;
    }
  `,p=(e,t,n,r="[disabled]")=>a.a`
    .control {
      background: padding-box linear-gradient(${d.Q}, ${d.Q}),
        border-box ${d.lb};
    }

    :host(${n}:hover) .control {
      background: padding-box linear-gradient(${d.G}, ${d.G}),
        border-box ${d.kb};
    }

    :host(${n}:active) .control {
      background: padding-box linear-gradient(${d.F}, ${d.F}),
        border-box ${d.jb};
    }

    :host(${r}) .control {
      background: padding-box linear-gradient(${d.Q}, ${d.Q}),
        border-box ${d.rb};
    }
  `.withBehaviors(Object(s.a)(a.a`
        .control {
          background: ${i.a.ButtonFace};
          border-color: ${i.a.ButtonText};
          color: ${i.a.ButtonText};
        }

        :host(${n}:hover) .control,
        :host(${n}:active) .control {
          forced-color-adjust: none;
          background: ${i.a.HighlightText};
          border-color: ${i.a.Highlight};
          color: ${i.a.Highlight};
        }

        :host(${r}) .control {
          background: transparent;
          border-color: ${i.a.GrayText};
          color: ${i.a.GrayText};
        }

        .control:${o.a} {
          outline-color: ${i.a.CanvasText};
        }

        :host([href]) .control {
          background: transparent;
          border-color: ${i.a.LinkText};
          color: ${i.a.LinkText};
        }

        :host([href]:hover) .control,
        :host([href]:active) .control {
          background: transparent;
          border-color: ${i.a.CanvasText};
          color: ${i.a.CanvasText};
        }
    `)),m=(e,t,n,r="[disabled]")=>a.a`
    .control {
      background: padding-box linear-gradient(${d.e}, ${d.e}),
        border-box ${d.m};
      color: ${d.C};
    }

    :host(${n}:hover) .control {
      background: padding-box linear-gradient(${d.d}, ${d.d}),
        border-box ${d.l};
      color: ${d.B};
    }

    :host(${n}:active) .control {
      background: padding-box linear-gradient(${d.b}, ${d.b}),
        border-box ${d.j};
      color: ${d.z};
    }

    :host(${r}) .control {
      background: ${d.e};
    }

    .control:${o.a} {
      box-shadow: 0 0 0 calc(${d.y} * 1px) ${d.w} inset !important;
    }
  `.withBehaviors(Object(s.a)(a.a`
        .control {
          forced-color-adjust: none;
          background: ${i.a.Highlight};
          color: ${i.a.HighlightText};
        }

        :host(${n}:hover) .control,
        :host(${n}:active) .control {
          background: ${i.a.HighlightText};
          border-color: ${i.a.Highlight};
          color: ${i.a.Highlight};
        }

        :host(${r}) .control {
          background: transparent;
          border-color: ${i.a.GrayText};
          color: ${i.a.GrayText};
        }

        .control:${o.a} {
          outline-color: ${i.a.CanvasText};
          box-shadow: 0 0 0 calc(${d.y} * 1px) ${i.a.HighlightText} inset !important;
        }

        :host([href]) .control {
          background: ${i.a.LinkText};
          color: ${i.a.HighlightText};
        }

        :host([href]:hover) .control,
        :host([href]:active) .control {
          background: ${i.a.ButtonFace};
          border-color: ${i.a.LinkText};
          color: ${i.a.LinkText};
        }
      `)),_=(e,t,n,r="[disabled]")=>a.a`
    :host {
      color: ${d.i};
    }

    .control {
      background: ${d.ab};
    }

    :host(${n}:hover) .control {
      background: ${d.Y};
      color: ${d.h};
    }

    :host(${n}:active) .control {
      background: ${d.W};
      color: ${d.f};
    }

    :host(${r}) .control {
      background: ${d.ab};
    }
  `.withBehaviors(Object(s.a)(a.a`
        :host {
          color: ${i.a.ButtonText};
        }

        .control {
          forced-color-adjust: none;
          background: transparent;
        }

        :host(${n}:hover) .control,
        :host(${n}:active) .control {
          background: transparent;
          border-color: ${i.a.ButtonText};
          color: ${i.a.ButtonText};
        }

        :host(${r}) .control {
          background: transparent;
          color: ${i.a.GrayText};
        }

        .control:${o.a} {
          outline-color: ${i.a.CanvasText};
        }

        :host([href]) .control {
          color: ${i.a.LinkText};
        }

        :host([href]:hover) .control,
        :host([href]:active) .control {
          border-color: ${i.a.LinkText};
          color: ${i.a.LinkText};
        }
      `)),h=(e,t,n,r="[disabled]")=>a.a`
    .control {
      background: transparent !important;
      border-color: ${d.rb};
    }

    :host(${n}:hover) .control {
      border-color: ${d.nb};
    }

    :host(${n}:active) .control {
      border-color: ${d.ib};
    }

    :host(${r}) .control {
      background: transparent !important;
      border-color: ${d.rb};
    }
  `.withBehaviors(Object(s.a)(a.a`
        .control {
          border-color: ${i.a.ButtonText};
          color: ${i.a.ButtonText};
        }

        :host(${n}:hover) .control,
        :host(${n}:active) .control {
          background: ${i.a.HighlightText};
          border-color: ${i.a.Highlight};
          color: ${i.a.Highlight};
        }

        :host(${r}) .control {
          border-color: ${i.a.GrayText};
          color: ${i.a.GrayText};
        }

        .control:${o.a} {
          outline-color: ${i.a.CanvasText};
        }

        :host([href]) .control {
          border-color: ${i.a.LinkText};
          color: ${i.a.LinkText};
        }

        :host([href]:hover) .control,
        :host([href]:active) .control {
          border-color: ${i.a.CanvasText};
          color: ${i.a.CanvasText};
        }
      `)),b=(e,t,n,r="[disabled]")=>a.a`
    .control {
      background: ${d.ab};
    }

    :host(${n}:hover) .control {
      background: ${d.Y};
    }

    :host(${n}:active) .control {
      background: ${d.W};
    }

    :host(${r}) .control {
      background: ${d.ab};
    }
  `.withBehaviors(Object(s.a)(a.a`
        .control {
          forced-color-adjust: none;
          background: transparent;
          color: ${i.a.ButtonText};
        }

        :host(${n}:hover) .control,
        :host(${n}:active) .control {
          background: transparent;
          border-color: ${i.a.ButtonText};
          color: ${i.a.ButtonText};
        }

        :host(${r}) .control {
          background: transparent;
          color: ${i.a.GrayText};
        }
        
        .control:${o.a} {
          outline-color: ${i.a.CanvasText};
        }

        :host([href]) .control {
          color: ${i.a.LinkText};
        }

        :host([href]:hover) .control,
        :host([href]:active) .control {
          background: transparent;
          border-color: ${i.a.LinkText};
          color: ${i.a.LinkText};
        }
      `))},"8GQ4":function(e,t,n){"use strict";n.d(t,"a",function(){return O});var a=n("eftJ"),i=n("olMv"),r=n("oePG"),o=n("P4Ao"),s=n("+Cud"),c=n("5ZAu"),d=n("+yEz");const l=document.createElement("div");class u{setProperty(e,t){c.a.queueUpdate(()=>this.target.setProperty(e,t))}removeProperty(e){c.a.queueUpdate(()=>this.target.removeProperty(e))}}class f extends u{constructor(){super();const e=new CSSStyleSheet;this.target=e.cssRules[e.insertRule(":root{}")].style,document.adoptedStyleSheets=[...document.adoptedStyleSheets,e]}}class p extends u{constructor(){super(),this.style=document.createElement("style"),document.head.appendChild(this.style);const{sheet:e}=this.style;if(e){const t=e.insertRule(":root{}",e.cssRules.length);this.target=e.cssRules[t].style}}}class m{constructor(e){this.store=new Map,this.target=null;const t=e.$fastController;this.style=document.createElement("style"),t.addStyles(this.style),r.b.getNotifier(t).subscribe(this,"isConnected"),this.handleChange(t,"isConnected")}targetChanged(){if(null!==this.target)for(const[e,t]of this.store.entries())this.target.setProperty(e,t)}setProperty(e,t){this.store.set(e,t),c.a.queueUpdate(()=>{null!==this.target&&this.target.setProperty(e,t)})}removeProperty(e){this.store.delete(e),c.a.queueUpdate(()=>{null!==this.target&&this.target.removeProperty(e)})}handleChange(e,t){const{sheet:n}=this.style;if(n){const e=n.insertRule(":host{}",n.cssRules.length);this.target=n.cssRules[e].style}else this.target=null}}Object(a.a)([r.d],m.prototype,"target",void 0);class _{constructor(e){this.target=e.style}setProperty(e,t){c.a.queueUpdate(()=>this.target.setProperty(e,t))}removeProperty(e){c.a.queueUpdate(()=>this.target.removeProperty(e))}}class h{setProperty(e,t){h.properties[e]=t;for(const n of h.roots.values())v.getOrCreate(h.normalizeRoot(n)).setProperty(e,t)}removeProperty(e){delete h.properties[e];for(const t of h.roots.values())v.getOrCreate(h.normalizeRoot(t)).removeProperty(e)}static registerRoot(e){const{roots:t}=h;if(!t.has(e)){t.add(e);const n=v.getOrCreate(this.normalizeRoot(e));for(const e in h.properties)n.setProperty(e,h.properties[e])}}static unregisterRoot(e){const{roots:t}=h;if(t.has(e)){t.delete(e);const n=v.getOrCreate(h.normalizeRoot(e));for(const e in h.properties)n.removeProperty(e)}}static normalizeRoot(e){return e===l?document:e}}h.roots=new Set,h.properties={};const b=new WeakMap,g=c.a.supportsAdoptedStyleSheets?class extends u{constructor(e){super();const t=new CSSStyleSheet;this.target=t.cssRules[t.insertRule(":host{}")].style,e.$fastController.addStyles(d.a.create([t]))}}:m,v=Object.freeze({getOrCreate(e){if(b.has(e))return b.get(e);let t;return t=e===l?new h:e instanceof Document?c.a.supportsAdoptedStyleSheets?new f:new p:e instanceof o.a?new g(e):new _(e),b.set(e,t),t}});class y extends i.a{constructor(e){super(),this.subscribers=new WeakMap,this._appliedTo=new Set,this.name=e.name,null!==e.cssCustomPropertyName&&(this.cssCustomProperty=`--${e.cssCustomPropertyName}`,this.cssVar=`var(${this.cssCustomProperty})`),this.id=y.uniqueId(),y.tokensById.set(this.id,this)}get appliedTo(){return[...this._appliedTo]}static from(e){return new y({name:"string"==typeof e?e:e.name,cssCustomPropertyName:"string"==typeof e?e:void 0===e.cssCustomPropertyName?e.name:e.cssCustomPropertyName})}static isCSSDesignToken(e){return"string"==typeof e.cssCustomProperty}static isDerivedDesignTokenValue(e){return"function"==typeof e}static getTokenById(e){return y.tokensById.get(e)}getOrCreateSubscriberSet(e=this){return this.subscribers.get(e)||this.subscribers.set(e,new Set)&&this.subscribers.get(e)}createCSS(){return this.cssVar||""}getValueFor(e){const t=C.getOrCreate(e).get(this);if(void 0!==t)return t;throw new Error(`Value could not be retrieved for token named "${this.name}". Ensure the value is set for ${e} or an ancestor of ${e}.`)}setValueFor(e,t){return this._appliedTo.add(e),t instanceof y&&(t=this.alias(t)),C.getOrCreate(e).set(this,t),this}deleteValueFor(e){return this._appliedTo.delete(e),C.existsFor(e)&&C.getOrCreate(e).delete(this),this}withDefault(e){return this.setValueFor(l,e),this}subscribe(e,t){const n=this.getOrCreateSubscriberSet(t);t&&!C.existsFor(t)&&C.getOrCreate(t),n.has(e)||n.add(e)}unsubscribe(e,t){const n=this.subscribers.get(t||this);n&&n.has(e)&&n.delete(e)}notify(e){const t=Object.freeze({token:this,target:e});this.subscribers.has(this)&&this.subscribers.get(this).forEach(e=>e.handleChange(t)),this.subscribers.has(e)&&this.subscribers.get(e).forEach(e=>e.handleChange(t))}alias(e){return t=>e.getValueFor(t)}}y.uniqueId=(()=>{let e=0;return()=>(e++,e.toString(16))})(),y.tokensById=new Map;class S{constructor(e,t,n){this.source=e,this.token=t,this.node=n,this.dependencies=new Set,this.observer=r.b.binding(e,this,!1),this.observer.handleChange=this.observer.call,this.handleChange()}disconnect(){this.observer.disconnect()}handleChange(){this.node.store.set(this.token,this.observer.observe(this.node.target,r.c))}}class D{constructor(){this.values=new Map}set(e,t){this.values.get(e)!==t&&(this.values.set(e,t),r.b.getNotifier(this).notify(e.id))}get(e){return r.b.track(this,e.id),this.values.get(e)}delete(e){this.values.delete(e)}all(){return this.values.entries()}}const I=new WeakMap,x=new WeakMap;class C{constructor(e){this.target=e,this.store=new D,this.children=[],this.assignedValues=new Map,this.reflecting=new Set,this.bindingObservers=new Map,this.tokenValueChangeHandler={handleChange:(e,t)=>{const n=y.getTokenById(t);if(n&&(n.notify(this.target),y.isCSSDesignToken(n))){const t=this.parent,a=this.isReflecting(n);if(t){const i=t.get(n),r=e.get(n);i===r||a?i===r&&a&&this.stopReflectToCSS(n):this.reflectToCSS(n)}else a||this.reflectToCSS(n)}}},I.set(e,this),r.b.getNotifier(this.store).subscribe(this.tokenValueChangeHandler),e instanceof o.a?e.$fastController.addBehaviors([this]):e.isConnected&&this.bind()}static getOrCreate(e){return I.get(e)||new C(e)}static existsFor(e){return I.has(e)}static findParent(e){if(l!==e.target){let t=Object(s.a)(e.target);for(;null!==t;){if(I.has(t))return I.get(t);t=Object(s.a)(t)}return C.getOrCreate(l)}return null}static findClosestAssignedNode(e,t){let n=t;do{if(n.has(e))return n;n=n.parent?n.parent:n.target!==l?C.getOrCreate(l):null}while(null!==n);return null}get parent(){return x.get(this)||null}has(e){return this.assignedValues.has(e)}get(e){const t=this.store.get(e);if(void 0!==t)return t;const n=this.getRaw(e);return void 0!==n?(this.hydrate(e,n),this.get(e)):void 0}getRaw(e){var t;return this.assignedValues.has(e)?this.assignedValues.get(e):null===(t=C.findClosestAssignedNode(e,this))||void 0===t?void 0:t.getRaw(e)}set(e,t){y.isDerivedDesignTokenValue(this.assignedValues.get(e))&&this.tearDownBindingObserver(e),this.assignedValues.set(e,t),y.isDerivedDesignTokenValue(t)?this.setupBindingObserver(e,t):this.store.set(e,t)}delete(e){this.assignedValues.delete(e),this.tearDownBindingObserver(e);const t=this.getRaw(e);t?this.hydrate(e,t):this.store.delete(e)}bind(){const e=C.findParent(this);e&&e.appendChild(this);for(const e of this.assignedValues.keys())e.notify(this.target)}unbind(){this.parent&&x.get(this).removeChild(this)}appendChild(e){e.parent&&x.get(e).removeChild(e);const t=this.children.filter(t=>e.contains(t));x.set(e,this),this.children.push(e),t.forEach(t=>e.appendChild(t)),r.b.getNotifier(this.store).subscribe(e);for(const[t,n]of this.store.all())e.hydrate(t,this.bindingObservers.has(t)?this.getRaw(t):n)}removeChild(e){const t=this.children.indexOf(e);return-1!==t&&this.children.splice(t,1),r.b.getNotifier(this.store).unsubscribe(e),e.parent===this&&x.delete(e)}contains(e){return function(e,t){let n=t;for(;null!==n;){if(n===e)return!0;n=Object(s.a)(n)}return!1}(this.target,e.target)}reflectToCSS(e){this.isReflecting(e)||(this.reflecting.add(e),C.cssCustomPropertyReflector.startReflection(e,this.target))}stopReflectToCSS(e){this.isReflecting(e)&&(this.reflecting.delete(e),C.cssCustomPropertyReflector.stopReflection(e,this.target))}isReflecting(e){return this.reflecting.has(e)}handleChange(e,t){const n=y.getTokenById(t);n&&this.hydrate(n,this.getRaw(n))}hydrate(e,t){if(!this.has(e)){const n=this.bindingObservers.get(e);y.isDerivedDesignTokenValue(t)?n?n.source!==t&&(this.tearDownBindingObserver(e),this.setupBindingObserver(e,t)):this.setupBindingObserver(e,t):(n&&this.tearDownBindingObserver(e),this.store.set(e,t))}}setupBindingObserver(e,t){const n=new S(t,e,this);return this.bindingObservers.set(e,n),n}tearDownBindingObserver(e){return!!this.bindingObservers.has(e)&&(this.bindingObservers.get(e).disconnect(),this.bindingObservers.delete(e),!0)}}C.cssCustomPropertyReflector=new class{startReflection(e,t){e.subscribe(this,t),this.handleChange({token:e,target:t})}stopReflection(e,t){e.unsubscribe(this,t),this.remove(e,t)}handleChange(e){const{token:t,target:n}=e;this.add(t,n)}add(e,t){v.getOrCreate(t).setProperty(e.cssCustomProperty,this.resolveCSSValue(C.getOrCreate(t).get(e)))}remove(e,t){v.getOrCreate(t).removeProperty(e.cssCustomProperty)}resolveCSSValue(e){return e&&"function"==typeof e.createCSS?e.createCSS():e}},Object(a.a)([r.d],C.prototype,"children",void 0);const O=Object.freeze({create:function(e){return y.from(e)},notifyConnection:e=>!(!e.isConnected||!C.existsFor(e)||(C.getOrCreate(e).bind(),0)),notifyDisconnection:e=>!(e.isConnected||!C.existsFor(e)||(C.getOrCreate(e).unbind(),0)),registerRoot(e=l){h.registerRoot(e)},unregisterRoot(e=l){h.unregisterRoot(e)}})},"8hiW":function(e,t,n){"use strict";n.d(t,"t",function(){return M}),n.d(t,"u",function(){return P}),n.d(t,"n",function(){return T}),n.d(t,"r",function(){return U}),n.d(t,"s",function(){return F}),n.d(t,"q",function(){return H}),n.d(t,"D",function(){return R}),n.d(t,"vb",function(){return N}),n.d(t,"y",function(){return B}),n.d(t,"p",function(){return j}),n.d(t,"wb",function(){return G}),n.d(t,"yb",function(){return K}),n.d(t,"xb",function(){return W}),n.d(t,"zb",function(){return q}),n.d(t,"Bb",function(){return Q}),n.d(t,"Ab",function(){return Y}),n.d(t,"Cb",function(){return J}),n.d(t,"Eb",function(){return X}),n.d(t,"Db",function(){return Z}),n.d(t,"Fb",function(){return $}),n.d(t,"Hb",function(){return ee}),n.d(t,"Gb",function(){return te}),n.d(t,"Ib",function(){return ne}),n.d(t,"Kb",function(){return ae}),n.d(t,"Jb",function(){return ie}),n.d(t,"Lb",function(){return re}),n.d(t,"Nb",function(){return oe}),n.d(t,"Mb",function(){return se}),n.d(t,"Ob",function(){return ce}),n.d(t,"Qb",function(){return de}),n.d(t,"Pb",function(){return le}),n.d(t,"Rb",function(){return ue}),n.d(t,"Tb",function(){return fe}),n.d(t,"Sb",function(){return pe}),n.d(t,"Ub",function(){return me}),n.d(t,"Wb",function(){return _e}),n.d(t,"Vb",function(){return he}),n.d(t,"o",function(){return be}),n.d(t,"E",function(){return _t}),n.d(t,"hb",function(){return ht}),n.d(t,"a",function(){return bt}),n.d(t,"gb",function(){return St}),n.d(t,"v",function(){return wt}),n.d(t,"e",function(){return Lt}),n.d(t,"d",function(){return kt}),n.d(t,"b",function(){return Mt}),n.d(t,"c",function(){return Pt}),n.d(t,"C",function(){return Ut}),n.d(t,"B",function(){return Ft}),n.d(t,"z",function(){return Ht}),n.d(t,"A",function(){return Rt}),n.d(t,"i",function(){return Bt}),n.d(t,"h",function(){return jt}),n.d(t,"f",function(){return Vt}),n.d(t,"g",function(){return zt}),n.d(t,"m",function(){return Kt}),n.d(t,"l",function(){return Wt}),n.d(t,"j",function(){return qt}),n.d(t,"k",function(){return Qt}),n.d(t,"Q",function(){return Jt}),n.d(t,"G",function(){return Xt}),n.d(t,"F",function(){return Zt}),n.d(t,"N",function(){return $t}),n.d(t,"O",function(){return en}),n.d(t,"M",function(){return tn}),n.d(t,"L",function(){return nn}),n.d(t,"K",function(){return rn}),n.d(t,"J",function(){return on}),n.d(t,"H",function(){return sn}),n.d(t,"I",function(){return cn}),n.d(t,"P",function(){return dn}),n.d(t,"U",function(){return un}),n.d(t,"V",function(){return fn}),n.d(t,"T",function(){return pn}),n.d(t,"R",function(){return mn}),n.d(t,"S",function(){return _n}),n.d(t,"Z",function(){return hn}),n.d(t,"ab",function(){return bn}),n.d(t,"Y",function(){return gn}),n.d(t,"W",function(){return vn}),n.d(t,"X",function(){return yn}),n.d(t,"fb",function(){return In}),n.d(t,"eb",function(){return xn}),n.d(t,"bb",function(){return Cn}),n.d(t,"db",function(){return On}),n.d(t,"cb",function(){return wn}),n.d(t,"rb",function(){return An}),n.d(t,"nb",function(){return Ln}),n.d(t,"ib",function(){return kn}),n.d(t,"lb",function(){return Pn}),n.d(t,"kb",function(){return Tn}),n.d(t,"jb",function(){return Un}),n.d(t,"mb",function(){return Hn}),n.d(t,"pb",function(){return Nn}),n.d(t,"ob",function(){return Bn}),n.d(t,"qb",function(){return Vn}),n.d(t,"ub",function(){return Gn}),n.d(t,"tb",function(){return Kn}),n.d(t,"sb",function(){return Wn}),n.d(t,"x",function(){return Qn}),n.d(t,"w",function(){return Jn});var a=n("8GQ4"),i=n("qibw"),r=n("hv+n"),o=n("YBl9"),s=n("i2HT");const c=s.a.create(1,1,1),d=s.a.create(0,0,0),l=s.a.create(.5,.5,.5),u=Object(o.a)("#0078D4"),f=s.a.create(u.r,u.g,u.b);function p(e,t,n,a,i){const r=e=>e.contrast(c)>=i?c:d,o=r(e),s=r(t);return{rest:o,hover:s,active:o.relativeLuminance===s.relativeLuminance?o:r(n),focus:r(a)}}var m,_=n("6eT7"),h=n("swXE");n("PT2o"),n("RN7+"),n("SiT+"),n("0q6d"),function(e){e[e.Burn=0]="Burn",e[e.Color=1]="Color",e[e.Darken=2]="Darken",e[e.Dodge=3]="Dodge",e[e.Lighten=4]="Lighten",e[e.Multiply=5]="Multiply",e[e.Overlay=6]="Overlay",e[e.Screen=7]="Screen"}(m||(m={}));var b=n("s2f2"),g=n("THNU");class v{constructor(e,t,n,a){this.toColorString=()=>this.cssGradient,this.contrast=g.a.bind(null,this),this.createCSS=this.toColorString,this.color=new _.a(e,t,n),this.cssGradient=a,this.relativeLuminance=Object(h.j)(this.color),this.r=e,this.g=t,this.b=n}static fromObject(e,t){return new v(e.r,e.g,e.b,t)}}const y=new _.a(0,0,0),S=new _.a(1,1,1);function D(e,t,n,a,i,r,c,d,l=10,u=!1){const f=e.closestIndexOf(t);function p(n){if(u){const a=e.closestIndexOf(t),i=e.get(a),r=n.relativeLuminance<t.relativeLuminance?y:S,c=Object(h.a)(Object(o.a)(n.toColorString()),Object(o.a)(i.toColorString()),r).roundToPrecision(2),d=function(e,t){if(t.a>=1)return t;if(t.a<=0)return new _.a(e.r,e.g,e.b,1);const n=t.a*t.r+(1-t.a)*e.r,a=t.a*t.g+(1-t.a)*e.g,i=t.a*t.b+(1-t.a)*e.b;return new _.a(n,a,i,1)}(Object(o.a)(t.toColorString()),c);return s.a.from(d)}return n}void 0===d&&(d=Object(b.a)(t));const m=f+d*n,g=m+d*(a-n),D=m+d*(i-n),I=m+d*(r-n),x=-1===d?0:100-l,C=-1===d?l:100;function O(t,n){const a=e.get(t);if(n){const n=e.get(t+d*c),i=-1===d?n:a,r=-1===d?a:n,o=`linear-gradient(${p(i).toColorString()} ${x}%, ${p(r).toColorString()} ${C}%)`;return v.fromObject(i,o)}return p(a)}return{rest:O(m,!0),hover:O(g,!0),active:O(D,!1),focus:O(I,!0)}}var I=n("Kct4");function x(e,t,n,a,i,r,o,s){null==s&&(s=Object(b.a)(t));const c=e.closestIndexOf(e.colorContrast(t,n));return{rest:e.get(c+s*a),hover:e.get(c+s*i),active:e.get(c+s*r),focus:e.get(c+s*o)}}function C(e,t,n,a,i,r,o){const s=e.closestIndexOf(t);return null==o&&(o=Object(b.a)(t)),{rest:e.get(s+o*n),hover:e.get(s+o*a),active:e.get(s+o*i),focus:e.get(s+o*r)}}function O(e,t,n,a,i,r,o=void 0,s,c,d,l,u=void 0){return Object(I.a)(t)?C(e,t,s,c,d,l,u):C(e,t,n,a,i,r,o)}var w=n("yQ0v");function E(e,t){return e.closestIndexOf(Object(w.b)(t))}function A(e,t,n){return e.get(E(e,t)+-1*n)}const{create:L}=a.a;function k(e){return a.a.create({name:e,cssCustomPropertyName:null})}const M=L("direction").withDefault(i.a.ltr),P=L("disabled-opacity").withDefault(.3),T=L("base-height-multiplier").withDefault(8),U=(L("base-horizontal-spacing-multiplier").withDefault(3),L("density").withDefault(0)),F=L("design-unit").withDefault(4),H=L("control-corner-radius").withDefault(4),R=L("layer-corner-radius").withDefault(8),N=L("stroke-width").withDefault(1),B=L("focus-stroke-width").withDefault(2),j=L("body-font").withDefault('"Segoe UI Variable", "Segoe UI", sans-serif'),V=L("font-weight").withDefault(400);function z(e){return t=>{const n=e.getValueFor(t),a=V.getValueFor(t);if(n.endsWith("px")){const e=Number.parseFloat(n.replace("px",""));if(e<=12)return`"wght" ${a}, "opsz" 8`;if(e>24)return`"wght" ${a}, "opsz" 36`}return`"wght" ${a}, "opsz" 10.5`}}const G=L("type-ramp-base-font-size").withDefault("14px"),K=L("type-ramp-base-line-height").withDefault("20px"),W=L("type-ramp-base-font-variations").withDefault(z(G)),q=L("type-ramp-minus-1-font-size").withDefault("12px"),Q=L("type-ramp-minus-1-line-height").withDefault("16px"),Y=L("type-ramp-minus-1-font-variations").withDefault(z(q)),J=L("type-ramp-minus-2-font-size").withDefault("10px"),X=L("type-ramp-minus-2-line-height").withDefault("14px"),Z=L("type-ramp-minus-2-font-variations").withDefault(z(J)),$=L("type-ramp-plus-1-font-size").withDefault("16px"),ee=L("type-ramp-plus-1-line-height").withDefault("22px"),te=L("type-ramp-plus-1-font-variations").withDefault(z($)),ne=L("type-ramp-plus-2-font-size").withDefault("20px"),ae=L("type-ramp-plus-2-line-height").withDefault("26px"),ie=L("type-ramp-plus-2-font-variations").withDefault(z(ne)),re=L("type-ramp-plus-3-font-size").withDefault("24px"),oe=L("type-ramp-plus-3-line-height").withDefault("32px"),se=L("type-ramp-plus-3-font-variations").withDefault(z(re)),ce=L("type-ramp-plus-4-font-size").withDefault("28px"),de=L("type-ramp-plus-4-line-height").withDefault("36px"),le=L("type-ramp-plus-4-font-variations").withDefault(z(ce)),ue=L("type-ramp-plus-5-font-size").withDefault("32px"),fe=L("type-ramp-plus-5-line-height").withDefault("40px"),pe=L("type-ramp-plus-5-font-variations").withDefault(z(ue)),me=L("type-ramp-plus-6-font-size").withDefault("40px"),_e=L("type-ramp-plus-6-line-height").withDefault("52px"),he=L("type-ramp-plus-6-font-variations").withDefault(z(me)),be=L("base-layer-luminance").withDefault(w.a.LightMode),ge=k("accent-fill-rest-delta").withDefault(0),ve=k("accent-fill-hover-delta").withDefault(-2),ye=k("accent-fill-active-delta").withDefault(-5),Se=k("accent-fill-focus-delta").withDefault(0),De=k("accent-foreground-rest-delta").withDefault(0),Ie=k("accent-foreground-hover-delta").withDefault(3),xe=k("accent-foreground-active-delta").withDefault(-8),Ce=k("accent-foreground-focus-delta").withDefault(0),Oe=k("neutral-fill-rest-delta").withDefault(-1),we=k("neutral-fill-hover-delta").withDefault(1),Ee=k("neutral-fill-active-delta").withDefault(0),Ae=k("neutral-fill-focus-delta").withDefault(0),Le=k("neutral-fill-input-rest-delta").withDefault(-1),ke=k("neutral-fill-input-hover-delta").withDefault(1),Me=k("neutral-fill-input-active-delta").withDefault(0),Pe=k("neutral-fill-input-focus-delta").withDefault(-2),Te=k("neutral-fill-input-alt-rest-delta").withDefault(2),Ue=k("neutral-fill-input-alt-hover-delta").withDefault(4),Fe=k("neutral-fill-input-alt-active-delta").withDefault(6),He=k("neutral-fill-input-alt-focus-delta").withDefault(2),Re=k("neutral-fill-layer-rest-delta").withDefault(-2),Ne=k("neutral-fill-layer-hover-delta").withDefault(-3),Be=k("neutral-fill-layer-active-delta").withDefault(-3),je=k("neutral-fill-layer-alt-rest-delta").withDefault(-1),Ve=k("neutral-fill-secondary-rest-delta").withDefault(3),ze=k("neutral-fill-secondary-hover-delta").withDefault(2),Ge=k("neutral-fill-secondary-active-delta").withDefault(1),Ke=k("neutral-fill-secondary-focus-delta").withDefault(3),We=k("neutral-fill-stealth-rest-delta").withDefault(0),qe=k("neutral-fill-stealth-hover-delta").withDefault(3),Qe=k("neutral-fill-stealth-active-delta").withDefault(2),Ye=k("neutral-fill-stealth-focus-delta").withDefault(0),Je=k("neutral-fill-strong-rest-delta").withDefault(0),Xe=k("neutral-fill-strong-hover-delta").withDefault(8),Ze=k("neutral-fill-strong-active-delta").withDefault(-5),$e=k("neutral-fill-strong-focus-delta").withDefault(0),et=k("neutral-stroke-rest-delta").withDefault(8),tt=k("neutral-stroke-hover-delta").withDefault(12),nt=k("neutral-stroke-active-delta").withDefault(6),at=k("neutral-stroke-focus-delta").withDefault(8),it=k("neutral-stroke-control-rest-delta").withDefault(3),rt=k("neutral-stroke-control-hover-delta").withDefault(5),ot=k("neutral-stroke-control-active-delta").withDefault(5),st=k("neutral-stroke-control-focus-delta").withDefault(5),ct=k("neutral-stroke-divider-rest-delta").withDefault(4),dt=k("neutral-stroke-layer-rest-delta").withDefault(3),lt=k("neutral-stroke-layer-hover-delta").withDefault(3),ut=k("neutral-stroke-layer-active-delta").withDefault(3),ft=k("neutral-stroke-strong-hover-delta").withDefault(0),pt=k("neutral-stroke-strong-active-delta").withDefault(0),mt=k("neutral-stroke-strong-focus-delta").withDefault(0),_t=L("neutral-base-color").withDefault(l),ht=k("neutral-palette").withDefault(e=>r.a.from(_t.getValueFor(e))),bt=L("accent-base-color").withDefault(f),gt=k("accent-palette").withDefault(e=>r.a.from(bt.getValueFor(e))),vt=k("neutral-layer-card-container-recipe").withDefault({evaluate:e=>A(ht.getValueFor(e),be.getValueFor(e),Re.getValueFor(e))}),yt=(L("neutral-layer-card-container").withDefault(e=>vt.getValueFor(e).evaluate(e)),k("neutral-layer-floating-recipe").withDefault({evaluate:e=>{return t=ht.getValueFor(e),n=be.getValueFor(e),a=Re.getValueFor(e),t.get(E(t,n)+a);var t,n,a}})),St=L("neutral-layer-floating").withDefault(e=>yt.getValueFor(e).evaluate(e)),Dt=k("neutral-layer-1-recipe").withDefault({evaluate:e=>{return t=ht.getValueFor(e),n=be.getValueFor(e),t.get(E(t,n));var t,n}}),It=L("neutral-layer-1").withDefault(e=>Dt.getValueFor(e).evaluate(e)),xt=k("neutral-layer-2-recipe").withDefault({evaluate:e=>A(ht.getValueFor(e),be.getValueFor(e),Re.getValueFor(e))}),Ct=(L("neutral-layer-2").withDefault(e=>xt.getValueFor(e).evaluate(e)),k("neutral-layer-3-recipe").withDefault({evaluate:e=>{return t=ht.getValueFor(e),n=be.getValueFor(e),a=Re.getValueFor(e),t.get(E(t,n)+-1*a*2);var t,n,a}})),Ot=(L("neutral-layer-3").withDefault(e=>Ct.getValueFor(e).evaluate(e)),k("neutral-layer-4-recipe").withDefault({evaluate:e=>{return t=ht.getValueFor(e),n=be.getValueFor(e),a=Re.getValueFor(e),t.get(E(t,n)+-1*a*3);var t,n,a}})),wt=(L("neutral-layer-4").withDefault(e=>Ot.getValueFor(e).evaluate(e)),L("fill-color").withDefault(e=>It.getValueFor(e)));var Et;!function(e){e[e.normal=4.5]="normal",e[e.large=3]="large"}(Et||(Et={}));const At=k("accent-fill-recipe").withDefault({evaluate:(e,t)=>function(e,t,n,a,i,r,o,s,c,d,l,u,f,p){return Object(I.a)(t)?x(e,t,8,d,l,u,f,void 0):x(e,t,5,a,i,r,o,void 0)}(gt.getValueFor(e),t||wt.getValueFor(e),0,ge.getValueFor(e),ve.getValueFor(e),ye.getValueFor(e),Se.getValueFor(e),0,0,ge.getValueFor(e),ve.getValueFor(e),ye.getValueFor(e),Se.getValueFor(e))}),Lt=L("accent-fill-rest").withDefault(e=>At.getValueFor(e).evaluate(e).rest),kt=L("accent-fill-hover").withDefault(e=>At.getValueFor(e).evaluate(e).hover),Mt=L("accent-fill-active").withDefault(e=>At.getValueFor(e).evaluate(e).active),Pt=L("accent-fill-focus").withDefault(e=>At.getValueFor(e).evaluate(e).focus),Tt=k("foreground-on-accent-recipe").withDefault({evaluate:e=>p(Lt.getValueFor(e),kt.getValueFor(e),Mt.getValueFor(e),Pt.getValueFor(e),Et.normal)}),Ut=L("foreground-on-accent-rest").withDefault(e=>Tt.getValueFor(e).evaluate(e).rest),Ft=L("foreground-on-accent-hover").withDefault(e=>Tt.getValueFor(e).evaluate(e).hover),Ht=L("foreground-on-accent-active").withDefault(e=>Tt.getValueFor(e).evaluate(e).active),Rt=L("foreground-on-accent-focus").withDefault(e=>Tt.getValueFor(e).evaluate(e).focus),Nt=k("accent-foreground-recipe").withDefault({evaluate:(e,t)=>x(gt.getValueFor(e),t||wt.getValueFor(e),9.5,De.getValueFor(e),Ie.getValueFor(e),xe.getValueFor(e),Ce.getValueFor(e))}),Bt=L("accent-foreground-rest").withDefault(e=>Nt.getValueFor(e).evaluate(e).rest),jt=L("accent-foreground-hover").withDefault(e=>Nt.getValueFor(e).evaluate(e).hover),Vt=L("accent-foreground-active").withDefault(e=>Nt.getValueFor(e).evaluate(e).active),zt=L("accent-foreground-focus").withDefault(e=>Nt.getValueFor(e).evaluate(e).focus),Gt=k("accent-stroke-control-recipe").withDefault({evaluate:(e,t)=>D(ht.getValueFor(e),t||wt.getValueFor(e),-3,-3,-3,-3,10,1,void 0,!0)}),Kt=L("accent-stroke-control-rest").withDefault(e=>Gt.getValueFor(e).evaluate(e,Lt.getValueFor(e)).rest),Wt=L("accent-stroke-control-hover").withDefault(e=>Gt.getValueFor(e).evaluate(e,kt.getValueFor(e)).hover),qt=L("accent-stroke-control-active").withDefault(e=>Gt.getValueFor(e).evaluate(e,Mt.getValueFor(e)).active),Qt=L("accent-stroke-control-focus").withDefault(e=>Gt.getValueFor(e).evaluate(e,Pt.getValueFor(e)).focus),Yt=k("neutral-fill-recipe").withDefault({evaluate:(e,t)=>O(ht.getValueFor(e),t||wt.getValueFor(e),Oe.getValueFor(e),we.getValueFor(e),Ee.getValueFor(e),Ae.getValueFor(e),void 0,2,3,1,2,void 0)}),Jt=L("neutral-fill-rest").withDefault(e=>Yt.getValueFor(e).evaluate(e).rest),Xt=L("neutral-fill-hover").withDefault(e=>Yt.getValueFor(e).evaluate(e).hover),Zt=L("neutral-fill-active").withDefault(e=>Yt.getValueFor(e).evaluate(e).active),$t=(L("neutral-fill-focus").withDefault(e=>Yt.getValueFor(e).evaluate(e).focus),k("neutral-fill-input-recipe").withDefault({evaluate:(e,t)=>O(ht.getValueFor(e),t||wt.getValueFor(e),Le.getValueFor(e),ke.getValueFor(e),Me.getValueFor(e),Pe.getValueFor(e),void 0,2,3,1,0,void 0)})),en=L("neutral-fill-input-rest").withDefault(e=>$t.getValueFor(e).evaluate(e).rest),tn=L("neutral-fill-input-hover").withDefault(e=>$t.getValueFor(e).evaluate(e).hover),nn=(L("neutral-fill-input-active").withDefault(e=>$t.getValueFor(e).evaluate(e).active),L("neutral-fill-input-focus").withDefault(e=>$t.getValueFor(e).evaluate(e).focus)),an=k("neutral-fill-input-alt-recipe").withDefault({evaluate:(e,t)=>O(ht.getValueFor(e),t||wt.getValueFor(e),Te.getValueFor(e),Ue.getValueFor(e),Fe.getValueFor(e),He.getValueFor(e),1,Te.getValueFor(e),Te.getValueFor(e)-Ue.getValueFor(e),Te.getValueFor(e)-Fe.getValueFor(e),He.getValueFor(e),1)}),rn=L("neutral-fill-input-alt-rest").withDefault(e=>an.getValueFor(e).evaluate(e).rest),on=L("neutral-fill-input-alt-hover").withDefault(e=>an.getValueFor(e).evaluate(e).hover),sn=L("neutral-fill-input-alt-active").withDefault(e=>an.getValueFor(e).evaluate(e).active),cn=L("neutral-fill-input-alt-focus").withDefault(e=>an.getValueFor(e).evaluate(e).focus),dn=k("neutral-fill-layer-recipe").withDefault({evaluate:(e,t)=>C(ht.getValueFor(e),t||wt.getValueFor(e),Re.getValueFor(e),Ne.getValueFor(e),Be.getValueFor(e),Re.getValueFor(e),1)}),ln=(L("neutral-fill-layer-rest").withDefault(e=>dn.getValueFor(e).evaluate(e).rest),L("neutral-fill-layer-hover").withDefault(e=>dn.getValueFor(e).evaluate(e).hover),L("neutral-fill-layer-active").withDefault(e=>dn.getValueFor(e).evaluate(e).active),k("neutral-fill-layer-alt-recipe").withDefault({evaluate:(e,t)=>C(ht.getValueFor(e),t||wt.getValueFor(e),je.getValueFor(e),je.getValueFor(e),je.getValueFor(e),je.getValueFor(e))})),un=(L("neutral-fill-layer-alt-rest").withDefault(e=>ln.getValueFor(e).evaluate(e).rest),k("neutral-fill-secondary-recipe").withDefault({evaluate:(e,t)=>C(ht.getValueFor(e),t||wt.getValueFor(e),Ve.getValueFor(e),ze.getValueFor(e),Ge.getValueFor(e),Ke.getValueFor(e))})),fn=L("neutral-fill-secondary-rest").withDefault(e=>un.getValueFor(e).evaluate(e).rest),pn=L("neutral-fill-secondary-hover").withDefault(e=>un.getValueFor(e).evaluate(e).hover),mn=L("neutral-fill-secondary-active").withDefault(e=>un.getValueFor(e).evaluate(e).active),_n=L("neutral-fill-secondary-focus").withDefault(e=>un.getValueFor(e).evaluate(e).focus),hn=k("neutral-fill-stealth-recipe").withDefault({evaluate:(e,t)=>C(ht.getValueFor(e),t||wt.getValueFor(e),We.getValueFor(e),qe.getValueFor(e),Qe.getValueFor(e),Ye.getValueFor(e))}),bn=L("neutral-fill-stealth-rest").withDefault(e=>hn.getValueFor(e).evaluate(e).rest),gn=L("neutral-fill-stealth-hover").withDefault(e=>hn.getValueFor(e).evaluate(e).hover),vn=L("neutral-fill-stealth-active").withDefault(e=>hn.getValueFor(e).evaluate(e).active),yn=L("neutral-fill-stealth-focus").withDefault(e=>hn.getValueFor(e).evaluate(e).focus),Sn=k("neutral-fill-strong-recipe").withDefault({evaluate:(e,t)=>x(ht.getValueFor(e),t||wt.getValueFor(e),4.5,Je.getValueFor(e),Xe.getValueFor(e),Ze.getValueFor(e),$e.getValueFor(e))}),Dn=(L("neutral-fill-strong-rest").withDefault(e=>Sn.getValueFor(e).evaluate(e).rest),L("neutral-fill-strong-hover").withDefault(e=>Sn.getValueFor(e).evaluate(e).hover),L("neutral-fill-strong-active").withDefault(e=>Sn.getValueFor(e).evaluate(e).active),L("neutral-fill-strong-focus").withDefault(e=>Sn.getValueFor(e).evaluate(e).focus),k("neutral-foreground-recipe").withDefault({evaluate:(e,t)=>x(ht.getValueFor(e),t||wt.getValueFor(e),16,0,-19,-30,0)})),In=L("neutral-foreground-rest").withDefault(e=>Dn.getValueFor(e).evaluate(e).rest),xn=L("neutral-foreground-hover").withDefault(e=>Dn.getValueFor(e).evaluate(e).hover),Cn=L("neutral-foreground-active").withDefault(e=>Dn.getValueFor(e).evaluate(e).active),On=(L("neutral-foreground-focus").withDefault(e=>Dn.getValueFor(e).evaluate(e).focus),k("neutral-foreground-hint-recipe").withDefault({evaluate:(e,t)=>function(e,t,n){return e.colorContrast(t,4.5)}(ht.getValueFor(e),t||wt.getValueFor(e))})),wn=L("neutral-foreground-hint").withDefault(e=>On.getValueFor(e).evaluate(e)),En=k("neutral-stroke-recipe").withDefault({evaluate:(e,t)=>C(ht.getValueFor(e),t||wt.getValueFor(e),et.getValueFor(e),tt.getValueFor(e),nt.getValueFor(e),at.getValueFor(e))}),An=L("neutral-stroke-rest").withDefault(e=>En.getValueFor(e).evaluate(e).rest),Ln=L("neutral-stroke-hover").withDefault(e=>En.getValueFor(e).evaluate(e).hover),kn=L("neutral-stroke-active").withDefault(e=>En.getValueFor(e).evaluate(e).active),Mn=(L("neutral-stroke-focus").withDefault(e=>En.getValueFor(e).evaluate(e).focus),k("neutral-stroke-control-recipe").withDefault({evaluate:(e,t)=>D(ht.getValueFor(e),t||wt.getValueFor(e),it.getValueFor(e),rt.getValueFor(e),ot.getValueFor(e),st.getValueFor(e),5)})),Pn=L("neutral-stroke-control-rest").withDefault(e=>Mn.getValueFor(e).evaluate(e).rest),Tn=L("neutral-stroke-control-hover").withDefault(e=>Mn.getValueFor(e).evaluate(e).hover),Un=L("neutral-stroke-control-active").withDefault(e=>Mn.getValueFor(e).evaluate(e).active),Fn=(L("neutral-stroke-control-focus").withDefault(e=>Mn.getValueFor(e).evaluate(e).focus),k("neutral-stroke-divider-recipe").withDefault({evaluate:(e,t)=>function(e,t,n){return e.get(e.closestIndexOf(t)+Object(b.a)(t)*n)}(ht.getValueFor(e),t||wt.getValueFor(e),ct.getValueFor(e))})),Hn=L("neutral-stroke-divider-rest").withDefault(e=>Fn.getValueFor(e).evaluate(e)),Rn=k("neutral-stroke-input-recipe").withDefault({evaluate:(e,t)=>function(e,t,n,a,i,r,o,s){const c=e.closestIndexOf(t),d=Object(b.a)(t),l=c+d*n,u=l+d*(a-n),f=l+d*(i-n),p=l+d*(r-n),m=`calc(100% - ${s})`;function _(t,n){const a=e.get(t);if(n){const n=e.get(t+20*d),i=`linear-gradient(${a.toColorString()} ${m}, ${n.toColorString()} ${m}, ${n.toColorString()})`;return v.fromObject(a,i)}return a}return{rest:_(l,!0),hover:_(u,!0),active:_(f,!1),focus:_(p,!0)}}(ht.getValueFor(e),t||wt.getValueFor(e),it.getValueFor(e),rt.getValueFor(e),ot.getValueFor(e),st.getValueFor(e),0,N.getValueFor(e)+"px")}),Nn=L("neutral-stroke-input-rest").withDefault(e=>Rn.getValueFor(e).evaluate(e).rest),Bn=L("neutral-stroke-input-hover").withDefault(e=>Rn.getValueFor(e).evaluate(e).hover),jn=(L("neutral-stroke-input-active").withDefault(e=>Rn.getValueFor(e).evaluate(e).active),L("neutral-stroke-input-focus").withDefault(e=>Rn.getValueFor(e).evaluate(e).focus),k("neutral-stroke-layer-recipe").withDefault({evaluate:(e,t)=>C(ht.getValueFor(e),t||wt.getValueFor(e),dt.getValueFor(e),lt.getValueFor(e),ut.getValueFor(e),dt.getValueFor(e))})),Vn=L("neutral-stroke-layer-rest").withDefault(e=>jn.getValueFor(e).evaluate(e).rest),zn=(L("neutral-stroke-layer-hover").withDefault(e=>jn.getValueFor(e).evaluate(e).hover),L("neutral-stroke-layer-active").withDefault(e=>jn.getValueFor(e).evaluate(e).active),k("neutral-stroke-strong-recipe").withDefault({evaluate:(e,t)=>x(ht.getValueFor(e),t||wt.getValueFor(e),5.5,0,ft.getValueFor(e),pt.getValueFor(e),mt.getValueFor(e))})),Gn=L("neutral-stroke-strong-rest").withDefault(e=>zn.getValueFor(e).evaluate(e).rest),Kn=L("neutral-stroke-strong-hover").withDefault(e=>zn.getValueFor(e).evaluate(e).hover),Wn=L("neutral-stroke-strong-active").withDefault(e=>zn.getValueFor(e).evaluate(e).active),qn=(L("neutral-stroke-strong-focus").withDefault(e=>zn.getValueFor(e).evaluate(e).focus),k("focus-stroke-outer-recipe").withDefault({evaluate:e=>{return ht.getValueFor(e),t=wt.getValueFor(e),Object(I.a)(t)?c:d;var t}})),Qn=L("focus-stroke-outer").withDefault(e=>qn.getValueFor(e).evaluate(e)),Yn=k("focus-stroke-inner-recipe").withDefault({evaluate:e=>{return gt.getValueFor(e),t=wt.getValueFor(e),Qn.getValueFor(e),Object(I.a)(t)?d:c;var t}}),Jn=L("focus-stroke-inner").withDefault(e=>Yn.getValueFor(e).evaluate(e)),Xn=k("foreground-on-accent-large-recipe").withDefault({evaluate:e=>p(Lt.getValueFor(e),kt.getValueFor(e),Mt.getValueFor(e),Pt.getValueFor(e),Et.large)}),Zn=(L("foreground-on-accent-rest-large").withDefault(e=>Xn.getValueFor(e).evaluate(e).rest),L("foreground-on-accent-hover-large").withDefault(e=>Xn.getValueFor(e).evaluate(e,kt.getValueFor(e)).hover),L("foreground-on-accent-active-large").withDefault(e=>Xn.getValueFor(e).evaluate(e,Mt.getValueFor(e)).active),L("foreground-on-accent-focus-large").withDefault(e=>Xn.getValueFor(e).evaluate(e,Pt.getValueFor(e)).focus),L("neutral-fill-inverse-rest-delta").withDefault(0)),$n=L("neutral-fill-inverse-hover-delta").withDefault(-3),ea=L("neutral-fill-inverse-active-delta").withDefault(7),ta=L("neutral-fill-inverse-focus-delta").withDefault(0),na=k("neutral-fill-inverse-recipe").withDefault({evaluate:(e,t)=>function(e,t,n,a,i,r){const o=Object(b.a)(t),s=e.closestIndexOf(e.colorContrast(t,14)),c=s+o*Math.abs(n-a);let d,l;return(1===o?n<a:o*n>o*a)?(d=s,l=c):(d=c,l=s),{rest:e.get(d),hover:e.get(l),active:e.get(d+o*i),focus:e.get(d+o*r)}}(ht.getValueFor(e),t||wt.getValueFor(e),Zn.getValueFor(e),$n.getValueFor(e),ea.getValueFor(e),ta.getValueFor(e))});L("neutral-fill-inverse-rest").withDefault(e=>na.getValueFor(e).evaluate(e).rest),L("neutral-fill-inverse-hover").withDefault(e=>na.getValueFor(e).evaluate(e).hover),L("neutral-fill-inverse-active").withDefault(e=>na.getValueFor(e).evaluate(e).active),L("neutral-fill-inverse-focus").withDefault(e=>na.getValueFor(e).evaluate(e).focus)},C001:function(e,t,n){"use strict";n.d(t,"a",function(){return d});var a=n("Y1A4"),i=n("cBsD"),r=n("HN6m"),o=n("qqMp"),s=n("ZgG/"),c=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};class d extends a.a{get personDetails(){return this._personDetails}set personDetails(e){this._personDetails!==e&&(this._personDetails=e,this.requestStateUpdate())}get isCompact(){return this._isCompact}constructor(){super(),this._isCompact=!1,this._personDetails=null}asCompactView(){return this._isCompact=!0,this.requestUpdate(),this}asFullView(){return this._isCompact=!1,this.requestUpdate(),this}clearState(){this._isCompact=!1,this._personDetails=null}render(){return this.isCompact?this.renderCompactView():this.renderFullView()}renderLoading(){return i.a`
      <div class="loading">
        <mgt-spinner></mgt-spinner>
      </div>
    `}renderNoData(){return o.c`
      <div class="no-data">No data</div>
    `}navigateCard(e){var t;let n=this.parentNode;for(;n;){n=n.parentNode;const e=n;if((null===(t=null==e?void 0:e.host)||void 0===t?void 0:t.tagName)===`${r.a.prefix}-PERSON-CARD`.toUpperCase()){n=e.host;break}}n.navigate(e)}}!function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);r>3&&o&&Object.defineProperty(t,n,o)}([Object(s.b)({attribute:"person-details",type:Object}),c("design:type",Object),c("design:paramtypes",[Object])],d.prototype,"personDetails",null)},C5kU:function(e,t,n){"use strict";n.d(t,"a",function(){return r}),n.d(t,"b",function(){return o}),n.d(t,"d",function(){return s}),n.d(t,"c",function(){return c}),n.d(t,"e",function(){return d});var a=n("6BDD"),i=n("+53S");class r{handleStartContentChange(){this.startContainer.classList.toggle("start",this.start.assignedNodes().length>0)}handleEndContentChange(){this.endContainer.classList.toggle("end",this.end.assignedNodes().length>0)}}const o=(e,t)=>a.a`
    <span
        part="end"
        ${Object(i.a)("endContainer")}
        class=${e=>t.end?"end":void 0}
    >
        <slot name="end" ${Object(i.a)("end")} @slotchange="${e=>e.handleEndContentChange()}">
            ${t.end||""}
        </slot>
    </span>
`,s=(e,t)=>a.a`
    <span
        part="start"
        ${Object(i.a)("startContainer")}
        class="${e=>t.start?"start":void 0}"
    >
        <slot
            name="start"
            ${Object(i.a)("start")}
            @slotchange="${e=>e.handleStartContentChange()}"
        >
            ${t.start||""}
        </slot>
    </span>
`,c=a.a`
    <span part="end" ${Object(i.a)("endContainer")}>
        <slot
            name="end"
            ${Object(i.a)("end")}
            @slotchange="${e=>e.handleEndContentChange()}"
        ></slot>
    </span>
`,d=a.a`
    <span part="start" ${Object(i.a)("startContainer")}>
        <slot
            name="start"
            ${Object(i.a)("start")}
            @slotchange="${e=>e.handleStartContentChange()}"
        ></slot>
    </span>
`},D4ZN:function(e,t,n){"use strict";n.d(t,"b",function(){return s}),n.d(t,"a",function(){return c});var a=n("qqMp"),i=n("dblZ");const r=[a.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}
`];var o=n("0mOt");const s=()=>Object(o.b)("spinner",c);class c extends i.b{static get styles(){return r}render(){return a.c`<fluent-progress-ring title="spinner"></fluent-progress-ring>`}}},DNu6:function(e,t,n){"use strict";n.d(t,"b",function(){return s}),n.d(t,"a",function(){return c});var a=n("zFbe"),i=n("/i08"),r=n("d3jT"),o=n("R2TB");const s="mgt-db-list";class c{static getCache(e,t){const n=`${e.name}/${t}`;return this.isInitialized||this.init(),this.cacheStore.has(t)||this.cacheStore.set(n,new r.a(e,t)),this.cacheStore.get(n)}static clearCacheById(e){const t=[],n=JSON.parse(localStorage.getItem(s));if(n){const a=[];n.forEach(n=>{n.includes(e)?t.push(new Promise((e,t)=>{const a=indexedDB.deleteDatabase(n);a.onsuccess=()=>e(),a.onerror=()=>{Object(o.a)(`${a.error.name} occurred deleting cache: ${n}`,a.error.message),t()}})):a.push(n)}),a.length>0?localStorage.setItem(s,JSON.stringify(a)):localStorage.removeItem(s)}return Promise.all(t)}static get config(){return this.cacheConfig}static init(){let e;a.a.globalProvider&&(e=a.a.globalProvider.state),a.a.onProviderUpdated(()=>{return t=this,void 0,r=function*(){if(e===i.c.SignedIn&&a.a.globalProvider.state===i.c.SignedOut){const e=yield a.a.getCacheId();null!==e&&(yield this.clearCacheById(e))}e=a.a.globalProvider.state},new((n=void 0)||(n=Promise))(function(e,a){function i(e){try{s(r.next(e))}catch(e){a(e)}}function o(e){try{s(r.throw(e))}catch(e){a(e)}}function s(t){var a;t.done?e(t.value):(a=t.value,a instanceof n?a:new n(function(e){e(a)})).then(i,o)}s((r=r.apply(t,[])).next())});var t,n,r}),this.isInitialized=!0}}c.cacheStore=new Map,c.isInitialized=!1,c.cacheConfig={defaultInvalidationPeriod:36e5,groups:{invalidationPeriod:null,isEnabled:!0},isEnabled:!0,people:{invalidationPeriod:null,isEnabled:!0},photos:{invalidationPeriod:null,isEnabled:!0},presence:{invalidationPeriod:3e5,isEnabled:!0},users:{invalidationPeriod:null,isEnabled:!0},response:{invalidationPeriod:null,isEnabled:!0},files:{invalidationPeriod:null,isEnabled:!0},fileLists:{invalidationPeriod:null,isEnabled:!0},conversation:{invalidationPeriod:432e6,isEnabled:!0}}},E7Zv:function(e,t,n){"use strict";n.d(t,"a",function(){return a});const a=(...e)=>{const t=e[0];let n=t;for(let t=1;t<e.length;++t){const a=e[t];n.setNext&&n.setNext(a),n=a}return t}},EKCE:function(e,t){var n;n=function(){return this}();try{n=n||new Function("return this")()}catch(e){"object"==typeof window&&(n=window)}e.exports=n},EYXE:function(e,t){e.exports=n},FGLN:function(e,t,n){"use strict";function a(...e){return e.every(e=>e instanceof HTMLElement)}function i(e,t){if(e&&t&&a(e))return Array.from(e.querySelectorAll(t)).filter(e=>null!==e.offsetParent)}let r;function o(){if("boolean"==typeof r)return r;if("undefined"==typeof window||!window.document||!window.document.createElement)return r=!1,r;const e=document.createElement("style"),t=function(){const e=document.querySelector('meta[property="csp-nonce"]');return e?e.getAttribute("content"):null}();null!==t&&e.setAttribute("nonce",t),document.head.appendChild(e);try{e.sheet.insertRule("foo:focus-visible {color:inherit}",0),r=!0}catch(e){r=!1}finally{document.head.removeChild(e)}return r}n.d(t,"c",function(){return a}),n.d(t,"b",function(){return i}),n.d(t,"a",function(){return o})},FVZ7:function(e,t,n){"use strict";n.d(t,"a",function(){return r}),n.d(t,"b",function(){return o});var a=n("4X57"),i=n("8hiW");const r=a.b`
  outline: calc(${i.y} * 1px) solid ${i.x};
  outline-offset: calc(${i.y} * -1px);
`,o=a.b`
  outline: calc(${i.y} * 1px) solid ${i.x};
  outline-offset: calc(${i.vb} * 1px);
`},GcG9:function(e,t,n){"use strict";n.d(t,"a",function(){return r}),n.d(t,"b",function(){return a}),n.d(t,"c",function(){return i});const a={ATTRIBUTE:1,CHILD:2,PROPERTY:3,BOOLEAN_ATTRIBUTE:4,EVENT:5,ELEMENT:6},i=e=>(...t)=>({_$litDirective$:e,values:t});class r{constructor(e){}get _$AU(){return this._$AM._$AU}_$AT(e,t,n){this._$Ct=e,this._$AM=t,this._$Ci=n}_$AS(e,t){return this.update(e,t)}update(e,t){return this.render(...t)}}},Gy7L:function(e,t,n){"use strict";var a;n.d(t,"b",function(){return i}),n.d(t,"c",function(){return r}),n.d(t,"d",function(){return o}),n.d(t,"e",function(){return s}),n.d(t,"g",function(){return c}),n.d(t,"h",function(){return d}),n.d(t,"j",function(){return l}),n.d(t,"f",function(){return u}),n.d(t,"i",function(){return f}),n.d(t,"k",function(){return p}),n.d(t,"l",function(){return m}),n.d(t,"m",function(){return _}),n.d(t,"n",function(){return h}),n.d(t,"a",function(){return b}),function(e){e[e.alt=18]="alt",e[e.arrowDown=40]="arrowDown",e[e.arrowLeft=37]="arrowLeft",e[e.arrowRight=39]="arrowRight",e[e.arrowUp=38]="arrowUp",e[e.back=8]="back",e[e.backSlash=220]="backSlash",e[e.break=19]="break",e[e.capsLock=20]="capsLock",e[e.closeBracket=221]="closeBracket",e[e.colon=186]="colon",e[e.colon2=59]="colon2",e[e.comma=188]="comma",e[e.ctrl=17]="ctrl",e[e.delete=46]="delete",e[e.end=35]="end",e[e.enter=13]="enter",e[e.equals=187]="equals",e[e.equals2=61]="equals2",e[e.equals3=107]="equals3",e[e.escape=27]="escape",e[e.forwardSlash=191]="forwardSlash",e[e.function1=112]="function1",e[e.function10=121]="function10",e[e.function11=122]="function11",e[e.function12=123]="function12",e[e.function2=113]="function2",e[e.function3=114]="function3",e[e.function4=115]="function4",e[e.function5=116]="function5",e[e.function6=117]="function6",e[e.function7=118]="function7",e[e.function8=119]="function8",e[e.function9=120]="function9",e[e.home=36]="home",e[e.insert=45]="insert",e[e.menu=93]="menu",e[e.minus=189]="minus",e[e.minus2=109]="minus2",e[e.numLock=144]="numLock",e[e.numPad0=96]="numPad0",e[e.numPad1=97]="numPad1",e[e.numPad2=98]="numPad2",e[e.numPad3=99]="numPad3",e[e.numPad4=100]="numPad4",e[e.numPad5=101]="numPad5",e[e.numPad6=102]="numPad6",e[e.numPad7=103]="numPad7",e[e.numPad8=104]="numPad8",e[e.numPad9=105]="numPad9",e[e.numPadDivide=111]="numPadDivide",e[e.numPadDot=110]="numPadDot",e[e.numPadMinus=109]="numPadMinus",e[e.numPadMultiply=106]="numPadMultiply",e[e.numPadPlus=107]="numPadPlus",e[e.openBracket=219]="openBracket",e[e.pageDown=34]="pageDown",e[e.pageUp=33]="pageUp",e[e.period=190]="period",e[e.print=44]="print",e[e.quote=222]="quote",e[e.scrollLock=145]="scrollLock",e[e.shift=16]="shift",e[e.space=32]="space",e[e.tab=9]="tab",e[e.tilde=192]="tilde",e[e.windowsLeft=91]="windowsLeft",e[e.windowsOpera=219]="windowsOpera",e[e.windowsRight=92]="windowsRight"}(a||(a={}));const i="ArrowDown",r="ArrowLeft",o="ArrowRight",s="ArrowUp",c="Enter",d="Escape",l="Home",u="End",f="F2",p="PageDown",m="PageUp",_=" ",h="Tab",b={ArrowDown:i,ArrowLeft:r,ArrowRight:o,ArrowUp:s}},HN6m:function(e,t,n){"use strict";n.d(t,"a",function(){return a});const a=new class{constructor(){this._disambiguation=""}get defaultPrefix(){return"mgt"}withDisambiguation(e){return e&&!this._disambiguation&&(this._disambiguation=e),this}get prefix(){return this._disambiguation?`${this.defaultPrefix}-${this._disambiguation}`:this.defaultPrefix}get disambiguation(){return this._disambiguation}get isDisambiguated(){return Boolean(this._disambiguation)}normalize(e){return this.isDisambiguated?e.toUpperCase().replace(this.prefix.toUpperCase(),this.defaultPrefix.toUpperCase()):e}}},KKsr:function(e,t,n){"use strict";n.d(t,"a",function(){return _}),n.d(t,"e",function(){return h}),n.d(t,"d",function(){return b}),n.d(t,"b",function(){return g}),n.d(t,"c",function(){return v});var a=n("4X57"),i=n("8GQ4"),r=n("2i1/"),o=n("j9Xn"),s=n("8hiW"),c=n("NLdk"),d=n("QkLF"),l=n("FVZ7");const u=i.a.create("input-placeholder-rest").withDefault(e=>{const t=s.N.getValueFor(e);return s.db.getValueFor(e).evaluate(e,t.evaluate(e).rest)}),f=i.a.create("input-placeholder-hover").withDefault(e=>{const t=s.N.getValueFor(e);return s.db.getValueFor(e).evaluate(e,t.evaluate(e).hover)}),p=i.a.create("input-filled-placeholder-rest").withDefault(e=>{const t=s.U.getValueFor(e);return s.db.getValueFor(e).evaluate(e,t.evaluate(e).rest)}),m=i.a.create("input-filled-placeholder-hover").withDefault(e=>{const t=s.U.getValueFor(e);return s.db.getValueFor(e).evaluate(e,t.evaluate(e).hover)}),_=(e,t,n)=>a.a`
  :host {
    ${c.a}
    color: ${s.fb};
    fill: currentcolor;
    user-select: none;
    position: relative;
  }

  ${n} {
    box-sizing: border-box;
    position: relative;
    color: inherit;
    border: calc(${s.vb} * 1px) solid transparent;
    border-radius: calc(${s.q} * 1px);
    height: calc(${d.a} * 1px);
    font-family: inherit;
    font-size: inherit;
    line-height: inherit;
  }

  .control {
    width: 100%;
    outline: none;
  }

  .label {
    display: block;
    color: ${s.fb};
    cursor: pointer;
    ${c.a}
    margin-bottom: 4px;
  }

  .label__hidden {
    display: none;
    visibility: hidden;
  }

  :host([disabled]) ${n},
  :host([readonly]) ${n},
  :host([disabled]) .label,
  :host([readonly]) .label,
  :host([disabled]) .control,
  :host([readonly]) .control {
    cursor: ${r.a};
  }

  :host([disabled]) {
    opacity: ${s.u};
  }
`,h=(e,t,n)=>a.a`
  @media (forced-colors: none) {
    :host(:not([disabled]):active)::after {
      left: 50%;
      width: 40%;
      transform: translateX(-50%);
      border-bottom-left-radius: 0;
      border-bottom-right-radius: 0;
    }

    :host(:not([disabled]):focus-within)::after {
      left: 0;
      width: 100%;
      transform: none;
    }

    :host(:not([disabled]):active)::after,
    :host(:not([disabled]):focus-within:not(:active))::after {
      content: '';
      position: absolute;
      height: calc(${s.y} * 1px);
      bottom: 0;
      border-bottom: calc(${s.y} * 1px) solid ${s.e};
      border-bottom-left-radius: calc(${s.q} * 1px);
      border-bottom-right-radius: calc(${s.q} * 1px);
      z-index: 2;
      transition: all 300ms cubic-bezier(0.1, 0.9, 0.2, 1);
    }
  }
`,b=(e,t,n,i=":not([disabled]):not(:focus-within)")=>a.a`
  ${n} {
    background: padding-box linear-gradient(${s.O}, ${s.O}),
      border-box ${s.pb};
  }

  :host(${i}:hover) ${n} {
    background: padding-box linear-gradient(${s.M}, ${s.M}),
      border-box ${s.ob};
  }

  :host(:not([disabled]):focus-within) ${n} {
    background: padding-box linear-gradient(${s.L}, ${s.L}),
      border-box ${s.pb};
  }
  
  :host([disabled]) ${n} {
    background: padding-box linear-gradient(${s.O}, ${s.O}),
      border-box ${s.rb};
  }

  .control::placeholder {
    color: ${u};
  }

  :host(${i}:hover) .control::placeholder {
    color: ${f};
  }
`,g=(e,t,n,i=":not([disabled]):not(:focus-within)")=>a.a`
  ${n} {
    background: ${s.V};
  }

  :host(${i}:hover) ${n} {
    background: ${s.T};
  }

  :host(:not([disabled]):focus-within) ${n} {
    background: ${s.S};
  }

  :host([disabled]) ${n} {
    background: ${s.V};
  }

  .control::placeholder {
    color: ${p};
  }

  :host(${i}:hover) .control::placeholder {
    color: ${m};
  }
`,v=(e,t,n,i=":not([disabled]):not(:focus-within)")=>a.a`
  :host {
    color: ${o.a.ButtonText};
  }

  ${n} {
    background: ${o.a.ButtonFace};
    border-color: ${o.a.ButtonText};
  }

  :host(${i}:hover) ${n},
  :host(:not([disabled]):focus-within) ${n} {
    border-color: ${o.a.Highlight};
  }

  :host([disabled]) ${n} {
    opacity: 1;
    background: ${o.a.ButtonFace};
    border-color: ${o.a.GrayText};
  }

  .control::placeholder,
  :host(${i}:hover) .control::placeholder {
    color: ${o.a.CanvasText};
  }

  :host(:not([disabled]):focus) ${n} {
    ${l.a}
    outline-color: ${o.a.Highlight};
  }

  :host([disabled]) {
    opacity: 1;
    color: ${o.a.GrayText};
  }

  :host([disabled]) ::placeholder,
  :host([disabled]) ::-webkit-input-placeholder {
    color: ${o.a.GrayText};
  }
`},Kct4:function(e,t,n){"use strict";n.d(t,"a",function(){return i});const a=(-.1+Math.sqrt(.21))/2;function i(e){return e.relativeLuminance<=a}},Mjrb:function(e,t,n){"use strict";n.d(t,"a",function(){return p});var a=n("eftJ"),i=n("QBS5"),r=n("oePG"),o=n("uXNP"),s=n("C5kU"),c=n("6fxl"),d=n("o87Z"),l=n("TDEi");class u extends l.a{}class f extends(Object(d.b)(u)){constructor(){super(...arguments),this.proxy=document.createElement("input")}}class p extends f{constructor(){super(...arguments),this.handleClick=e=>{var t;this.disabled&&(null===(t=this.defaultSlottedContent)||void 0===t?void 0:t.length)<=1&&e.stopPropagation()},this.handleSubmission=()=>{if(!this.form)return;const e=this.proxy.isConnected;e||this.attachProxy(),"function"==typeof this.form.requestSubmit?this.form.requestSubmit(this.proxy):this.proxy.click(),e||this.detachProxy()},this.handleFormReset=()=>{var e;null===(e=this.form)||void 0===e||e.reset()},this.handleUnsupportedDelegatesFocus=()=>{var e;window.ShadowRoot&&!window.ShadowRoot.prototype.hasOwnProperty("delegatesFocus")&&(null===(e=this.$fastController.definition.shadowOptions)||void 0===e?void 0:e.delegatesFocus)&&(this.focus=()=>{this.control.focus()})}}formactionChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.formAction=this.formaction)}formenctypeChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.formEnctype=this.formenctype)}formmethodChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.formMethod=this.formmethod)}formnovalidateChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.formNoValidate=this.formnovalidate)}formtargetChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.formTarget=this.formtarget)}typeChanged(e,t){this.proxy instanceof HTMLInputElement&&(this.proxy.type=this.type),"submit"===t&&this.addEventListener("click",this.handleSubmission),"submit"===e&&this.removeEventListener("click",this.handleSubmission),"reset"===t&&this.addEventListener("click",this.handleFormReset),"reset"===e&&this.removeEventListener("click",this.handleFormReset)}validate(){super.validate(this.control)}connectedCallback(){var e;super.connectedCallback(),this.proxy.setAttribute("type",this.type),this.handleUnsupportedDelegatesFocus();const t=Array.from(null===(e=this.control)||void 0===e?void 0:e.children);t&&t.forEach(e=>{e.addEventListener("click",this.handleClick)})}disconnectedCallback(){var e;super.disconnectedCallback();const t=Array.from(null===(e=this.control)||void 0===e?void 0:e.children);t&&t.forEach(e=>{e.removeEventListener("click",this.handleClick)})}}Object(a.a)([Object(i.c)({mode:"boolean"})],p.prototype,"autofocus",void 0),Object(a.a)([Object(i.c)({attribute:"form"})],p.prototype,"formId",void 0),Object(a.a)([i.c],p.prototype,"formaction",void 0),Object(a.a)([i.c],p.prototype,"formenctype",void 0),Object(a.a)([i.c],p.prototype,"formmethod",void 0),Object(a.a)([Object(i.c)({mode:"boolean"})],p.prototype,"formnovalidate",void 0),Object(a.a)([i.c],p.prototype,"formtarget",void 0),Object(a.a)([i.c],p.prototype,"type",void 0),Object(a.a)([r.d],p.prototype,"defaultSlottedContent",void 0);class m{}Object(a.a)([Object(i.c)({attribute:"aria-expanded"})],m.prototype,"ariaExpanded",void 0),Object(a.a)([Object(i.c)({attribute:"aria-pressed"})],m.prototype,"ariaPressed",void 0),Object(c.a)(m,o.a),Object(c.a)(p,s.a,m)},NLdk:function(e,t,n){"use strict";n.d(t,"a",function(){return r});var a=n("4X57"),i=n("8hiW");const r=a.b`
  font-family: ${i.p};
  font-size: ${i.wb};
  line-height: ${i.yb};
  font-weight: initial;
  font-variation-settings: ${i.xb};
`;a.b`
  font-family: ${i.p};
  font-size: ${i.zb};
  line-height: ${i.Bb};
  font-weight: initial;
  font-variation-settings: ${i.Ab};
`,a.b`
  font-family: ${i.p};
  font-size: ${i.Cb};
  line-height: ${i.Eb};
  font-weight: initial;
  font-variation-settings: ${i.Db};
`,a.b`
  font-family: ${i.p};
  font-size: ${i.Fb};
  line-height: ${i.Hb};
  font-weight: initial;
  font-variation-settings: ${i.Gb};
`,a.b`
  font-family: ${i.p};
  font-size: ${i.Ib};
  line-height: ${i.Kb};
  font-weight: initial;
  font-variation-settings: ${i.Jb};
`,a.b`
  font-family: ${i.p};
  font-size: ${i.Lb};
  line-height: ${i.Nb};
  font-weight: initial;
  font-variation-settings: ${i.Mb};
`,a.b`
  font-family: ${i.p};
  font-size: ${i.Ob};
  line-height: ${i.Qb};
  font-weight: initial;
  font-variation-settings: ${i.Pb};
`,a.b`
  font-family: ${i.p};
  font-size: ${i.Rb};
  line-height: ${i.Tb};
  font-weight: initial;
  font-variation-settings: ${i.Sb};
`,a.b`
  font-family: ${i.p};
  font-size: ${i.Ub};
  line-height: ${i.Wb};
  font-weight: initial;
  font-variation-settings: ${i.Vb};
`},Ne8q:function(e,t,n){"use strict";n.d(t,"a",function(){return i});const a=document.createRange();class i{constructor(e,t){this.fragment=e,this.behaviors=t,this.source=null,this.context=null,this.firstChild=e.firstChild,this.lastChild=e.lastChild}appendTo(e){e.appendChild(this.fragment)}insertBefore(e){if(this.fragment.hasChildNodes())e.parentNode.insertBefore(this.fragment,e);else{const t=this.lastChild;if(e.previousSibling===t)return;const n=e.parentNode;let a,i=this.firstChild;for(;i!==t;)a=i.nextSibling,n.insertBefore(i,e),i=a;n.insertBefore(t,e)}}remove(){const e=this.fragment,t=this.lastChild;let n,a=this.firstChild;for(;a!==t;)n=a.nextSibling,e.appendChild(a),a=n;e.appendChild(t)}dispose(){const e=this.firstChild.parentNode,t=this.lastChild;let n,a=this.firstChild;for(;a!==t;)n=a.nextSibling,e.removeChild(a),a=n;e.removeChild(t);const i=this.behaviors,r=this.source;for(let e=0,t=i.length;e<t;++e)i[e].unbind(r)}bind(e,t){const n=this.behaviors;if(this.source!==e)if(null!==this.source){const a=this.source;this.source=e,this.context=t;for(let i=0,r=n.length;i<r;++i){const r=n[i];r.unbind(a),r.bind(e,t)}}else{this.source=e,this.context=t;for(let a=0,i=n.length;a<i;++a)n[a].bind(e,t)}}unbind(){if(null===this.source)return;const e=this.behaviors,t=this.source;for(let n=0,a=e.length;n<a;++n)e[n].unbind(t);this.source=null}static disposeContiguousBatch(e){if(0!==e.length){a.setStartBefore(e[0].firstChild),a.setEndAfter(e[e.length-1].lastChild),a.deleteContents();for(let t=0,n=e.length;t<n;++t){const n=e[t],a=n.behaviors,i=n.source;for(let e=0,t=a.length;e<t;++e)a[e].unbind(i)}}}}},"O/9U":function(e,t,n){"use strict";n.d(t,"b",function(){return a}),n.d(t,"a",function(){return i});class a{constructor(e,t){this.sub1=void 0,this.sub2=void 0,this.spillover=void 0,this.source=e,this.sub1=t}has(e){return void 0===this.spillover?this.sub1===e||this.sub2===e:-1!==this.spillover.indexOf(e)}subscribe(e){const t=this.spillover;if(void 0===t){if(this.has(e))return;if(void 0===this.sub1)return void(this.sub1=e);if(void 0===this.sub2)return void(this.sub2=e);this.spillover=[this.sub1,this.sub2,e],this.sub1=void 0,this.sub2=void 0}else-1===t.indexOf(e)&&t.push(e)}unsubscribe(e){const t=this.spillover;if(void 0===t)this.sub1===e?this.sub1=void 0:this.sub2===e&&(this.sub2=void 0);else{const n=t.indexOf(e);-1!==n&&t.splice(n,1)}}notify(e){const t=this.spillover,n=this.source;if(void 0===t){const t=this.sub1,a=this.sub2;void 0!==t&&t.handleChange(n,e),void 0!==a&&a.handleChange(n,e)}else for(let a=0,i=t.length;a<i;++a)t[a].handleChange(n,e)}}class i{constructor(e){this.subscribers={},this.sourceSubscribers=null,this.source=e}notify(e){var t;const n=this.subscribers[e];void 0!==n&&n.notify(e),null===(t=this.sourceSubscribers)||void 0===t||t.notify(e)}subscribe(e,t){var n;if(t){let n=this.subscribers[t];void 0===n&&(this.subscribers[t]=n=new a(this.source)),n.subscribe(e)}else this.sourceSubscribers=null!==(n=this.sourceSubscribers)&&void 0!==n?n:new a(this.source),this.sourceSubscribers.subscribe(e)}unsubscribe(e,t){var n;if(t){const n=this.subscribers[t];void 0!==n&&n.unsubscribe(e)}else null===(n=this.sourceSubscribers)||void 0===n||n.unsubscribe(e)}}},P2Ap:function(e,t,n){"use strict";n.d(t,"a",function(){return A});var a=n("3rEL"),i=n("QBS5"),r=n("eftJ"),o=n("5ZAu"),s=n("oePG"),c=n("uXNP"),d=n("C5kU"),l=n("6fxl"),u=n("o87Z"),f=n("TDEi");class p extends f.a{}class m extends(Object(u.b)(p)){constructor(){super(...arguments),this.proxy=document.createElement("input")}}class _ extends m{constructor(){super(...arguments),this.type="text"}readOnlyChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.readOnly=this.readOnly,this.validate())}autofocusChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.autofocus=this.autofocus,this.validate())}placeholderChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.placeholder=this.placeholder)}typeChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.type=this.type,this.validate())}listChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.setAttribute("list",this.list),this.validate())}maxlengthChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.maxLength=this.maxlength,this.validate())}minlengthChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.minLength=this.minlength,this.validate())}patternChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.pattern=this.pattern,this.validate())}sizeChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.size=this.size)}spellcheckChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.spellcheck=this.spellcheck)}connectedCallback(){super.connectedCallback(),this.proxy.setAttribute("type",this.type),this.validate(),this.autofocus&&o.a.queueUpdate(()=>{this.focus()})}select(){this.control.select(),this.$emit("select")}handleTextInput(){this.value=this.control.value}handleChange(){this.$emit("change")}validate(){super.validate(this.control)}}Object(r.a)([Object(i.c)({attribute:"readonly",mode:"boolean"})],_.prototype,"readOnly",void 0),Object(r.a)([Object(i.c)({mode:"boolean"})],_.prototype,"autofocus",void 0),Object(r.a)([i.c],_.prototype,"placeholder",void 0),Object(r.a)([i.c],_.prototype,"type",void 0),Object(r.a)([i.c],_.prototype,"list",void 0),Object(r.a)([Object(i.c)({converter:i.e})],_.prototype,"maxlength",void 0),Object(r.a)([Object(i.c)({converter:i.e})],_.prototype,"minlength",void 0),Object(r.a)([i.c],_.prototype,"pattern",void 0),Object(r.a)([Object(i.c)({converter:i.e})],_.prototype,"size",void 0),Object(r.a)([Object(i.c)({mode:"boolean"})],_.prototype,"spellcheck",void 0),Object(r.a)([s.d],_.prototype,"defaultSlottedNodes",void 0);class h{}Object(l.a)(h,c.a),Object(l.a)(_,d.a,h);var b=n("6BDD"),g=n("UauI"),v=n("+53S"),y=n("6gCt"),S=n("4X57"),D=n("xY0q"),I=n("oqLQ"),x=n("KKsr"),C=n("57ob"),O=n("8hiW");const w=".root";class E extends _{appearanceChanged(e,t){e!==t&&(this.classList.add(t),this.classList.remove(e))}connectedCallback(){super.connectedCallback(),this.appearance||(this.appearance="outline")}}Object(a.a)([i.c],E.prototype,"appearance",void 0);const A=E.compose({baseName:"text-field",baseClass:_,template:(e,t)=>b.a`
    <template
        class="
            ${e=>e.readOnly?"readonly":""}
        "
    >
        <label
            part="label"
            for="control"
            class="${e=>e.defaultSlottedNodes&&e.defaultSlottedNodes.length?"label":"label label__hidden"}"
        >
            <slot
                ${Object(g.a)({property:"defaultSlottedNodes",filter:y.a})}
            ></slot>
        </label>
        <div class="root" part="root">
            ${Object(d.d)(e,t)}
            <input
                class="control"
                part="control"
                id="control"
                @input="${e=>e.handleTextInput()}"
                @change="${e=>e.handleChange()}"
                ?autofocus="${e=>e.autofocus}"
                ?disabled="${e=>e.disabled}"
                list="${e=>e.list}"
                maxlength="${e=>e.maxlength}"
                minlength="${e=>e.minlength}"
                pattern="${e=>e.pattern}"
                placeholder="${e=>e.placeholder}"
                ?readonly="${e=>e.readOnly}"
                ?required="${e=>e.required}"
                size="${e=>e.size}"
                ?spellcheck="${e=>e.spellcheck}"
                :value="${e=>e.value}"
                type="${e=>e.type}"
                aria-atomic="${e=>e.ariaAtomic}"
                aria-busy="${e=>e.ariaBusy}"
                aria-controls="${e=>e.ariaControls}"
                aria-current="${e=>e.ariaCurrent}"
                aria-describedby="${e=>e.ariaDescribedby}"
                aria-details="${e=>e.ariaDetails}"
                aria-disabled="${e=>e.ariaDisabled}"
                aria-errormessage="${e=>e.ariaErrormessage}"
                aria-flowto="${e=>e.ariaFlowto}"
                aria-haspopup="${e=>e.ariaHaspopup}"
                aria-hidden="${e=>e.ariaHidden}"
                aria-invalid="${e=>e.ariaInvalid}"
                aria-keyshortcuts="${e=>e.ariaKeyshortcuts}"
                aria-label="${e=>e.ariaLabel}"
                aria-labelledby="${e=>e.ariaLabelledby}"
                aria-live="${e=>e.ariaLive}"
                aria-owns="${e=>e.ariaOwns}"
                aria-relevant="${e=>e.ariaRelevant}"
                aria-roledescription="${e=>e.ariaRoledescription}"
                ${Object(v.a)("control")}
            />
            ${Object(d.b)(e,t)}
        </div>
    </template>
`,styles:(e,t)=>S.a`
    ${Object(D.a)("inline-block")}

    ${Object(x.a)(e,t,w)}

    ${Object(x.e)(e,t,w)}

    .root {
      display: flex;
      flex-direction: row;
    }

    .control {
      -webkit-appearance: none;
      color: inherit;
      background: transparent;
      border: 0;
      height: calc(100% - 4px);
      margin-top: auto;
      margin-bottom: auto;
      padding: 0 calc(${O.s} * 2px + 1px);
      font-family: inherit;
      font-size: inherit;
      line-height: inherit;
    }

    .start,
    .end {
      display: flex;
      margin: auto;
    }

    .start {
      display: flex;
      margin-inline-start: 11px;
    }

    .end {
      display: flex;
      margin-inline-end: 11px;
    }
  `.withBehaviors(Object(C.a)("outline",Object(x.d)(e,t,w)),Object(C.a)("filled",Object(x.b)(e,t,w)),Object(I.a)(Object(x.c)(e,t,w))),shadowOptions:{delegatesFocus:!0}})},P4Ao:function(e,t,n){"use strict";n.d(t,"a",function(){return f});var a=n("5ZAu"),i=n("O/9U"),r=n("oePG"),o=n("WniA");const s=new WeakMap,c={bubbles:!0,composed:!0,cancelable:!0};function d(e){return e.shadowRoot||s.get(e)||null}class l extends i.a{constructor(e,t){super(e),this.boundObservables=null,this.behaviors=null,this.needsInitialization=!0,this._template=null,this._styles=null,this._isConnected=!1,this.$fastController=this,this.view=null,this.element=e,this.definition=t;const n=t.shadowOptions;if(void 0!==n){const t=e.attachShadow(n);"closed"===n.mode&&s.set(e,t)}const a=r.b.getAccessors(e);if(a.length>0){const t=this.boundObservables=Object.create(null);for(let n=0,i=a.length;n<i;++n){const i=a[n].name,r=e[i];void 0!==r&&(delete e[i],t[i]=r)}}}get isConnected(){return r.b.track(this,"isConnected"),this._isConnected}setIsConnected(e){this._isConnected=e,r.b.notify(this,"isConnected")}get template(){return this._template}set template(e){this._template!==e&&(this._template=e,this.needsInitialization||this.renderTemplate(e))}get styles(){return this._styles}set styles(e){this._styles!==e&&(null!==this._styles&&this.removeStyles(this._styles),this._styles=e,this.needsInitialization||null===e||this.addStyles(e))}addStyles(e){const t=d(this.element)||this.element.getRootNode();if(e instanceof HTMLStyleElement)t.append(e);else if(!e.isAttachedTo(t)){const n=e.behaviors;e.addStylesTo(t),null!==n&&this.addBehaviors(n)}}removeStyles(e){const t=d(this.element)||this.element.getRootNode();if(e instanceof HTMLStyleElement)t.removeChild(e);else if(e.isAttachedTo(t)){const n=e.behaviors;e.removeStylesFrom(t),null!==n&&this.removeBehaviors(n)}}addBehaviors(e){const t=this.behaviors||(this.behaviors=new Map),n=e.length,a=[];for(let i=0;i<n;++i){const n=e[i];t.has(n)?t.set(n,t.get(n)+1):(t.set(n,1),a.push(n))}if(this._isConnected){const e=this.element;for(let t=0;t<a.length;++t)a[t].bind(e,r.c)}}removeBehaviors(e,t=!1){const n=this.behaviors;if(null===n)return;const a=e.length,i=[];for(let r=0;r<a;++r){const a=e[r];if(n.has(a)){const e=n.get(a)-1;0===e||t?n.delete(a)&&i.push(a):n.set(a,e)}}if(this._isConnected){const e=this.element;for(let t=0;t<i.length;++t)i[t].unbind(e)}}onConnectedCallback(){if(this._isConnected)return;const e=this.element;this.needsInitialization?this.finishInitialization():null!==this.view&&this.view.bind(e,r.c);const t=this.behaviors;if(null!==t)for(const[n]of t)n.bind(e,r.c);this.setIsConnected(!0)}onDisconnectedCallback(){if(!this._isConnected)return;this.setIsConnected(!1);const e=this.view;null!==e&&e.unbind();const t=this.behaviors;if(null!==t){const e=this.element;for(const[n]of t)n.unbind(e)}}onAttributeChangedCallback(e,t,n){const a=this.definition.attributeLookup[e];void 0!==a&&a.onAttributeChangedCallback(this.element,n)}emit(e,t,n){return!!this._isConnected&&this.element.dispatchEvent(new CustomEvent(e,Object.assign(Object.assign({detail:t},c),n)))}finishInitialization(){const e=this.element,t=this.boundObservables;if(null!==t){const n=Object.keys(t);for(let a=0,i=n.length;a<i;++a){const i=n[a];e[i]=t[i]}this.boundObservables=null}const n=this.definition;null===this._template&&(this.element.resolveTemplate?this._template=this.element.resolveTemplate():n.template&&(this._template=n.template||null)),null!==this._template&&this.renderTemplate(this._template),null===this._styles&&(this.element.resolveStyles?this._styles=this.element.resolveStyles():n.styles&&(this._styles=n.styles||null)),null!==this._styles&&this.addStyles(this._styles),this.needsInitialization=!1}renderTemplate(e){const t=this.element,n=d(t)||t;null!==this.view?(this.view.dispose(),this.view=null):this.needsInitialization||a.a.removeChildNodes(n),e&&(this.view=e.render(t,n,t))}static forCustomElement(e){const t=e.$fastController;if(void 0!==t)return t;const n=o.a.forType(e.constructor);if(void 0===n)throw new Error("Missing FASTElement definition.");return e.$fastController=new l(e,n)}}function u(e){return class extends e{constructor(){super(),l.forCustomElement(this)}$emit(e,t,n){return this.$fastController.emit(e,t,n)}connectedCallback(){this.$fastController.onConnectedCallback()}disconnectedCallback(){this.$fastController.onDisconnectedCallback()}attributeChangedCallback(e,t,n){this.$fastController.onAttributeChangedCallback(e,t,n)}}}const f=Object.assign(u(HTMLElement),{from:e=>u(e),define:(e,t)=>new o.a(e,t).define().type})},PT2o:function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("0q6d");class i{constructor(e,t,n){this.h=e,this.s=t,this.l=n}static fromObject(e){return!e||isNaN(e.h)||isNaN(e.s)||isNaN(e.l)?null:new i(e.h,e.s,e.l)}equalValue(e){return this.h===e.h&&this.s===e.s&&this.l===e.l}roundToPrecision(e){return new i(Object(a.i)(this.h,e),Object(a.i)(this.s,e),Object(a.i)(this.l,e))}toObject(){return{h:this.h,s:this.s,l:this.l}}}},Q5AN:function(e,t,n){"use strict";n.d(t,"b",function(){return r}),n.d(t,"a",function(){return o});var a=n("oePG"),i=n("oZuh");function r(e){return e?function(t,n,a){return 1===t.nodeType&&t.matches(e)}:function(e,t,n){return 1===e.nodeType}}class o{constructor(e,t){this.target=e,this.options=t,this.source=null}bind(e){const t=this.options.property;this.shouldUpdate=a.b.getAccessors(e).some(e=>e.name===t),this.source=e,this.updateTarget(this.computeNodes()),this.shouldUpdate&&this.observe()}unbind(){this.updateTarget(i.d),this.source=null,this.shouldUpdate&&this.disconnect()}handleEvent(){this.updateTarget(this.computeNodes())}computeNodes(){let e=this.getNodes();return void 0!==this.options.filter&&(e=e.filter(this.options.filter)),e}updateTarget(e){this.source[this.options.property]=e}}},QBS5:function(e,t,n){"use strict";n.d(t,"a",function(){return o}),n.d(t,"d",function(){return s}),n.d(t,"e",function(){return c}),n.d(t,"b",function(){return d}),n.d(t,"c",function(){return l});var a=n("oePG"),i=n("5ZAu"),r=n("oZuh");const o=Object.freeze({locate:Object(r.c)()}),s={toView:e=>e?"true":"false",fromView:e=>null!=e&&"false"!==e&&!1!==e&&0!==e},c={toView(e){if(null==e)return null;const t=1*e;return isNaN(t)?null:t.toString()},fromView(e){if(null==e)return null;const t=1*e;return isNaN(t)?null:t}};class d{constructor(e,t,n=t.toLowerCase(),a="reflect",i){this.guards=new Set,this.Owner=e,this.name=t,this.attribute=n,this.mode=a,this.converter=i,this.fieldName=`_${t}`,this.callbackName=`${t}Changed`,this.hasCallback=this.callbackName in e.prototype,"boolean"===a&&void 0===i&&(this.converter=s)}setValue(e,t){const n=e[this.fieldName],a=this.converter;void 0!==a&&(t=a.fromView(t)),n!==t&&(e[this.fieldName]=t,this.tryReflectToAttribute(e),this.hasCallback&&e[this.callbackName](n,t),e.$fastController.notify(this.name))}getValue(e){return a.b.track(e,this.name),e[this.fieldName]}onAttributeChangedCallback(e,t){this.guards.has(e)||(this.guards.add(e),this.setValue(e,t),this.guards.delete(e))}tryReflectToAttribute(e){const t=this.mode,n=this.guards;n.has(e)||"fromView"===t||i.a.queueUpdate(()=>{n.add(e);const a=e[this.fieldName];switch(t){case"reflect":const t=this.converter;i.a.setAttribute(e,this.attribute,void 0!==t?t.toView(a):a);break;case"boolean":i.a.setBooleanAttribute(e,this.attribute,a)}n.delete(e)})}static collect(e,...t){const n=[];t.push(o.locate(e));for(let a=0,i=t.length;a<i;++a){const i=t[a];if(void 0!==i)for(let t=0,a=i.length;t<a;++t){const a=i[t];"string"==typeof a?n.push(new d(e,a)):n.push(new d(e,a.property,a.attribute,a.mode,a.converter))}}return n}}function l(e,t){let n;function a(e,t){arguments.length>1&&(n.property=t),o.locate(e.constructor).push(n)}return arguments.length>1?(n={},void a(e,t)):(n=void 0===e?{}:e,a)}},QkLF:function(e,t,n){"use strict";n.d(t,"a",function(){return r});var a=n("4X57"),i=n("8hiW");const r=a.b`(${i.n} + ${i.r}) * ${i.s}`},R2TB:function(e,t,n){"use strict";n.d(t,"b",function(){return i}),n.d(t,"c",function(){return r}),n.d(t,"a",function(){return o});const a=(e,...t)=>[(new Date).toISOString(),"🦒: ",e,t],i=(e,...t)=>console.log(...a(e,t)),r=(e,...t)=>console.warn(...a(e,t)),o=(e,...t)=>console.error(...a(e,t))},RIOo:function(e,t,n){"use strict";n.d(t,"a",function(){return r});var a=n("EYXE"),i=n("zFbe");const r=(...e)=>{const t={scopes:e};return i.a.globalProvider.isIncrementalConsentDisabled?[]:[new a.AuthenticationHandlerOptions(void 0,t)]}},"RN7+":function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("0q6d");class i{constructor(e,t,n){this.l=e,this.a=t,this.b=n}static fromObject(e){return!e||isNaN(e.l)||isNaN(e.a)||isNaN(e.b)?null:new i(e.l,e.a,e.b)}equalValue(e){return this.l===e.l&&this.a===e.a&&this.b===e.b}roundToPrecision(e){return new i(Object(a.i)(this.l,e),Object(a.i)(this.a,e),Object(a.i)(this.b,e))}toObject(){return{l:this.l,a:this.a,b:this.b}}}i.epsilon=216/24389,i.kappa=24389/27},SUtl:function(e,t,n){"use strict";n.d(t,"a",function(){return h}),n.d(t,"b",function(){return b});var a=n("EYXE"),i=n("/49y"),r=n("4Eu1"),o=n("RIOo");class s{constructor(e,t,n,a){this.resource=n.startsWith("/")?n:`/${n}`,this.method=a,this.index=e,this.id=t}}var c=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};class d{static get baseUrl(){return"https://graph.microsoft.com"}constructor(e){this.graph=e,this.allRequests=[],this.requestsQueue=[],this.scopes=[],this.nextIndex=0,this.retryAfter=0}get hasRequests(){return this.requestsQueue.length>0}get(e,t,n,a){const i=this.nextIndex++,r=new s(i,e,t,"GET");r.headers=a,this.allRequests.push(r),this.requestsQueue.push(i),n&&(this.scopes=this.scopes.concat(n))}executeNext(){return c(this,void 0,void 0,function*(){const e=new Map;var t,n,i;if(this.retryAfter&&(yield(t=1e3*this.retryAfter,void 0,void 0,n=void 0,i=function*(){return new Promise(e=>{setTimeout(e,t)})},new(n||(n=Promise))(function(e,t){function a(e){try{o(i.next(e))}catch(e){t(e)}}function r(e){try{o(i.throw(e))}catch(e){t(e)}}function o(t){var i;t.done?e(t.value):(i=t.value,i instanceof n?i:new n(function(e){e(i)})).then(a,r)}o((i=i.apply(undefined,[])).next())})),this.retryAfter=0),!this.hasRequests)return e;const s=this.requestsQueue.splice(0,20),c=new a.BatchRequestContent;for(const e of s.map(e=>this.allRequests[e]))c.addRequest({id:e.index.toString(),request:new Request(d.baseUrl+e.resource,{method:e.method,headers:e.headers})});const l=this.scopes.length?Object(o.a)(...this.scopes):[],u=this.graph.api("$batch").middlewareOptions(l),f=yield c.getContent(),p=yield u.post(f);for(const t of p.responses){const n=new r.a,a=parseInt(t.id,10),i=this.allRequests[a];if(n.id=i.id,n.index=i.index,n.headers=t.headers,200===t.status)"string"==typeof t.body?t.headers["Content-Type"].includes("image/jpeg")?n.content="data:image/jpeg;base64,"+t.body:t.headers["Content-Type"].includes("image/pjpeg")?n.content="data:image/pjpeg;base64,"+t.body:t.headers["Content-Type"].includes("image/png")&&(n.content="data:image/png;base64,"+t.body):n.content=t.body,e.set(i.id,n);else if(429===t.status){this.requestsQueue.unshift(a);const e=t.headers["Retry-After"];this.retryAfter=Math.max(this.retryAfter,parseInt(e,10)||1)}}return e})}executeAll(){return c(this,void 0,void 0,function*(){const e=new Map;for(;this.hasRequests;){const t=yield this.executeNext();for(const[n,a]of t)e.set(n,a)}return e})}}class l{constructor(e){this.componentName="string"==typeof e?e:e.tagName.toLowerCase()}}var u=n("E7Zv");Object.create,Object.create,"function"==typeof SuppressedError&&SuppressedError;var f=n("q/PQ");class p{constructor(e,t){this._packageVersion=e,this._providerName=t}execute(e){var t,n,a,i;return n=this,void 0,i=function*(){try{if("string"==typeof e.request)if(Object(f.a)(e.request)){const t=[],n=e.middlewareControl.getMiddlewareOptions(l);if(n){const e=`${n.componentName}/${this._packageVersion}`;t.push(e)}if(this._providerName){const e=`${this._providerName}/${this._packageVersion}`;t.push(e)}const a=`mgt/${this._packageVersion}`;t.push(a),t.push(((e,t,n)=>{let a=null;if("undefined"!=typeof Request&&e instanceof Request)a=e.headers.get(n);else if(void 0!==t&&void 0!==t.headers)if("undefined"!=typeof Headers&&t.headers instanceof Headers)a=t.headers.get(n);else if(t.headers instanceof Array){const e=t.headers;for(let t=0,i=e.length;t<i;t++)if(e[t][0]===n){a=e[t][1];break}}else void 0!==t.headers[n]&&(a=t.headers[n]);return a})(e.request,e.options,"SdkVersion"));const i=t.join(", ");((e,t,n,a)=>{if("undefined"!=typeof Request&&e instanceof Request)e.headers.set(n,a);else if(void 0!==t)if(void 0===t.headers)t.headers=new Headers({[n]:a});else if("undefined"!=typeof Headers&&t.headers instanceof Headers)t.headers.set(n,a);else if(t.headers instanceof Array){let e=0;const i=t.headers.length;for(;e<i;e++){const i=t.headers[e];if(i[0]===n){i[1]=a;break}}e===i&&t.headers.push([n,a])}else Object.assign(t.headers,{[n]:a})})(e.request,e.options,"SdkVersion",i)}else null===(t=null==e?void 0:e.options)||void 0===t||delete t.headers.SdkVersion}catch(e){}return yield this._nextMiddleware.execute(e)},new((a=void 0)||(a=Promise))(function(e,t){function r(e){try{s(i.next(e))}catch(e){t(e)}}function o(e){try{s(i.throw(e))}catch(e){t(e)}}function s(t){var n;t.done?e(t.value):(n=t.value,n instanceof a?n:new a(function(e){e(n)})).then(r,o)}s((i=i.apply(n,[])).next())})}setNext(e){this._nextMiddleware=e}}var m=n("2oY9"),_=n("HN6m");class h{get client(){return this._client}get componentName(){return this._componentName}get version(){return this._version}constructor(e,t="v1.0"){this._client=e,this._version=t}forComponent(e,t){const n=new h(this._client,t||this._version);return n.setComponent(e),n}api(e){let t=this._client.api(e).version(this._version);return this._componentName&&(t.middlewareOptions=e=>(t._middlewareOptions=t._middlewareOptions.concat(e),t),t=t.middlewareOptions([new l(this._componentName)])),t}createBatch(){return new d(this)}setComponent(e){this._componentName=e instanceof Element?_.a.normalize(e.tagName):e}}const b=(e,t,n)=>{const r=[new a.AuthenticationHandler(e),new a.RetryHandler(new a.RetryHandlerOptions),new a.TelemetryHandler,new p(m.a,e.name),new a.HTTPMessageHandler],o=e.baseURL?e.baseURL:i.a,s=a.Client.initWithMiddleware({middleware:Object(u.a)(...r),customHosts:e.customHosts?new Set(e.customHosts):null,baseUrl:o}),c=new h(s,t);return n?c.forComponent(n):c}},"SiT+":function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("0q6d");class i{constructor(e,t,n){this.l=e,this.c=t,this.h=n}static fromObject(e){return!e||isNaN(e.l)||isNaN(e.c)||isNaN(e.h)?null:new i(e.l,e.c,e.h)}equalValue(e){return this.l===e.l&&this.c===e.c&&this.h===e.h}roundToPrecision(e){return new i(Object(a.i)(this.l,e),Object(a.i)(this.c,e),Object(a.i)(this.h,e))}toObject(){return{l:this.l,c:this.c,h:this.h}}}},TDEi:function(e,t,n){"use strict";n.d(t,"a",function(){return s});var a=n("eftJ"),i=n("P4Ao"),r=n("oePG"),o=n("4gKT");class s extends i.a{constructor(){super(...arguments),this._presentation=void 0}get $presentation(){return void 0===this._presentation&&(this._presentation=o.a.forTag(this.tagName,this)),this._presentation}templateChanged(){void 0!==this.template&&(this.$fastController.template=this.template)}stylesChanged(){void 0!==this.styles&&(this.$fastController.styles=this.styles)}connectedCallback(){null!==this.$presentation&&this.$presentation.applyTo(this),super.connectedCallback()}static compose(e){return(t={})=>new d(this===s?class extends s{}:this,e,t)}}function c(e,t,n){return"function"==typeof e?e(t,n):e}Object(a.a)([r.d],s.prototype,"template",void 0),Object(a.a)([r.d],s.prototype,"styles",void 0);class d{constructor(e,t,n){this.type=e,this.elementDefinition=t,this.overrideDefinition=n,this.definition=Object.assign(Object.assign({},this.elementDefinition),this.overrideDefinition)}register(e,t){const n=this.definition,a=this.overrideDefinition,i=`${n.prefix||t.elementPrefix}-${n.baseName}`;t.tryDefineElement({name:i,type:this.type,baseClass:this.elementDefinition.baseClass,callback:e=>{const t=new o.b(c(n.template,e,n),c(n.styles,e,n));e.definePresentation(t);let i=c(n.shadowOptions,e,n);e.shadowRootMode&&(i?a.shadowOptions||(i.mode=e.shadowRootMode):null!==i&&(i={mode:e.shadowRootMode})),e.defineElement({elementOptions:c(n.elementOptions,e,n),shadowOptions:i,attributes:c(n.attributes,e,n)})}})}}},THNU:function(e,t,n){"use strict";function a(e,t){const n=e.relativeLuminance>t.relativeLuminance?e:t,a=e.relativeLuminance>t.relativeLuminance?t:e;return(n.relativeLuminance+.05)/(a.relativeLuminance+.05)}n.d(t,"a",function(){return a})},UauI:function(e,t,n){"use strict";n.d(t,"a",function(){return o});var a=n("/w5G"),i=n("Q5AN");class r extends i.a{constructor(e,t){super(e,t)}observe(){this.target.addEventListener("slotchange",this)}disconnect(){this.target.removeEventListener("slotchange",this)}getNodes(){return this.target.assignedNodes(this.options)}}function o(e){return"string"==typeof e&&(e={property:e}),new a.a("fast-slotted",r,e)}},V07F:function(e,t,n){"use strict";n.d(t,"a",function(){return f});var a=n("5rba"),i=n("GcG9");const{D:r}=a.a,o=()=>document.createComment(""),s=(e,t,n)=>{const a=e._$AA.parentNode,i=void 0===t?e._$AB:t._$AA;if(void 0===n){const t=a.insertBefore(o(),i),s=a.insertBefore(o(),i);n=new r(t,s,e,e.options)}else{const t=n._$AB.nextSibling,r=n._$AM,o=r!==e;if(o){var s,c;let t;null!==(s=(c=n)._$AQ)&&void 0!==s&&s.call(c,e),n._$AM=e,void 0!==n._$AP&&(t=e._$AU)!==r._$AU&&n._$AP(t)}if(t!==i||o){let e=n._$AA;for(;e!==t;){const t=e.nextSibling;a.insertBefore(e,i),e=t}}}return n},c=(e,t,n=e)=>(e._$AI(t,n),e),d={},l=e=>{var t;null===(t=e._$AP)||void 0===t||t.call(e,!1,!0);let n=e._$AA;const a=e._$AB.nextSibling;for(;n!==a;){const e=n.nextSibling;n.remove(),n=e}},u=(e,t,n)=>{const a=new Map;for(let i=t;i<=n;i++)a.set(e[i],i);return a},f=Object(i.c)(class extends i.a{constructor(e){if(super(e),e.type!==i.b.CHILD)throw Error("repeat() can only be used in text expressions")}ht(e,t,n){let a;void 0===n?n=t:void 0!==t&&(a=t);const i=[],r=[];let o=0;for(const t of e)i[o]=a?a(t,o):o,r[o]=n(t,o),o++;return{values:r,keys:i}}render(e,t,n){return this.ht(e,t,n).values}update(e,[t,n,i]){var r;const o=e._$AH,{values:f,keys:p}=this.ht(t,n,i);if(!Array.isArray(o))return this.dt=p,f;const m=null!==(r=this.dt)&&void 0!==r?r:this.dt=[],_=[];let h,b,g=0,v=o.length-1,y=0,S=f.length-1;for(;g<=v&&y<=S;)if(null===o[g])g++;else if(null===o[v])v--;else if(m[g]===p[y])_[y]=c(o[g],f[y]),g++,y++;else if(m[v]===p[S])_[S]=c(o[v],f[S]),v--,S--;else if(m[g]===p[S])_[S]=c(o[g],f[S]),s(e,_[S+1],o[g]),g++,S--;else if(m[v]===p[y])_[y]=c(o[v],f[y]),s(e,o[g],o[v]),v--,y++;else if(void 0===h&&(h=u(p,y,S),b=u(m,g,v)),h.has(m[g]))if(h.has(m[v])){const t=b.get(p[y]),n=void 0!==t?o[t]:null;if(null===n){const t=s(e,o[g]);c(t,f[y]),_[y]=t}else _[y]=c(n,f[y]),s(e,o[g],n),o[t]=null;y++}else l(o[v]),v--;else l(o[g]),g++;for(;y<=S;){const t=s(e,_[S+1]);c(t,f[y]),_[y++]=t}for(;g<=v;){const e=o[g++];null!==e&&l(e)}return this.dt=p,((e,t=d)=>{e._$AH=t})(e,_),a.c}})},VDxL:function(e,t,n){"use strict";n.d(t,"a",function(){return b});var a=n("WniA"),i=n("TDEi"),r=n("q5bN"),o=n("8GQ4"),s=n("4gKT");const c=Object.freeze({definitionCallbackOnly:null,ignoreDuplicate:Symbol()}),d=new Map,l=new Map;let u=null;const f=r.a.createInterface(e=>e.cachedCallback(e=>(null===u&&(u=new m(null,e)),u))),p=Object.freeze({tagFor:e=>l.get(e),responsibleFor:e=>e.$$designSystem$$||r.a.findResponsibleContainer(e).get(f),getOrCreate(e){if(!e)return null===u&&(u=r.a.getOrCreateDOMContainer().get(f)),u;const t=e.$$designSystem$$;if(t)return t;const n=r.a.getOrCreateDOMContainer(e);if(n.has(f,!1))return n.get(f);{const t=new m(e,n);return n.register(r.b.instance(f,t)),t}}});class m{constructor(e,t){this.owner=e,this.container=t,this.designTokensInitialized=!1,this.prefix="fast",this.shadowRootMode=void 0,this.disambiguate=()=>c.definitionCallbackOnly,null!==e&&(e.$$designSystem$$=this)}withPrefix(e){return this.prefix=e,this}withShadowRootMode(e){return this.shadowRootMode=e,this}withElementDisambiguation(e){return this.disambiguate=e,this}withDesignTokenRoot(e){return this.designTokenRoot=e,this}register(...e){const t=this.container,n=[],a=this.disambiguate,r=this.shadowRootMode,s={elementPrefix:this.prefix,tryDefineElement(e,o,s){const u=function(e,t,n){return"string"==typeof e?{name:e,type:t,callback:n}:e}(e,o,s),{name:f,callback:p,baseClass:m}=u;let{type:h}=u,b=f,g=d.get(b),v=!0;for(;g;){const e=a(b,h,g);switch(e){case c.ignoreDuplicate:return;case c.definitionCallbackOnly:v=!1,g=void 0;break;default:b=e,g=d.get(b)}}v&&((l.has(h)||h===i.a)&&(h=class extends h{}),d.set(b,h),l.set(h,b),m&&l.set(m,b)),n.push(new _(t,b,h,r,p,v))}};this.designTokensInitialized||(this.designTokensInitialized=!0,null!==this.designTokenRoot&&o.a.registerRoot(this.designTokenRoot)),t.registerWithContext(s,...e);for(const e of n)e.callback(e),e.willDefine&&null!==e.definition&&e.definition.define();return this}}class _{constructor(e,t,n,a,i,r){this.container=e,this.name=t,this.type=n,this.shadowRootMode=a,this.callback=i,this.willDefine=r,this.definition=null}definePresentation(e){s.a.define(this.name,e,this.container)}defineElement(e){this.definition=new a.a(this.type,Object.assign(Object.assign({},e),{name:this.name}))}tagFor(e){return p.tagFor(e)}}const h=p.getOrCreate(void 0).withPrefix("fluent"),b=(...e)=>{if(null==e?void 0:e.length)for(const t of e)h.register(t())}},VRJB:function(e,t,n){"use strict";function a(e,t,n){return Math.min(Math.max(n,e),t)}function i(e,t,n=0){return[t,n]=[t,n].sort((e,t)=>e-t),t<=e&&e<n}n.d(t,"b",function(){return a}),n.d(t,"a",function(){return i})},VcPv:function(e,t,n){"use strict";n.d(t,"a",function(){return p});var a=n("tezw"),i=n("6BDD"),r=n("6vBc"),o=n("4X57"),s=n("j9Xn"),c=n("xY0q"),d=n("oqLQ"),l=n("QkLF"),u=n("8hiW");class f extends a.a{}const p=f.compose({baseName:"progress-ring",template:(e,t)=>i.a`
    <template
        role="progressbar"
        aria-valuenow="${e=>e.value}"
        aria-valuemin="${e=>e.min}"
        aria-valuemax="${e=>e.max}"
        class="${e=>e.paused?"paused":""}"
    >
        ${Object(r.a)(e=>"number"==typeof e.value,i.a`
                <svg
                    class="progress"
                    part="progress"
                    viewBox="0 0 16 16"
                    slot="determinate"
                >
                    <circle
                        class="background"
                        part="background"
                        cx="8px"
                        cy="8px"
                        r="7px"
                    ></circle>
                    <circle
                        class="determinate"
                        part="determinate"
                        style="stroke-dasharray: ${e=>44*e.percentComplete/100}px ${44}px"
                        cx="8px"
                        cy="8px"
                        r="7px"
                    ></circle>
                </svg>
            `,i.a`
                <slot name="indeterminate" slot="indeterminate">
                    ${t.indeterminateIndicator||""}
                </slot>
            `)}
    </template>
`,styles:(e,t)=>o.a`
    ${Object(c.a)("flex")} :host {
      align-items: center;
      height: calc(${l.a} * 1px);
      width: calc(${l.a} * 1px);
    }

    .progress {
      height: 100%;
      width: 100%;
    }

    .background {
      fill: none;
      stroke-width: 2px;
    }

    .determinate {
      stroke: ${u.e};
      fill: none;
      stroke-width: 2px;
      stroke-linecap: round;
      transform-origin: 50% 50%;
      transform: rotate(-90deg);
      transition: all 0.2s ease-in-out;
    }

    .indeterminate-indicator-1 {
      stroke: ${u.e};
      fill: none;
      stroke-width: 2px;
      stroke-linecap: round;
      transform-origin: 50% 50%;
      transform: rotate(-90deg);
      transition: all 0.2s ease-in-out;
      animation: spin-infinite 2s linear infinite;
    }

    :host(.paused) .indeterminate-indicator-1 {
      animation: none;
      stroke: ${u.cb};
    }

    :host(.paused) .determinate {
      stroke: ${u.cb};
    }

    @keyframes spin-infinite {
      0% {
        stroke-dasharray: 0.01px 43.97px;
        transform: rotate(0deg);
      }
      50% {
        stroke-dasharray: 21.99px 21.99px;
        transform: rotate(450deg);
      }
      100% {
        stroke-dasharray: 0.01px 43.97px;
        transform: rotate(1080deg);
      }
    }
  `.withBehaviors(Object(d.a)(o.a`
        .background {
          stroke: ${s.a.Field};
        }
        .determinate,
        .indeterminate-indicator-1 {
          stroke: ${s.a.ButtonText};
        }
        :host(.paused) .determinate,
        :host(.paused) .indeterminate-indicator-1 {
          stroke: ${s.a.GrayText};
        }
      `)),indeterminateIndicator:'\n    <svg class="progress" part="progress" viewBox="0 0 16 16">\n        <circle\n            class="background"\n            part="background"\n            cx="8px"\n            cy="8px"\n            r="7px"\n        ></circle>\n        <circle\n            class="indeterminate-indicator-1"\n            part="indeterminate-indicator-1"\n            cx="8px"\n            cy="8px"\n            r="7px"\n        ></circle>\n    </svg>\n  '})},WHff:function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};class i{get value(){return this._value}get hasNext(){return Boolean(this._nextLink)}static create(e,t,n){return a(this,void 0,void 0,function*(){const a=yield t.get();if(null==a?void 0:a.value){const t=new i;return t._graph=e,t._value=a.value,t._nextLink=a["@odata.nextLink"],t._version=n||e.version,t}return null})}static createFromValue(e,t,n=null){const a=new i;return a._graph=e,a._value=t,a._nextLink=n,a._version=e.version,a}get nextLink(){return this._nextLink||""}next(){var e;return a(this,void 0,void 0,function*(){if(this._nextLink){const t=this._nextLink.split(this._version)[1],n=yield this._graph.api(t).version(this._version).get();if(null===(e=null==n?void 0:n.value)||void 0===e?void 0:e.length)return this._value=this._value.concat(n.value),this._nextLink=n["@odata.nextLink"],n.value}return null})}}},WniA:function(e,t,n){"use strict";n.d(t,"a",function(){return l});var a=n("oZuh"),i=n("oePG"),r=n("+yEz"),o=n("QBS5");const s={mode:"open"},c={},d=a.b.getById(4,()=>{const e=new Map;return Object.freeze({register:t=>!e.has(t.type)&&(e.set(t.type,t),!0),getByType:t=>e.get(t)})});class l{constructor(e,t=e.definition){"string"==typeof t&&(t={name:t}),this.type=e,this.name=t.name,this.template=t.template;const n=o.b.collect(e,t.attributes),a=new Array(n.length),i={},d={};for(let e=0,t=n.length;e<t;++e){const t=n[e];a[e]=t.attribute,i[t.name]=t,d[t.attribute]=t}this.attributes=n,this.observedAttributes=a,this.propertyLookup=i,this.attributeLookup=d,this.shadowOptions=void 0===t.shadowOptions?s:null===t.shadowOptions?void 0:Object.assign(Object.assign({},s),t.shadowOptions),this.elementOptions=void 0===t.elementOptions?c:Object.assign(Object.assign({},c),t.elementOptions),this.styles=void 0===t.styles?void 0:Array.isArray(t.styles)?r.a.create(t.styles):t.styles instanceof r.a?t.styles:r.a.create([t.styles])}get isDefined(){return!!d.getByType(this.type)}define(e=customElements){const t=this.type;if(d.register(this)){const e=this.attributes,n=t.prototype;for(let t=0,a=e.length;t<a;++t)i.b.defineProperty(n,e[t]);Reflect.defineProperty(t,"observedAttributes",{value:this.observedAttributes,enumerable:!0})}return e.get(this.name)||e.define(this.name,t,this.elementOptions),this}}l.forType=d.getByType},Y1A4:function(e,t,n){"use strict";n.d(t,"a",function(){return l});var a=n("ZgG/"),i=n("qqMp"),r=n("ylMe"),o=n("dblZ"),s=n("oaDa"),c=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},d=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};class l extends o.b{constructor(){super(),this.templates={},this._renderedSlots=!1,this._renderedTemplates={},this._slotNamesAddedDuringRender=[],this.templateContext=this.templateContext||{}}update(e){this.templates=this.getTemplates(),this._slotNamesAddedDuringRender=[],super.update(e)}updated(e){super.updated(e),this.removeUnusedSlottedElements()}renderTemplate(e,t,n){if(!this.hasTemplate(e))return null;n=n||e,this._slotNamesAddedDuringRender.push(n),this._renderedSlots=!0;const a=i.c`
      <slot name=${n}></slot>
    `,o=Object.assign(Object.assign({},t),this.templateContext);if(Object.prototype.hasOwnProperty.call(this._renderedTemplates,n)){const{context:e,slot:t}=this._renderedTemplates[n];if(Object(r.b)(e,o))return a;this.removeChild(t)}const c=document.createElement("div");c.slot=n,c.dataset.generated="template",s.a.renderTemplate(c,this.templates[e],o),this.appendChild(c),this._renderedTemplates[n]={context:o,slot:c};const d={templateType:e,context:o,element:c};return this.fireCustomEvent("templateRendered",d),a}hasTemplate(e){var t;return Boolean(null===(t=this.templates)||void 0===t?void 0:t[e])}getTemplates(){const e={};for(let t=0;t<this.children.length;t++){const n=this.children[t];if("TEMPLATE"===n.nodeName){const a=n;a.dataset.type?e[a.dataset.type]=a:e.default=a,a.templateOrder=t}}return e}renderError(){return this.hasTemplate("error")?this.renderTemplate("error",this.error):i.c`
      <div class="error">
        ${this.error}
      </div>
    `}removeUnusedSlottedElements(){var e;if(this._renderedSlots){for(let t=0;t<this.children.length;t++){const n=this.children[t];(null===(e=n.dataset)||void 0===e?void 0:e.generated)&&!this._slotNamesAddedDuringRender.includes(n.slot)&&(this.removeChild(n),delete this._renderedTemplates[n.slot],t--)}this._renderedSlots=!1}}}c([Object(a.b)({attribute:!1}),d("design:type",Object)],l.prototype,"templateContext",void 0),c([Object(a.c)(),d("design:type",Object)],l.prototype,"error",void 0)},Y1eq:function(e,t,n){"use strict";n.d(t,"b",function(){return v}),n.d(t,"a",function(){return y});var a=n("qqMp"),i=n("ZgG/");const r=[a.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{font-size:var(--default-font-size);font-weight:var(--default-font-weight,400);border:var(--file-border,1px solid transparent);border-radius:var(--file-border-radius,4px);background-color:var(--file-background-color)}:host .item{display:flex;flex-flow:row nowrap;align-items:center;background-color:var(--file-background-color);padding:var(--file-padding,0);margin:var(--file-margin,0)}:host .item:hover{background-color:var(--file-background-color-hover,var(--neutral-fill-hover));border-radius:var(--file-border-radius,4px);cursor:pointer;--secondary-text-color:var(--secondary-text-hover-color)}:host .item:focus,:host .item:focus-visible{background-color:var(--file-background-color-focus,var(--neutral-fill-hover));border-radius:var(--file-border-radius,4px)}:host .item__file-type-icon{height:var(--file-type-icon-height,28px);display:flex;align-items:center;justify-content:center}:host .item__file-type-icon img{height:var(--file-type-icon-height,28px)}:host .item__details{padding-inline-start:var(--file-padding-inline-start,14px)}:host .item__details .line1{font-size:var(--default-font-size);font-weight:var(--file-line1-font-weight,var(--default-font-weight,400));text-transform:var(--file-line1-text-transform,initial);line-height:20px;color:var(--file-line1-color,var(--neutral-foreground-rest))}:host .item__details .line2{color:var(--file-line2-color,var(--secondary-text-color));font-size:var(--file-line2-font-size,var(--last-modified-font-size,12px));font-weight:var(--file-line2-font-weight,var(--default-font-weight,400));text-transform:var(--file-line2-text-transform,initial);line-height:16px}:host .item__details .line3{color:var(--file-line3-color,var(--secondary-text-color));font-size:var(--file-line3-font-size,var(--size-font-size,12px));font-weight:var(--file-line3-font-weight,var(--default-font-weight,400));text-transform:var(--file-line3-text-transform,initial);line-height:16px}[dir=rtl] .item{direction:rtl}[dir=rtl] .item__details{direction:rtl}
`];var o=n("Y1A4"),s=n("zFbe"),c=n("/i08"),d=n("0Uyf"),l=n("/4Fm"),u=n("ZzBS");const f={PowerPoint:"pptx",Word:"docx",Excel:"xlsx",Pdf:"pdf",OneNote:"onetoc",OneNotePage:"one",InfoPath:"xsn",Visio:"vstx",Publisher:"pub",Project:"mpp",Access:"accdb",Mail:"email",Csv:"csv",Archive:"archive",Xps:"vector",Audio:"audio",Video:"video",Image:"photo",Web:"html",Text:"txt",Xml:"xml",Story:"genericfile",ExternalContent:"genericfile",Folder:"folder",Spsite:"spo",Other:"genericfile"},p="https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201008.001/assets/item-types";var m=n("7Cdu");const _={modifiedSubtitle:"Modified",sizeSubtitle:"Size"};var h=n("0mOt"),b=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},g=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};const v=()=>Object(h.b)("file",y);class y extends o.a{static get styles(){return r}get strings(){return _}get fileQuery(){return this._fileQuery}set fileQuery(e){e!==this._fileQuery&&(this._fileQuery=e,this.requestStateUpdate())}get siteId(){return this._siteId}set siteId(e){e!==this._siteId&&(this._siteId=e,this.requestStateUpdate())}get driveId(){return this._driveId}set driveId(e){e!==this._driveId&&(this._driveId=e,this.requestStateUpdate())}get groupId(){return this._groupId}set groupId(e){e!==this._groupId&&(this._groupId=e,this.requestStateUpdate())}get listId(){return this._listId}set listId(e){e!==this._listId&&(this._listId=e,this.requestStateUpdate())}get userId(){return this._userId}set userId(e){e!==this._userId&&(this._userId=e,this.requestStateUpdate())}get itemId(){return this._itemId}set itemId(e){e!==this._itemId&&(this._itemId=e,this.requestStateUpdate())}get itemPath(){return this._itemPath}set itemPath(e){e!==this._itemPath&&(this._itemPath=e,this.requestStateUpdate())}get insightType(){return this._insightType}set insightType(e){e!==this._insightType&&(this._insightType=e,this.requestStateUpdate())}get insightId(){return this._insightId}set insightId(e){e!==this._insightId&&(this._insightId=e,this.requestStateUpdate())}get fileDetails(){return this._fileDetails}set fileDetails(e){e!==this._fileDetails&&(this._fileDetails=e,this.requestStateUpdate())}get fileIcon(){return this._fileIcon}set fileIcon(e){e!==this._fileIcon&&(this._fileIcon=e,this.requestStateUpdate())}static get requiredScopes(){return[...new Set(["files.read","files.read.all","sites.read.all"])]}constructor(){super(),this.line1Property="name",this.line2Property="lastModifiedDateTime",this.line3Property="size",this.view=u.a.threelines}render(){if(!this.driveItem&&this.isLoadingState)return this.renderLoading();if(!this.driveItem)return this.renderNoData();const e=this.driveItem;let t;if(t=this.renderTemplate("default",{file:e}),!t){const n=this.renderDetails(e),i=this.renderFileTypeIcon();t=a.c`
        <div class="item">
          ${i} ${n}
        </div>`}return t}renderLoading(){return this.renderTemplate("loading",null)||a.c``}renderNoData(){return this.renderTemplate("no-data",null)||a.c``}renderFileTypeIcon(){if(!this.fileIcon&&!this.driveItem.name)return a.c``;let e;if(this.fileIcon)e=this.fileIcon;else{const n=/(?:\.([^.]+))?$/,a=void 0===this.driveItem.package&&void 0===this.driveItem.folder?n.exec(this.driveItem.name)[1]?n.exec(this.driveItem.name)[1].toLowerCase():this.driveItem.size?"null":"folder":void 0!==this.driveItem.package&&"oneNote"===this.driveItem.package.type?"onetoc":"folder";t=a,e=Object.keys(f).find(e=>f[e]===t)?`${p}/${48..toString()}/${t}.svg`:"jpg"===t||"png"===t?(t="photo",`${p}/${48..toString()}/${t}.svg`):null}var t;return a.c`
      <div class="item__file-type-icon">
        ${e?a.c`
              <img src=${e} alt="File icon" />
            `:a.c`
              ${Object(m.b)(m.a.File)}
            `}
      </div>
    `}renderDetails(e){if(!e||this.view===u.a.image)return a.c``;const t=[];if(this.view>u.a.image){const n=this.getTextFromProperty(e,this.line1Property);n&&t.push(a.c`
          <div class="line1" aria-label="${n}">${n}</div>
        `)}if(this.view>u.a.oneline){const n=this.getTextFromProperty(e,this.line2Property);n&&t.push(a.c`
          <div class="line2" aria-label="${n}">${n}</div>
        `)}if(this.view>u.a.twolines){const n=this.getTextFromProperty(e,this.line3Property);n&&t.push(a.c`
          <div class="line3" aria-label="${n}">${n}</div>
        `)}return a.c`
      <div class="item__details">
        ${t}
      </div>
    `}loadState(){return e=this,void 0,n=function*(){if(this.fileDetails)return void(this.driveItem=this.fileDetails);const e=s.a.globalProvider;if(!e||e.state===c.c.Loading)return;if(e.state===c.c.SignedOut)return void(this.driveItem=null);const t=e.graph.forComponent(this);let n;const a=!(this.driveId||this.siteId||this.groupId||this.listId||this.userId);this.driveId&&!this.itemId&&!this.itemPath||this.siteId&&!this.itemId&&!this.itemPath||this.groupId&&!this.itemId&&!this.itemPath||this.listId&&!this.siteId&&!this.itemId||this.insightType&&!this.insightId||this.userId&&!this.itemId&&!this.itemPath&&!this.insightType&&!this.insightId?n=null:this.fileQuery?n=yield Object(d.i)(t,this.fileQuery):this.itemId&&a?n=yield Object(d.u)(t,this.itemId):this.itemPath&&a?n=yield Object(d.v)(t,this.itemPath):this.userId?this.itemId?n=yield Object(d.D)(t,this.userId,this.itemId):this.itemPath?n=yield Object(d.E)(t,this.userId,this.itemPath):this.insightType&&this.insightId&&(n=yield Object(d.H)(t,this.userId,this.insightType,this.insightId)):this.driveId?this.itemId?n=yield Object(d.g)(t,this.driveId,this.itemId):this.itemPath&&(n=yield Object(d.h)(t,this.driveId,this.itemPath)):this.siteId&&!this.listId?this.itemId?n=yield Object(d.y)(t,this.siteId,this.itemId):this.itemPath&&(n=yield Object(d.z)(t,this.siteId,this.itemPath)):this.listId?n=yield Object(d.t)(t,this.siteId,this.listId,this.itemId):this.groupId?this.itemId?n=yield Object(d.p)(t,this.groupId,this.itemId):this.itemPath&&(n=yield Object(d.q)(t,this.groupId,this.itemPath)):this.insightType&&!this.userId&&(n=yield Object(d.w)(t,this.insightType,this.insightId)),this.driveItem=n},new((t=void 0)||(t=Promise))(function(a,i){function r(e){try{s(n.next(e))}catch(e){i(e)}}function o(e){try{s(n.throw(e))}catch(e){i(e)}}function s(e){var n;e.done?a(e.value):(n=e.value,n instanceof t?n:new t(function(e){e(n)})).then(r,o)}s((n=n.apply(e,[])).next())});var e,t,n}getTextFromProperty(e,t){if(!t||0===t.length)return null;const n=t.trim().split(",");let a,i=0;for(;!a&&i<n.length;){const t=n[i].trim();switch(t){case"size":{let t="0";e.size&&(t=Object(l.d)(e.size)),a=`${this.strings.sizeSubtitle}: ${t}`;break}case"lastModifiedDateTime":{let t,n;if(e.lastModifiedDateTime){const a=new Date(e.lastModifiedDateTime);t=Object(l.l)(a),n=`${this.strings.modifiedSubtitle} ${t}`}else n="";a=n;break}default:a=e[t]}i++}return a}}b([Object(i.b)({attribute:"file-query"}),g("design:type",String),g("design:paramtypes",[String])],y.prototype,"fileQuery",null),b([Object(i.b)({attribute:"site-id"}),g("design:type",String),g("design:paramtypes",[String])],y.prototype,"siteId",null),b([Object(i.b)({attribute:"drive-id"}),g("design:type",String),g("design:paramtypes",[String])],y.prototype,"driveId",null),b([Object(i.b)({attribute:"group-id"}),g("design:type",String),g("design:paramtypes",[String])],y.prototype,"groupId",null),b([Object(i.b)({attribute:"list-id"}),g("design:type",String),g("design:paramtypes",[String])],y.prototype,"listId",null),b([Object(i.b)({attribute:"user-id"}),g("design:type",String),g("design:paramtypes",[String])],y.prototype,"userId",null),b([Object(i.b)({attribute:"item-id"}),g("design:type",String),g("design:paramtypes",[String])],y.prototype,"itemId",null),b([Object(i.b)({attribute:"item-path"}),g("design:type",String),g("design:paramtypes",[String])],y.prototype,"itemPath",null),b([Object(i.b)({attribute:"insight-type"}),g("design:type",String),g("design:paramtypes",[String])],y.prototype,"insightType",null),b([Object(i.b)({attribute:"insight-id"}),g("design:type",String),g("design:paramtypes",[String])],y.prototype,"insightId",null),b([Object(i.b)({type:Object}),g("design:type",Object),g("design:paramtypes",[Object])],y.prototype,"fileDetails",null),b([Object(i.b)({attribute:"file-icon"}),g("design:type",String),g("design:paramtypes",[String])],y.prototype,"fileIcon",null),b([Object(i.b)({type:Object}),g("design:type",Object)],y.prototype,"driveItem",void 0),b([Object(i.b)({attribute:"line1-property"}),g("design:type",String)],y.prototype,"line1Property",void 0),b([Object(i.b)({attribute:"line2-property"}),g("design:type",String)],y.prototype,"line2Property",void 0),b([Object(i.b)({attribute:"line3-property"}),g("design:type",String)],y.prototype,"line3Property",void 0),b([Object(i.b)({attribute:"view",converter:e=>e&&0!==e.length?(e=e.toLowerCase(),void 0===u.a[e]?u.a.threelines:u.a[e]):u.a.threelines}),g("design:type",Number)],y.prototype,"view",void 0)},YBl9:function(e,t,n){"use strict";n.d(t,"a",function(){return o});var a=n("6eT7"),i=n("0q6d");const r=/^#((?:[0-9a-f]{6}|[0-9a-f]{3}))$/i;function o(e){const t=r.exec(e);if(null===t)return null;let n=t[1];if(3===n.length){const e=n.charAt(0),t=n.charAt(1),a=n.charAt(2);n=e.concat(e,t,t,a,a)}const o=parseInt(n,16);return isNaN(o)?null:new a.a(Object(i.g)((16711680&o)>>>16,0,255),Object(i.g)((65280&o)>>>8,0,255),Object(i.g)(255&o,0,255),1)}},YZrM:function(e,t,n){"use strict";n.d(t,"a",function(){return a});class a{}a.sections={files:!0,mailMessages:!0,organization:{showWorksWith:!0},profile:!0},a.useContactApis=!0,a.isSendMessageVisible=!0},Yn9m:function(e,t,n){"use strict";n.d(t,"a",function(){return a});class a{static get microsoftTeamsLib(){return this._microsoftTeamsLib||window.microsoftTeams}static set microsoftTeamsLib(e){this._microsoftTeamsLib=e}static get isAvailable(){return!(!this.microsoftTeamsLib||(window.parent!==window.self||!window.nativeInterface)&&"embedded-page-container"!==window.name&&"extension-tab-frame"!==window.name)}static executeDeepLink(e,t){const n=this.microsoftTeamsLib;n.initialize(),n.executeDeepLink(e,t)}}},Ys1X:function(e,t,n){"use strict";n.d(t,"b",function(){return f}),n.d(t,"a",function(){return p});var a=n("qqMp"),i=n("C001"),r=n("7Cdu");const o=[a.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host .loading,:host .no-data{margin:0 20px;display:flex;justify-content:center}:host .no-data{font-style:normal;font-weight:600;font-size:14px;color:var(--font-color,#323130);line-height:19px}:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{position:relative;user-select:none}:host .root.compact{padding:0}:host .root.compact .coworker .coworker__image{height:40px;width:40px;border-radius:40px;--avatar-size:40px;--avatar-size-s:40px;margin-right:12px}:host .root.compact .coworker .coworker__name{font-size:14px}:host .root.compact .coworker .coworker__title{font-size:12px}:host .root .subtitle{color:var(--organization-sub-title-color,var(--neutral-foreground-hint));font-size:14px;margin:0 20px 8px}:host .root .divider{display:flex;background:var(--organization-coworker-border-color,var(--neutral-stroke-rest));height:1px;margin:26px 20px 18px}:host .root .org-member{height:74px;box-sizing:border-box;border-radius:2px;padding:12px;display:flex;align-items:center;margin-left:20px;margin-right:20px}:host .root .org-member.org-member--target{background-color:var(--organization-active-org-member-target-background-color,var(--neutral-fill-active));border:1px solid var(--organization-active-org-member-border-color,var(--accent-foreground-rest))}:host .root .org-member:not(.org-member--target){border:1px solid var(--organization-coworker-border-color,var(--neutral-stroke-rest))}:host .root .org-member:not(.org-member--target):hover{cursor:pointer;background-color:var(--organization-hover-color,var(--neutral-fill-hover))}:host .root .org-member .org-member__person{flex-grow:1}:host .root .org-member .org-member__details{flex-grow:1}:host .root .org-member .org-member__details .org-member__name{font-size:16px;color:var(--organization-title-color,var(--neutral-foreground-rest));font-weight:600}:host .root .org-member .org-member__details .org-member__department,:host .root .org-member .org-member__details .org-member__title{font-weight:14px;color:var(--organization-sub-title-color,var(--neutral-foreground-hint))}:host .root .org-member__separator:not(:last-child){border:1px solid var(--organization-coworker-border-color,var(--neutral-stroke-rest));box-sizing:border-box;width:0;margin:0 50%;height:14px}:host .root .coworker{display:flex;align-items:center;padding:10px 20px}:host .root .coworker:hover{cursor:pointer;background-color:var(--organization-coworker-hover-color,var(--neutral-fill-hover))}:host .root .coworker .coworker__person{height:46px;border-radius:46px;margin-right:8px}:host .root .direct-report__compact{padding:12px 20px}:host .root .direct-report__compact .direct-report{cursor:pointer;width:38px;margin-right:4px;display:inline;--avatar-size:38px}[dir=rtl] .org-member .org-member__more{transform:scaleX(-1);filter:fliph;filter:FlipH}@media (forced-colors:active) and (prefers-color-scheme:dark){:host svg,:host svg>path{fill:#fff!important;fill-rule:nonzero!important;clip-rule:nonzero!important}}@media (forced-colors:active) and (prefers-color-scheme:light){:host svg,:host svg>path{fill:#000!important;fill-rule:nonzero!important;clip-rule:nonzero!important}}
`],s={reportsToSectionTitle:"Reports to",directReportsSectionTitle:"Direct reports",organizationSectionTitle:"Organization",youWorkWithSubSectionTitle:"You work with",userWorksWithSubSectionTitle:"works with"};var c=n("ZzBS"),d=n("cBsD"),l=n("0mOt"),u=n("7SX6");const f=()=>{Object(u.c)(),Object(l.b)("organization",p)};class p extends i.a{static get styles(){return o}get strings(){return s}constructor(e,t){super(),this._state=e,this._me=t}clearState(){super.clearState(),this._state=null,this._me=null}get displayName(){const{person:e,directReports:t}=this._state;return!e.manager&&(null==t?void 0:t.length)?`${this.strings.directReportsSectionTitle} (${t.length})`:this.strings.reportsToSectionTitle}get cardTitle(){return this.strings.organizationSectionTitle}renderIcon(){return Object(r.b)(r.a.Organization)}renderCompactView(){var e;let t;if(!(null===(e=this._state)||void 0===e?void 0:e.person))return null;const{person:n,directReports:i}=this._state;return n?(n.manager?t=this.renderCoworker(n.manager):(null==i?void 0:i.length)&&(t=this.renderCompactDirectReports()),a.c`
        <div class="root compact">
          ${t}
        </div>
      `):null}renderFullView(){var e;let t;if(!(null===(e=this._state)||void 0===e?void 0:e.person))return null;const{person:n,directReports:i,people:r}=this._state;if(!(n||i||r))return null;{const e=this.renderManagers(),n=this.renderCurrentUser(),i=this.renderDirectReports(),r=this.renderCoworkers();t=a.c`
          ${e} ${n} ${i} ${r}
        `}return a.c`
       <div class="root" dir=${this.direction}>
         ${t}
       </div>
     `}renderManager(e){return d.a`
      <div
        class="org-member"
        @keydown=${t=>{"Enter"!==t.code&&" "!==t.code||this.navigateCard(e)}}
        @click=${()=>this.navigateCard(e)}
      >
        <div class="org-member__person">
          <mgt-person
            .personDetails=${e}
            .fetchImage=${!0}
            .view=${c.a.twolines}
            .showPresence=${!0}
          ></mgt-person>
        </div>
        <div tabindex="0" class="org-member__more">
          ${Object(r.b)(r.a.ExpandRight)}
        </div>
      </div>
      <div class="org-member__separator"></div>
     `}renderManagers(){const{person:e}=this._state;if(!(null==e?void 0:e.manager))return null;const t=[];let n=e;for(;n.manager;)t.push(n.manager),n=n.manager;return t.length?t.reverse().map(e=>this.renderManager(e)):null}renderDirectReports(){const{directReports:e}=this._state;return(null==e?void 0:e.length)?a.c`
      <div class="org-member__separator"></div>
      <div>
        ${e.map(e=>d.a`
            <div
              class="org-member org-member--direct-report"
              @keydown=${t=>{"Enter"!==t.code&&" "!==t.code||this.navigateCard(e)}}
              @click=${()=>this.navigateCard(e)}
            >
              <div class="org-member__person">
                <mgt-person
                  .personDetails=${e}
                  .fetchImage=${!0}
                  .showPresence=${!0}
                  .view=${c.a.twolines}
                ></mgt-person>
              </div>
              <div tabindex="0" class="org-member__more">
                ${Object(r.b)(r.a.ExpandRight)}
              </div>
            </div>
            <div class="org-member__separator"></div>
          `)}
      </div>
     `:null}renderCompactDirectReports(){const{directReports:e}=this._state;return a.c`
      <div class="direct-report__compact">
        ${e.slice(0,6).map(e=>d.a`
            <div
              class="direct-report"
              @keydown=${t=>{"Enter"!==t.code&&" "!==t.code||this.navigateCard(e)}}
              @click=${()=>this.navigateCard(e)}
            >
              <mgt-person
                .personDetails=${e}
                .fetchImage=${!0}
                .showPresence=${!0}
                .view=${c.a.twolines}
              ></mgt-person>
            </div>
          `)}
      </div>
    `}renderCurrentUser(){const{person:e}=this._state;return d.a`
       <div class="org-member org-member--target">
         <div class="org-member__person">
           <mgt-person
             .personDetails=${e}
             .fetchImage=${!0}
             .showPresence=${!0}
             .view=${c.a.twolines}
           ></mgt-person>
         </div>
       </div>
     `}renderCoworker(e){return d.a`
      <div
        class="coworker"
        @keydown=${t=>{"Enter"!==t.code&&" "!==t.code||this.navigateCard(e)}}
        @click=${()=>this.navigateCard(e)}
      >
        <div class="coworker__person">
          <mgt-person
            .personDetails=${e}
            .fetchImage=${!0}
            .showPresence=${!0}
            .view=${c.a.twolines}
          ></mgt-person>
        </div>
      </div>
    `}renderCoworkers(){const{people:e}=this._state;if(!(null==e?void 0:e.length))return null;const t=this._me.id===this._state.person.id?this.strings.youWorkWithSubSectionTitle:`${this._state.person.givenName} ${this.strings.userWorksWithSubSectionTitle}`;return a.c`
       <div class="divider"></div>
       <div class="subtitle" tabindex="0">${t}</div>
       <div>
         ${e.slice(0,6).map(e=>this.renderCoworker(e))}
       </div>
     `}}},"ZgG/":function(e,t,n){"use strict";n.d(t,"a",function(){return a}),n.d(t,"b",function(){return s}),n.d(t,"c",function(){return c});const a=e=>(t,n)=>{void 0!==n?n.addInitializer(()=>{customElements.define(e,t)}):customElements.define(e,t)};var i=n("ptXv");const r={attribute:!0,type:String,converter:i.c,reflect:!1,hasChanged:i.d},o=(e=r,t,n)=>{const{kind:a,metadata:i}=n;let o=globalThis.litPropertyMetadata.get(i);if(void 0===o&&globalThis.litPropertyMetadata.set(i,o=new Map),o.set(n.name,e),"accessor"===a){const{name:a}=n;return{set(n){const i=t.get.call(this);t.set.call(this,n),this.requestUpdate(a,i,e)},init(t){return void 0!==t&&this.C(a,void 0,e),t}}}if("setter"===a){const{name:a}=n;return function(n){const i=this[a];t.call(this,n),this.requestUpdate(a,i,e)}}throw Error("Unsupported decorator location: "+a)};function s(e){return(t,n)=>"object"==typeof n?o(e,t,n):((e,t,n)=>{const a=t.hasOwnProperty(n);return t.constructor.createProperty(n,a?{...e,wrapped:!0}:e),a?Object.getOwnPropertyDescriptor(t,n):void 0})(e,t,n)}function c(e){return s({...e,state:!0,attribute:!1})}},ZzBS:function(e,t,n){"use strict";var a;n.d(t,"a",function(){return a}),function(e){e[e.image=2]="image",e[e.oneline=3]="oneline",e[e.twolines=4]="twolines",e[e.threelines=5]="threelines",e[e.fourlines=6]="fourlines"}(a||(a={}))},bica:function(e,t,n){"use strict";n.d(t,"a",function(){return y});var a=n("3rEL"),i=n("TDEi");class r extends i.a{}var o=n("+Cud"),s=n("6BDD"),c=n("oePG"),d=n("QBS5"),l=n("YBl9"),u=n("8hiW"),f=n("i2HT"),p=n("hv+n"),m=n("4X57"),_=n("xY0q"),h=n("oqLQ"),b=n("j9Xn"),g=n("cQsl");class v extends r{cardFillColorChanged(e,t){if(t){const e=Object(l.a)(t);null!==e&&(this.neutralPaletteSource=t,u.v.setValueFor(this,f.a.create(e.r,e.g,e.b)))}}neutralPaletteSourceChanged(e,t){if(t){const e=Object(l.a)(t),n=f.a.create(e.r,e.g,e.b);u.hb.setValueFor(this,p.a.create(n))}}handleChange(e,t){this.cardFillColor||u.v.setValueFor(this,t=>u.P.getValueFor(t).evaluate(t,u.v.getValueFor(e)).rest)}connectedCallback(){super.connectedCallback();const e=Object(o.a)(this);if(e){const t=c.b.getNotifier(e);t.subscribe(this,"fillColor"),t.subscribe(this,"neutralPalette"),this.handleChange(e,"fillColor")}}}Object(a.a)([Object(d.c)({attribute:"card-fill-color",mode:"fromView"})],v.prototype,"cardFillColor",void 0),Object(a.a)([Object(d.c)({attribute:"neutral-palette-source",mode:"fromView"})],v.prototype,"neutralPaletteSource",void 0);const y=v.compose({baseName:"card",baseClass:r,template:(e,t)=>s.a`
    <slot></slot>
`,styles:(e,t)=>m.a`
    ${Object(_.a)("block")} :host {
      display: block;
      contain: content;
      height: var(--card-height, 100%);
      width: var(--card-width, 100%);
      box-sizing: border-box;
      background: ${u.v};
      color: ${u.fb};
      border: calc(${u.vb} * 1px) solid ${u.qb};
      border-radius: calc(${u.D} * 1px);
      box-shadow: ${g.a};
    }

    :host {
      content-visibility: auto;
    }
  `.withBehaviors(Object(h.a)(m.a`
        :host {
          background: ${b.a.Canvas};
          color: ${b.a.CanvasText};
        }
      `))})},c1DA:function(e,t,n){"use strict";n.d(t,"a",function(){return f});var a=n("qqMp"),i=n("ZgG/"),r=n("h2QR");const o=()=>void 0!==window.getWindowSegments,s=[a.b`
.root .scout-top{position:fixed;top:0;left:0}.root .scout-bottom{position:fixed;bottom:0;right:0}.root .flyout{display:none;position:fixed;z-index:9999999;opacity:0;box-shadow:var(--mgt-flyout-box-shadow,var(--elevation-shadow-card-rest));border-radius:8px}.root .flyout.small{overflow:hidden auto}.root.visible .flyout{display:inline-block;animation-name:fade-in;animation-duration:.3s;animation-timing-function:cubic-bezier(.1,.9,.2,1);animation-fill-mode:both;transition:top .3s ease,bottom .3s ease,left .3s ease;z-index:9999999}.root.visible .flyout.small{overflow:hidden auto}@keyframes fade-in{from{opacity:0;margin-top:-10px;margin-bottom:-10px}to{opacity:1;margin-top:0;margin-bottom:0}}
`];var c=n("dblZ"),d=n("0mOt"),l=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},u=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};const f=()=>Object(d.b)("flyout",p);class p extends c.b{static get styles(){return s}get isOpen(){return this._isOpen}set isOpen(e){const t=this._isOpen;t!==e&&(this._isOpen=e,window.requestAnimationFrame(()=>{this.setupWindowEvents(this.isOpen);const e=this._flyout;!this.isOpen&&e&&(e.style.width=null,e.style.setProperty("--mgt-flyout-set-width",null),e.style.setProperty("--mgt-flyout-set-height",null),e.style.maxHeight=null,e.style.top=null,e.style.left=null,e.style.bottom=null)}),this.requestUpdate("isOpen",t),this.dispatchEvent(new Event(e?"opened":"closed")))}get _edgePadding(){return 24}get _flyout(){return this.renderRoot.querySelector(".flyout")}get _anchor(){return this.renderRoot.querySelector(".anchor")}get _topScout(){return this.renderRoot.querySelector(".scout-top")}get _bottomScout(){return this.renderRoot.querySelector(".scout-bottom")}constructor(){super(),this._renderedOnce=!1,this._isOpen=!1,this._smallView=!1,this.handleWindowEvent=e=>{const t=this._flyout;if(t)if(e.composedPath){const n=e.composedPath();if(n.includes(t)||"pointerdown"===e.type&&n.includes(this))return}else{let n=e.target;for(;n;)if(n=n.parentElement,n===t||"pointerdown"===e.type&&n===this)return}this.close()},this.handleResize=()=>{this.close()},this.handleKeyUp=e=>{"Escape"===e.key&&this.close()},this.handleFlyoutWheel=e=>{e.preventDefault()},this.avoidHidingAnchor=!0,this.addEventListener("expanded",()=>{window.requestAnimationFrame(()=>{this.updateFlyout()})})}open(){this.isOpen=!0}close(){this.isOpen=!1}disconnectedCallback(){this.setupWindowEvents(!1),super.disconnectedCallback()}updated(e){super.updated(e),window.requestAnimationFrame(()=>{this.updateFlyout()})}render(){const e={root:!0,visible:this.isOpen},t=this.renderAnchor();let n=null;if(this._windowHeight=window.innerHeight&&document.documentElement.clientHeight?Math.min(window.innerHeight,document.documentElement.clientHeight):window.innerHeight||document.documentElement.clientHeight,this._windowHeight<250&&(this._smallView=!0),this.isOpen||this._renderedOnce){this._renderedOnce=!0;const e=Object(r.a)({flyout:!0,small:this._smallView});n=a.c`
        <div class=${e} @wheel=${this.handleFlyoutWheel}>
          ${this.renderFlyout()}
        </div>
      `}return a.c`
      <div class=${Object(r.a)(e)}>
        <div class="anchor">
          ${t}
        </div>
        <div class="scout-top"></div>
        <div class="scout-bottom"></div>
        ${n}
      </div>
    `}renderAnchor(){return a.c`
      <slot name="anchor"></slot>
    `}renderFlyout(){return a.c`
      <slot name="flyout"></slot>
    `}updateFlyout(e=!0){if(!this.isOpen)return;const t=this._anchor,n=this._flyout;if(n&&t){const a=window.innerWidth&&document.documentElement.clientWidth?Math.min(window.innerWidth,document.documentElement.clientWidth):window.innerWidth||document.documentElement.clientWidth;this._windowHeight=window.innerHeight&&document.documentElement.clientHeight?Math.min(window.innerHeight,document.documentElement.clientHeight):window.innerHeight||document.documentElement.clientHeight;let i,r,s,c=0,d=0;const l=n.getBoundingClientRect(),u=t.getBoundingClientRect(),f=this._topScout.getBoundingClientRect(),p=this._bottomScout.getBoundingClientRect(),m={height:this._windowHeight,left:0,top:0,width:a};if(o()){const e=(o()?window:null).getWindowSegments();let t;const n=u.left+u.width/2,a=u.top+u.height/2;for(const i of e)if(n>=i.left&&a>=i.top){t=i;break}t&&(m.left=t.left,m.top=t.top,m.width=t.width,m.height=t.height)}l.width+2*this._edgePadding>m.width?l.width>m.width?(s=m.width,c=0):c=(m.width-l.width)/2:c=u.left+l.width+this._edgePadding>m.width?u.left-(u.left+l.width+this._edgePadding-m.width):u.left<this._edgePadding?this._edgePadding:u.left;const _=m.height-(u.top+u.height),h=u.top;this.avoidHidingAnchor?_<=l.height?h<l.height?h>_?(i=m.height-u.top,r=h):(d=u.bottom,r=_):(i=m.height-u.top,r=h):(d=u.bottom,r=_):l.height+2*this._edgePadding>m.height?l.height>=m.height?(r=m.height,d=0):d=(m.height-l.height)/2:d=u.top+u.height+l.height+this._edgePadding>m.height?m.height-l.height-this._edgePadding:Math.max(u.top+u.height,this._edgePadding),0===f.top&&0===f.left||(c-=f.left,void 0!==i?i+=p.top-this._windowHeight:d-=f.top),"rtl"===this.direction?c>100&&this.offsetLeft>100&&(n.style.left=m.width-c+l.left-l.width-30+"px"):n.style.left=`${c+m.left}px`,void 0!==i?(n.style.top="unset",n.style.bottom=`${i}px`):(n.style.bottom="unset",n.style.top=`${d+m.top}px`),s&&(n.style.width=`${s}px`,n.style.setProperty("--mgt-flyout-set-width",`${s}px`),window.requestAnimationFrame(()=>this.updateFlyout())),r&&!e?(n.style.maxHeight=`${r}px`,n.style.setProperty("--mgt-flyout-set-height",`${r}px`)):(n.style.maxHeight=null,n.style.setProperty("--mgt-flyout-set-height","unset")),e&&window.requestAnimationFrame(()=>this.updateFlyout(!1))}}setupWindowEvents(e){e&&this.isLightDismiss?(window.addEventListener("wheel",this.handleWindowEvent),window.addEventListener("pointerdown",this.handleWindowEvent),window.addEventListener("resize",this.handleResize),window.addEventListener("keyup",this.handleKeyUp)):(window.removeEventListener("wheel",this.handleWindowEvent),window.removeEventListener("pointerdown",this.handleWindowEvent),window.removeEventListener("resize",this.handleResize),window.removeEventListener("keyup",this.handleKeyUp))}}l([Object(i.b)({attribute:"light-dismiss",type:Boolean}),u("design:type",Boolean)],p.prototype,"isLightDismiss",void 0),l([Object(i.b)({attribute:null,type:Boolean}),u("design:type",Boolean)],p.prototype,"avoidHidingAnchor",void 0),l([Object(i.b)({attribute:"isOpen",type:Boolean}),u("design:type",Boolean),u("design:paramtypes",[Boolean])],p.prototype,"isOpen",null)},cBsD:function(e,t,n){"use strict";n.d(t,"a",function(){return s});var a=n("qqMp"),i=n("HN6m");const r=new WeakMap,o=(e,t,n)=>{const a=[];for(const i of e)a.push(i.replace(t,n));return a},s=(e,...t)=>{if(i.a.isDisambiguated){let t=r.get(e);if(!t){const n=new RegExp("(</?)mgt-(?!"+i.a.disambiguation+"-)"),a=`$1${i.a.prefix}-`;t=Object.assign(o(e,n,a),{raw:o(e.raw,n,a)}),r.set(e,t)}e=t}return Object(a.c)(e,...t)}},cQsl:function(e,t,n){"use strict";n.d(t,"a",function(){return d}),n.d(t,"d",function(){return u}),n.d(t,"c",function(){return p}),n.d(t,"b",function(){return _});var a=n("8GQ4");const i=a.a.create({name:"elevation-shadow",cssCustomPropertyName:null}).withDefault({evaluate:(e,t,n)=>{let a=.12,i=.14;return t>16&&(a=.2,i=.24),`0 0 2px rgba(0, 0, 0, ${a}), 0 calc(${t} * 0.5px) calc((${t} * 1px)) rgba(0, 0, 0, ${i})`}}),r=a.a.create("elevation-shadow-card-rest-size").withDefault(4),o=a.a.create("elevation-shadow-card-hover-size").withDefault(8),s=a.a.create("elevation-shadow-card-active-size").withDefault(0),c=a.a.create("elevation-shadow-card-focus-size").withDefault(8),d=a.a.create("elevation-shadow-card-rest").withDefault(e=>i.getValueFor(e).evaluate(e,r.getValueFor(e))),l=(a.a.create("elevation-shadow-card-hover").withDefault(e=>i.getValueFor(e).evaluate(e,o.getValueFor(e))),a.a.create("elevation-shadow-card-active").withDefault(e=>i.getValueFor(e).evaluate(e,s.getValueFor(e))),a.a.create("elevation-shadow-card-focus").withDefault(e=>i.getValueFor(e).evaluate(e,c.getValueFor(e))),a.a.create("elevation-shadow-tooltip-size").withDefault(16)),u=a.a.create("elevation-shadow-tooltip").withDefault(e=>i.getValueFor(e).evaluate(e,l.getValueFor(e))),f=a.a.create("elevation-shadow-flyout-size").withDefault(32),p=a.a.create("elevation-shadow-flyout").withDefault(e=>i.getValueFor(e).evaluate(e,f.getValueFor(e))),m=a.a.create("elevation-shadow-dialog-size").withDefault(128),_=a.a.create("elevation-shadow-dialog").withDefault(e=>i.getValueFor(e).evaluate(e,m.getValueFor(e)))},d3jT:function(e,t,n){"use strict";n.d(t,"a",function(){return D});const a=(e,t)=>t.some(t=>e instanceof t);let i,r;const o=new WeakMap,s=new WeakMap,c=new WeakMap,d=new WeakMap,l=new WeakMap;let u={get(e,t,n){if(e instanceof IDBTransaction){if("done"===t)return s.get(e);if("objectStoreNames"===t)return e.objectStoreNames||c.get(e);if("store"===t)return n.objectStoreNames[1]?void 0:n.objectStore(n.objectStoreNames[0])}return f(e[t])},set:(e,t,n)=>(e[t]=n,!0),has:(e,t)=>e instanceof IDBTransaction&&("done"===t||"store"===t)||t in e};function f(e){if(e instanceof IDBRequest)return function(e){const t=new Promise((t,n)=>{const a=()=>{e.removeEventListener("success",i),e.removeEventListener("error",r)},i=()=>{t(f(e.result)),a()},r=()=>{n(e.error),a()};e.addEventListener("success",i),e.addEventListener("error",r)});return t.then(t=>{t instanceof IDBCursor&&o.set(t,e)}).catch(()=>{}),l.set(t,e),t}(e);if(d.has(e))return d.get(e);const t=function(e){return"function"==typeof e?(t=e)!==IDBDatabase.prototype.transaction||"objectStoreNames"in IDBTransaction.prototype?(r||(r=[IDBCursor.prototype.advance,IDBCursor.prototype.continue,IDBCursor.prototype.continuePrimaryKey])).includes(t)?function(...e){return t.apply(p(this),e),f(o.get(this))}:function(...e){return f(t.apply(p(this),e))}:function(e,...n){const a=t.call(p(this),e,...n);return c.set(a,e.sort?e.sort():[e]),f(a)}:(e instanceof IDBTransaction&&function(e){if(s.has(e))return;const t=new Promise((t,n)=>{const a=()=>{e.removeEventListener("complete",i),e.removeEventListener("error",r),e.removeEventListener("abort",r)},i=()=>{t(),a()},r=()=>{n(e.error||new DOMException("AbortError","AbortError")),a()};e.addEventListener("complete",i),e.addEventListener("error",r),e.addEventListener("abort",r)});s.set(e,t)}(e),a(e,i||(i=[IDBDatabase,IDBObjectStore,IDBIndex,IDBCursor,IDBTransaction]))?new Proxy(e,u):e);var t}(e);return t!==e&&(d.set(e,t),l.set(t,e)),t}const p=e=>l.get(e),m=["get","getKey","getAll","getAllKeys","count"],_=["put","add","delete","clear"],h=new Map;function b(e,t){if(!(e instanceof IDBDatabase)||t in e||"string"!=typeof t)return;if(h.get(t))return h.get(t);const n=t.replace(/FromIndex$/,""),a=t!==n,i=_.includes(n);if(!(n in(a?IDBIndex:IDBObjectStore).prototype)||!i&&!m.includes(n))return;const r=async function(e,...t){const r=this.transaction(e,i?"readwrite":"readonly");let o=r.store;return a&&(o=o.index(t.shift())),(await Promise.all([o[n](...t),i&&r.done]))[0]};return h.set(t,r),r}var g;g=u,u={...g,get:(e,t,n)=>b(e,t)||g.get(e,t,n),has:(e,t)=>!!b(e,t)||g.has(e,t)};var v=n("zFbe"),y=n("DNu6"),S=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};class D{constructor(e,t){if(!(t in e.stores))throw Error('"store" must be defined in the "schema"');this.schema=e,this.store=t}getValue(e){return S(this,void 0,void 0,function*(){if(!window.indexedDB)return null;try{return(yield this.getDb()).get(this.store,e)}catch(e){return null}})}delete(e){return S(this,void 0,void 0,function*(){if(window.indexedDB)try{return(yield this.getDb()).delete(this.store,e)}catch(e){return}})}putValue(e,t){return S(this,void 0,void 0,function*(){if(window.indexedDB)try{yield(yield this.getDb()).put(this.store,Object.assign(Object.assign({},t),{timeCached:Date.now()}),e)}catch(e){return}})}clearStore(){return S(this,void 0,void 0,function*(){if(window.indexedDB)try{yield(yield this.getDb()).clear(this.store)}catch(e){return}})}getDBName(){return S(this,void 0,void 0,function*(){const e=yield v.a.getCacheId();if(e)return`mgt-${this.schema.name}-${e}`})}getDb(){return S(this,void 0,void 0,function*(){const e=yield this.getDBName();if(e)return function(e,t,{blocked:n,upgrade:a,blocking:i,terminated:r}={}){const o=indexedDB.open(e,t),s=f(o);return a&&o.addEventListener("upgradeneeded",e=>{a(f(o.result),e.oldVersion,e.newVersion,f(o.transaction))}),n&&o.addEventListener("blocked",()=>n()),s.then(e=>{r&&e.addEventListener("close",()=>r()),i&&e.addEventListener("versionchange",()=>i())}).catch(()=>{}),s}(e,this.schema.version,{upgrade:(t,n,a,i)=>{var r,o;const s=JSON.parse(localStorage.getItem(y.b))||[];s.includes(e)||s.push(e),localStorage.setItem(y.b,JSON.stringify(s));for(const e in this.schema.stores)if(Object.prototype.hasOwnProperty.call(this.schema.stores,e)){const n=null!==(o=null===(r=this.schema.indexes)||void 0===r?void 0:r[e])&&void 0!==o?o:[];if(t.objectStoreNames.contains(e)){const t=i.objectStore(e);n.forEach(e=>{t&&!t.indexNames.contains(e.name)&&t.createIndex(e.name,e.field)})}else{const a=t.createObjectStore(e);n.forEach(e=>{a.createIndex(e.name,e.field)})}}}})})}queryDb(e,t){return S(this,void 0,void 0,function*(){const n=yield this.getDb();return yield n.getAllFromIndex(this.store,e,t)})}transaction(e){return S(this,void 0,void 0,function*(){const t=(yield this.getDb()).transaction(this.store,"readwrite"),n=t.objectStore(this.store);yield e(n),yield t.done})}}},dP6N:function(e,t,n){"use strict";var a,i;n.d(t,"a",function(){return a}),n.d(t,"b",function(){return i}),function(e){e[e.avatar=2]="avatar",e[e.oneline=3]="oneline",e[e.twolines=4]="twolines",e[e.threelines=5]="threelines",e[e.fourlines=6]="fourlines"}(a||(a={})),function(e){e.photo="photo",e.initials="initials"}(i||(i={}))},dblZ:function(e,t,n){"use strict";n.d(t,"a",function(){return a}),n.d(t,"b",function(){return u});var a,i=n("qqMp"),r=n("ZgG/"),o=n("/i08"),s=n("zFbe"),c=n("zb6c"),d=n("2oY9"),l=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};!function(e){e.mobile="",e.tablet="tablet",e.desktop="desktop"}(a||(a={}));class u extends i.a{static get packageVersion(){return d.a}get mediaQuery(){return this.offsetWidth<768?a.mobile:this.offsetWidth<1200?a.tablet:a.desktop}get isLoadingState(){return this._isLoadingState}get isFirstUpdated(){return this._isFirstUpdated}get strings(){return{}}constructor(){super(),this.direction="ltr",this._isLoadingState=!1,this._isFirstUpdated=!1,this.setLoadingState=e=>{this._isLoadingState!==e&&(this._isLoadingState=e,this.requestUpdate("isLoadingState"))},this.handleProviderUpdates=()=>{this.requestStateUpdate()},this.handleActiveAccountUpdates=()=>{this.clearState(),this.requestStateUpdate()},this.handleLocalizationChanged=()=>{c.a.updateStringsForTag(this.tagName,this.strings),this.requestUpdate()},this.handleDirectionChanged=()=>{this.direction=c.a.getDocumentDirection()},this.handleDirectionChanged(),this.handleLocalizationChanged()}connectedCallback(){super.connectedCallback(),c.a.onStringsUpdated(this.handleLocalizationChanged),c.a.onDirectionUpdated(this.handleDirectionChanged)}disconnectedCallback(){super.disconnectedCallback(),c.a.removeOnStringsUpdated(this.handleLocalizationChanged),c.a.removeOnDirectionUpdated(this.handleDirectionChanged),s.a.removeProviderUpdatedListener(this.handleProviderUpdates),s.a.removeActiveAccountChangedListener(this.handleActiveAccountUpdates)}firstUpdated(e){super.firstUpdated(e),this._isFirstUpdated=!0,s.a.onProviderUpdated(this.handleProviderUpdates),s.a.onActiveAccountChanged(this.handleActiveAccountUpdates),this.requestStateUpdate()}loadState(){return Promise.resolve()}clearState(){}fireCustomEvent(e,t,n=!1,a=!1,i=!1){const r=new CustomEvent(e,{bubbles:n,cancelable:a,composed:i,detail:t});return this.dispatchEvent(r)}updated(e){super.updated(e);const t=new CustomEvent("updated",{bubbles:!0,cancelable:!0});this.dispatchEvent(t)}requestStateUpdate(e=!1){return l(this,void 0,void 0,function*(){if(!this._isFirstUpdated)return;this.isLoadingState&&!e&&(yield this._currentLoadStatePromise);const t=s.a.globalProvider;if(!t)return Promise.resolve();if(t.state!==o.c.SignedOut){if(t.state===o.c.Loading)return Promise.resolve();{const t=new Promise((n,a)=>l(this,void 0,void 0,function*(){try{this.setLoadingState(!0),this.fireCustomEvent("loadingInitiated"),yield this.loadState(),this.setLoadingState(!1),this.fireCustomEvent("loadingCompleted"),n()}catch(e){this.clearState(),this.setLoadingState(!1),this.fireCustomEvent("loadingFailed"),a(e)}return this._currentLoadStatePromise=this.isLoadingState&&this._currentLoadStatePromise&&e?this._currentLoadStatePromise.then(()=>t):t}))}}else this.clearState()})}}!function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);r>3&&o&&Object.defineProperty(t,n,o)}([Object(r.c)(),function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata("design:type",t)}(0,String)],u.prototype,"direction",void 0)},dupE:function(e,t,n){"use strict";n.d(t,"b",function(){return de}),n.d(t,"a",function(){return le});var a=n("Y1A4"),i=n("ylMe"),r=n("zFbe"),o=n("cBsD"),s=n("/i08"),c=n("qqMp"),d=n("ZgG/"),l=n("V07F"),u=n("0Uyf"),f=n("7Cdu"),p=n("ZzBS");const m=[c.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{font-size:var(--default-font-size)}:host .title{font-size:14px;font-weight:600;padding:20px 0 12px;line-height:19px}:host .file-list-wrapper{background-color:var(--file-list-background-color,var(--neutral-layer-floating));border:var(--file-list-border,none);position:relative;display:flex;flex-direction:column;border-radius:8px}:host .file-list-wrapper .title{font-size:14px;font-weight:600;margin:0 20px -15px}:host .file-list-wrapper .file-list{display:flex;padding:var(--file-list-padding,0);margin:var(--file-list-margin,0);flex-direction:column;list-style:none}:host .file-list-wrapper .file-list .file-item{cursor:pointer;border-radius:var(--file-border-radius)}:host .file-list-wrapper .file-list .file-item:focus,:host .file-list-wrapper .file-list .file-item:focus-within{--file-background-color:var(--file-background-color-focus, var(--neutral-layer-2))}:host .file-list-wrapper .file-list .file-item.selected{--file-background-color:var(--file-background-color-active, var(--neutral-layer-3))}:host .file-list-wrapper .file-list .file-item .mgt-file-item{--file-padding:10px 20px 10px 20px;--file-padding-inline-start:24px;--file-border-radius:2px;--file-background-color:var(--file-item-background-color, var(--neutral-layer-1))}:host .file-list-wrapper .progress-ring{margin:4px auto;width:var(--progress-ring-size,24px);height:var(--progress-ring-size,24px)}:host .file-list-wrapper .show-more{text-align:center;font-size:var(--show-more-button-font-size,12px);padding:var(--show-more-button-padding,0);border-radius:0 0 var(--show-more-button-border-bottom-right-radius,var(--file-list-border-radius,8px)) var(--show-more-button-border-bottom-left-radius,var(--file-list-border-radius,8px));background-color:var(--show-more-button-background-color,var(--neutral-stroke-divider-rest))}:host .file-list-wrapper .show-more:hover{background-color:var(--show-more-button-background-color-hover,var(--neutral-fill-input-alt-active))}:host .file-list-wrapper .show-more-text{font-size:var(--show-more-button-font-size,12px)}
`],_={showMoreSubtitle:"Show more items",filesSectionTitle:"Files",sharedTextSubtitle:"Shared"};var h=n("Y1eq"),b=n("tezw"),g=n("6BDD"),v=n("6vBc"),y=n("4X57"),S=n("j9Xn"),D=n("xY0q"),I=n("oqLQ"),x=n("8hiW");class C extends b.a{}const O=C.compose({baseName:"progress",template:(e,t)=>g.a`
    <template
        role="progressbar"
        aria-valuenow="${e=>e.value}"
        aria-valuemin="${e=>e.min}"
        aria-valuemax="${e=>e.max}"
        class="${e=>e.paused?"paused":""}"
    >
        ${Object(v.a)(e=>"number"==typeof e.value,g.a`
                <div class="progress" part="progress" slot="determinate">
                    <div
                        class="determinate"
                        part="determinate"
                        style="width: ${e=>e.percentComplete}%"
                    ></div>
                </div>
            `,g.a`
                <div class="progress" part="progress" slot="indeterminate">
                    <slot class="indeterminate" name="indeterminate">
                        ${t.indeterminateIndicator1||""}
                        ${t.indeterminateIndicator2||""}
                    </slot>
                </div>
            `)}
    </template>
`,styles:(e,t)=>y.a`
    ${Object(D.a)("flex")} :host {
      align-items: center;
      height: calc((${x.vb} * 3) * 1px);
    }

    .progress {
      background-color: ${x.ub};
      border-radius: calc(${x.s} * 1px);
      width: 100%;
      height: calc(${x.vb} * 1px);
      display: flex;
      align-items: center;
      position: relative;
    }

    .determinate {
      background-color: ${x.e};
      border-radius: calc(${x.s} * 1px);
      height: calc((${x.vb} * 3) * 1px);
      transition: all 0.2s ease-in-out;
      display: flex;
    }

    .indeterminate {
      height: calc((${x.vb} * 3) * 1px);
      border-radius: calc(${x.s} * 1px);
      display: flex;
      width: 100%;
      position: relative;
      overflow: hidden;
    }

    .indeterminate-indicator-1 {
      position: absolute;
      opacity: 0;
      height: 100%;
      background-color: ${x.e};
      border-radius: calc(${x.s} * 1px);
      animation-timing-function: cubic-bezier(0.4, 0, 0.6, 1);
      width: 40%;
      animation: indeterminate-1 2s infinite;
    }

    .indeterminate-indicator-2 {
      position: absolute;
      opacity: 0;
      height: 100%;
      background-color: ${x.e};
      border-radius: calc(${x.s} * 1px);
      animation-timing-function: cubic-bezier(0.4, 0, 0.6, 1);
      width: 60%;
      animation: indeterminate-2 2s infinite;
    }

    :host(.paused) .indeterminate-indicator-1,
    :host(.paused) .indeterminate-indicator-2 {
      animation: none;
      background-color: ${x.cb};
      width: 100%;
      opacity: 1;
    }

    :host(.paused) .determinate {
      background-color: ${x.cb};
    }

    @keyframes indeterminate-1 {
      0% {
        opacity: 1;
        transform: translateX(-100%);
      }
      70% {
        opacity: 1;
        transform: translateX(300%);
      }
      70.01% {
        opacity: 0;
      }
      100% {
        opacity: 0;
        transform: translateX(300%);
      }
    }

    @keyframes indeterminate-2 {
      0% {
        opacity: 0;
        transform: translateX(-150%);
      }
      29.99% {
        opacity: 0;
      }
      30% {
        opacity: 1;
        transform: translateX(-150%);
      }
      100% {
        transform: translateX(166.66%);
        opacity: 1;
      }
    }
  `.withBehaviors(Object(I.a)(y.a`
        .indeterminate-indicator-1,
        .indeterminate-indicator-2,
        .determinate,
        .progress {
          background-color: ${S.a.ButtonText};
        }
        :host(.paused) .indeterminate-indicator-1,
        :host(.paused) .indeterminate-indicator-2,
        :host(.paused) .determinate {
          background-color: ${S.a.GrayText};
        }
      `)),indeterminateIndicator1:'\n    <span class="indeterminate-indicator-1" part="indeterminate-indicator-1"></span>\n  ',indeterminateIndicator2:'\n    <span class="indeterminate-indicator-2" part="indeterminate-indicator-2"></span>\n  '});var w=n("m1Vi"),E=n("kd7Q"),A=n("eftJ"),L=n("5ZAu"),k=n("oePG"),M=n("QBS5"),P=n("Gy7L"),T=["input","select","textarea","a[href]","button","[tabindex]:not(slot)","audio[controls]","video[controls]",'[contenteditable]:not([contenteditable="false"])',"details>summary:first-of-type","details"].join(","),U="undefined"==typeof Element,F=U?function(){}:Element.prototype.matches||Element.prototype.msMatchesSelector||Element.prototype.webkitMatchesSelector,H=!U&&Element.prototype.getRootNode?function(e){return e.getRootNode()}:function(e){return e.ownerDocument},R=function(e){return"INPUT"===e.tagName},N=function(e){var t=e.getBoundingClientRect(),n=t.width,a=t.height;return 0===n&&0===a},B=function(e,t){if(t=t||{},!e)throw new Error("No node provided");return!1!==F.call(e,T)&&function(e,t){return!(function(e){return function(e){return R(e)&&"radio"===e.type}(e)&&!function(e){if(!e.name)return!0;var t,n=e.form||H(e),a=function(e){return n.querySelectorAll('input[type="radio"][name="'+e+'"]')};if("undefined"!=typeof window&&void 0!==window.CSS&&"function"==typeof window.CSS.escape)t=a(window.CSS.escape(e.name));else try{t=a(e.name)}catch(e){return console.error("Looks like you have a radio button with a name attribute containing invalid CSS selector characters and need the CSS.escape polyfill: %s",e.message),!1}var i=function(e,t){for(var n=0;n<e.length;n++)if(e[n].checked&&e[n].form===t)return e[n]}(t,e.form);return!i||i===e}(e)}(t)||function(e,t){return e.tabIndex<0&&(/^(AUDIO|VIDEO|DETAILS)$/.test(e.tagName)||e.isContentEditable)&&isNaN(parseInt(e.getAttribute("tabindex"),10))?0:e.tabIndex}(t)<0||!function(e,t){return!(t.disabled||function(e){return R(e)&&"hidden"===e.type}(t)||function(e,t){var n=t.displayCheck,a=t.getShadowRoot;if("hidden"===getComputedStyle(e).visibility)return!0;var i=F.call(e,"details>summary:first-of-type")?e.parentElement:e;if(F.call(i,"details:not([open]) *"))return!0;var r=H(e).host,o=(null==r?void 0:r.ownerDocument.contains(r))||e.ownerDocument.contains(e);if(n&&"full"!==n){if("non-zero-area"===n)return N(e)}else{if("function"==typeof a){for(var s=e;e;){var c=e.parentElement,d=H(e);if(c&&!c.shadowRoot&&!0===a(c))return N(e);e=e.assignedSlot?e.assignedSlot:c||d===e.ownerDocument?c:d.host}e=s}if(o)return!e.getClientRects().length}return!1}(t,e)||function(e){return"DETAILS"===e.tagName&&Array.prototype.slice.apply(e.children).some(function(e){return"SUMMARY"===e.tagName})}(t)||function(e){if(/^(INPUT|BUTTON|SELECT|TEXTAREA)$/.test(e.tagName))for(var t=e.parentElement;t;){if("FIELDSET"===t.tagName&&t.disabled){for(var n=0;n<t.children.length;n++){var a=t.children.item(n);if("LEGEND"===a.tagName)return!!F.call(t,"fieldset[disabled] *")||!a.contains(e)}return!0}t=t.parentElement}return!1}(t))}(e,t))}(t,e)},j=n("TDEi");class V extends j.a{constructor(){super(...arguments),this.modal=!0,this.hidden=!1,this.trapFocus=!0,this.trapFocusChanged=()=>{this.$fastController.isConnected&&this.updateTrapFocus()},this.isTrappingFocus=!1,this.handleDocumentKeydown=e=>{if(!e.defaultPrevented&&!this.hidden)switch(e.key){case P.h:this.dismiss(),e.preventDefault();break;case P.n:this.handleTabKeyDown(e)}},this.handleDocumentFocus=e=>{!e.defaultPrevented&&this.shouldForceFocus(e.target)&&(this.focusFirstElement(),e.preventDefault())},this.handleTabKeyDown=e=>{if(!this.trapFocus||this.hidden)return;const t=this.getTabQueueBounds();return 0!==t.length?1===t.length?(t[0].focus(),void e.preventDefault()):void(e.shiftKey&&e.target===t[0]?(t[t.length-1].focus(),e.preventDefault()):e.shiftKey||e.target!==t[t.length-1]||(t[0].focus(),e.preventDefault())):void 0},this.getTabQueueBounds=()=>V.reduceTabbableItems([],this),this.focusFirstElement=()=>{const e=this.getTabQueueBounds();e.length>0?e[0].focus():this.dialog instanceof HTMLElement&&this.dialog.focus()},this.shouldForceFocus=e=>this.isTrappingFocus&&!this.contains(e),this.shouldTrapFocus=()=>this.trapFocus&&!this.hidden,this.updateTrapFocus=e=>{const t=void 0===e?this.shouldTrapFocus():e;t&&!this.isTrappingFocus?(this.isTrappingFocus=!0,document.addEventListener("focusin",this.handleDocumentFocus),L.a.queueUpdate(()=>{this.shouldForceFocus(document.activeElement)&&this.focusFirstElement()})):!t&&this.isTrappingFocus&&(this.isTrappingFocus=!1,document.removeEventListener("focusin",this.handleDocumentFocus))}}dismiss(){this.$emit("dismiss"),this.$emit("cancel")}show(){this.hidden=!1}hide(){this.hidden=!0,this.$emit("close")}connectedCallback(){super.connectedCallback(),document.addEventListener("keydown",this.handleDocumentKeydown),this.notifier=k.b.getNotifier(this),this.notifier.subscribe(this,"hidden"),this.updateTrapFocus()}disconnectedCallback(){super.disconnectedCallback(),document.removeEventListener("keydown",this.handleDocumentKeydown),this.updateTrapFocus(!1),this.notifier.unsubscribe(this,"hidden")}handleChange(e,t){"hidden"===t&&this.updateTrapFocus()}static reduceTabbableItems(e,t){return"-1"===t.getAttribute("tabindex")?e:B(t)||V.isFocusableFastElement(t)&&V.hasTabbableShadow(t)?(e.push(t),e):t.childElementCount?e.concat(Array.from(t.children).reduce(V.reduceTabbableItems,[])):e}static isFocusableFastElement(e){var t,n;return!!(null===(n=null===(t=e.$fastController)||void 0===t?void 0:t.definition.shadowOptions)||void 0===n?void 0:n.delegatesFocus)}static hasTabbableShadow(e){var t,n;return Array.from(null!==(n=null===(t=e.shadowRoot)||void 0===t?void 0:t.querySelectorAll("*"))&&void 0!==n?n:[]).some(e=>B(e))}}Object(A.a)([Object(M.c)({mode:"boolean"})],V.prototype,"modal",void 0),Object(A.a)([Object(M.c)({mode:"boolean"})],V.prototype,"hidden",void 0),Object(A.a)([Object(M.c)({attribute:"trap-focus",mode:"boolean"})],V.prototype,"trapFocus",void 0),Object(A.a)([Object(M.c)({attribute:"aria-describedby"})],V.prototype,"ariaDescribedby",void 0),Object(A.a)([Object(M.c)({attribute:"aria-labelledby"})],V.prototype,"ariaLabelledby",void 0),Object(A.a)([Object(M.c)({attribute:"aria-label"})],V.prototype,"ariaLabel",void 0);var z=n("+53S"),G=n("cQsl");const K=V.compose({baseName:"dialog",template:(e,t)=>g.a`
    <div class="positioning-region" part="positioning-region">
        ${Object(v.a)(e=>e.modal,g.a`
                <div
                    class="overlay"
                    part="overlay"
                    role="presentation"
                    @click="${e=>e.dismiss()}"
                ></div>
            `)}
        <div
            role="dialog"
            tabindex="-1"
            class="control"
            part="control"
            aria-modal="${e=>e.modal}"
            aria-describedby="${e=>e.ariaDescribedby}"
            aria-labelledby="${e=>e.ariaLabelledby}"
            aria-label="${e=>e.ariaLabel}"
            ${Object(z.a)("dialog")}
        >
            <slot></slot>
        </div>
    </div>
`,styles:(e,t)=>y.a`
  :host([hidden]) {
    display: none;
  }

  :host {
    --dialog-height: 480px;
    --dialog-width: 640px;
    display: block;
  }

  .overlay {
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: rgba(0, 0, 0, 0.3);
    touch-action: none;
  }

  .positioning-region {
    display: flex;
    justify-content: center;
    position: fixed;
    top: 0;
    bottom: 0;
    left: 0;
    right: 0;
    overflow: auto;
  }

  .control {
    box-shadow: ${G.b};
    margin-top: auto;
    margin-bottom: auto;
    border-radius: calc(${x.D} * 1px);
    width: var(--dialog-width);
    height: var(--dialog-height);
    background: ${x.v};
    z-index: 1;
    border: calc(${x.vb} * 1px) solid transparent;
  }
`});var W=n("dblZ"),q=n("h2QR"),Q=n("VDxL"),Y=n("/4Fm");const J=[c.b`
:host .file-upload-area-button{width:auto;display:flex;align-items:end;justify-content:end;margin-inline-end:36px;margin-top:30px}:host fluent-button .upload-icon path{fill:var(--file-upload-button-text-color,var(--foreground-on-accent-rest))}:host fluent-button.file-upload-button::part(control){border:var(--file-upload-button-border,none);background:var(--file-upload-button-background-color,var(--accent-fill-rest))}:host fluent-button.file-upload-button::part(control):hover{background:var(--file-upload-button-background-color-hover,var(--accent-fill-hover))}:host fluent-button.file-upload-button .upload-text{color:var(--file-upload-button-text-color,var(--foreground-on-accent-rest));font-weight:400;line-height:20px}:host input{display:none}:host fluent-progress.file-upload-bar{width:180px;margin-top:10px}:host fluent-dialog::part(overlay){opacity:.5}:host fluent-dialog::part(control){--dialog-width:$file-upload-dialog-width;--dialog-height:$file-upload-dialog-height;padding:var(--file-upload-dialog-padding,24px);border:var(--file-upload-dialog-border,1px solid var(--neutral-fill-rest))}:host fluent-dialog .file-upload-dialog-ok{background:var(--file-upload-dialog-keep-both-button-background-color,var(--accent-fill-rest));border:var(--file-upload-dialog-keep-both-button-border,none);color:var(--file-upload-dialog-keep-both-button-text-color,var(--foreground-on-accent-rest))}:host fluent-dialog .file-upload-dialog-ok:hover{background:var(--file-upload-dialog-keep-both-button-background-color-hover,var(--accent-fill-hover))}:host fluent-dialog .file-upload-dialog-cancel{background:var(--file-upload-dialog-replace-button-background-color,var(--accent-fill-rest));border:var(--file-upload-dialog-replace-button-border,1px solid var(--neutral-foreground-rest));color:var(--file-upload-dialog-replace-button-text-color,var(--neutral-foreground-rest))}:host fluent-dialog .file-upload-dialog-cancel:hover{background:var(--file-upload-dialog-replace-button-background-color-hover,var(--accent-fill-hover))}:host fluent-checkbox{margin-top:12px}:host fluent-checkbox .file-upload-dialog-check{color:var(--file-upload-dialog-text-color,--foreground-on-accent-rest)}:host .file-upload-table{display:flex}:host .file-upload-table.upload{display:flex}:host .file-upload-table .file-upload-cell{padding:1px 0 1px 1px;display:table-cell;vertical-align:middle;position:relative}:host .file-upload-table .file-upload-cell.percent-indicator{padding-inline-start:10px}:host .file-upload-table .file-upload-cell .description{opacity:.5;position:relative}:host .file-upload-table .file-upload-cell .file-upload-filename{max-width:250px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}:host .file-upload-table .file-upload-cell .file-upload-status{position:absolute;left:28px}:host .file-upload-table .file-upload-cell .file-upload-cancel{cursor:pointer;margin-inline-start:20px}:host .file-upload-table .file-upload-cell .file-upload-name{width:auto}:host .file-upload-table .file-upload-cell .cancel-icon{fill:var(--file-upload-dialog-text-color,var(--neutral-foreground-rest))}:host .mgt-file-item{--file-background-color:transparent;--file-padding:0 12px;--file-padding-inline-start:24px}:host .file-upload-template{clear:both}:host .file-upload-template .file-upload-folder-tab{padding-inline-start:20px}:host .file-upload-dialog{display:none}:host .file-upload-dialog .file-upload-dialog-content{background-color:var(--file-upload-dialog-background-color,var(--accent-fill-rest));color:var(--file-upload-dialog-text-color,var(--neutral-foreground-rest))}:host .file-upload-dialog .file-upload-dialog-content-text{margin-bottom:36px}:host .file-upload-dialog .file-upload-dialog-title{margin-top:0}:host .file-upload-dialog .file-upload-dialog-editor{display:flex;align-items:end;justify-content:end;gap:5px}:host .file-upload-dialog .file-upload-dialog-close{float:right;cursor:pointer}:host .file-upload-dialog .file-upload-dialog-close svg{fill:var(--file-upload-dialog-text-color,var(--neutral-foreground-rest));padding-right:5px}:host .file-upload-dialog.visible{display:block}:host fluent-checkbox.file-upload-dialog-check.hide{display:none}:host .file-upload-dialog-success{cursor:pointer;opacity:.5}:host #file-upload-border{display:none}:host #file-upload-border.visible{border:var(--file-upload-border-drag,1px dashed #0078d4);background-color:var(--file-upload-background-color-drag,rgba(0,120,212,.1));position:absolute;inset:0;z-index:1;display:inline-block}[dir=rtl] :host .file-upload-status{left:0;right:28px}
`],X={failUploadFile:"File upload fail.",cancelUploadFile:"File cancel.",buttonUploadFile:"Upload Files",maximumFilesTitle:"Maximum files",maximumFiles:"Sorry, the maximum number of files you can upload at once is {MaxNumber}. Do you want to upload the first {MaxNumber} files or reselect?",maximumFileSizeTitle:"Maximum files size",maximumFileSize:'Sorry, the maximum file size to upload is {FileSize}. The file "{FileName}" has ',fileTypeTitle:"File type",fileType:'Sorry, the format of following file "{FileName}" cannot be uploaded.',checkAgain:"Don't show again",checkApplyAll:"Apply to all",buttonOk:"OK",buttonCancel:"Cancel",buttonUpload:"Upload",buttonKeep:"Keep both",buttonReplace:"Replace",buttonReselect:"Reselect",fileReplaceTitle:"Replace file",fileReplace:'Do you want to replace the file "{FileName}" or keep both files?',uploadButtonLabel:"File upload button"};var Z=n("0mOt"),$=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},ee=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)},te=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const ne=e=>e.isDirectory,ae=e=>"getAsEntry"in e&&"function"==typeof e.getAsEntry;class ie extends W.b{static get styles(){return J}get strings(){return X}static get requiredScopes(){return[...new Set(["files.readwrite","files.readwrite.all","sites.readwrite.all"])]}get _dropEffect(){return"copy"}constructor(){super(),this._dragCounter=0,this._maxChunkSize=4194304,this._dialogTitle="",this._dialogContent="",this._dialogPrimaryButton="",this._dialogSecondaryButton="",this._dialogCheckBox="",this._applyAll=!1,this._applyAllConflictBehavior=null,this._maximumFileSize=!1,this._excludedFileType=!1,this.onFileUploadChange=e=>{const t=e.target;!e||t.files.length<1||this.readUploadedFiles(t.files,()=>t.value=null)},this.onFileUploadClick=()=>{this.renderRoot.querySelector("#file-upload-input").click()},this.handleonDragOver=e=>{e.preventDefault(),e.stopPropagation(),e.dataTransfer.items&&e.dataTransfer.items.length>0&&(e.dataTransfer.dropEffect=e.dataTransfer.dropEffect=this._dropEffect)},this.handleonDragEnter=e=>{e.preventDefault(),e.stopPropagation(),this._dragCounter++,e.dataTransfer.items&&e.dataTransfer.items.length>0&&(e.dataTransfer.dropEffect=this._dropEffect,this.renderRoot.querySelector("#file-upload-border").classList.add("visible"))},this.handleonDragLeave=e=>{e.preventDefault(),e.stopPropagation(),this._dragCounter--,0===this._dragCounter&&this.renderRoot.querySelector("#file-upload-border").classList.remove("visible")},this.handleonDrop=e=>{var t;e.preventDefault(),e.stopPropagation(),this.renderRoot.querySelector("#file-upload-border").classList.remove("visible"),(null===(t=e.dataTransfer)||void 0===t?void 0:t.items)&&this.readUploadedFiles(e.dataTransfer.items,()=>{e.dataTransfer.clearData()}),this._dragCounter=0},this.filesToUpload=[]}render(){if(null!==this.parentElement){const e=this.parentElement;e.addEventListener("dragenter",this.handleonDragEnter),e.addEventListener("dragleave",this.handleonDragLeave),e.addEventListener("dragover",this.handleonDragOver),e.addEventListener("drop",this.handleonDrop)}return c.c`
        <div id="file-upload-dialog" class="file-upload-dialog">
          <!-- Modal content -->
          <fluent-dialog modal="true" class="file-upload-dialog-content">
            <span
              class="file-upload-dialog-close"
              id="file-upload-dialog-close">
                ${Object(f.b)(f.a.Cancel)}
            </span>
            <div class="file-upload-dialog-content-text">
              <h2 class="file-upload-dialog-title">${this._dialogTitle}</h2>
              <div>${this._dialogContent}</div>
                <fluent-checkbox
                  id="file-upload-dialog-check"
                  class="file-upload-dialog-check">
                    ${this._dialogCheckBox}
                </fluent-checkbox>
            </div>
            <div class="file-upload-dialog-editor">
              <fluent-button
                appearance="accent"
                class="file-upload-dialog-ok">
                ${this._dialogPrimaryButton}
              </fluent-button>
              <fluent-button
                appearance="outline"
                class="file-upload-dialog-cancel">
                ${this._dialogSecondaryButton}
              </fluent-button>
            </div>
          </fluent-dialog>
        </div>
        <div id="file-upload-border"></div>
        <div class="file-upload-area-button">
          <input
            id="file-upload-input"
            title="${this.strings.uploadButtonLabel}"
            tabindex="-1"
            aria-label="file upload input"
            type="file"
            multiple
            @change="${this.onFileUploadChange}"
          />
          <fluent-button
            appearance="accent"
            class="file-upload-button"
            @click=${this.onFileUploadClick}
            label=${this.strings.uploadButtonLabel}>
              <span slot="start">${Object(f.b)(f.a.Upload)}</span>
              <span class="upload-text">${this.strings.buttonUploadFile}</span>
          </fluent-button>
        </div>
        <div class="file-upload-template">
          ${this.renderFolderTemplate(this.filesToUpload)}
        </div>
       `}renderFolderTemplate(e){const t=[];if(e.length>0){const n=e.map(e=>-1===t.indexOf(e.fullPath.substring(0,e.fullPath.lastIndexOf("/")))?e.fullPath.endsWith("/")?c.c`${this.renderFileTemplate(e,"")}`:(t.push(e.fullPath.substring(0,e.fullPath.lastIndexOf("/"))),o.a`
            <div class='file-upload-table'>
              <div class='file-upload-cell'>
                <mgt-file
                  .fileDetails=${{name:e.fullPath.substring(1,e.fullPath.lastIndexOf("/")),folder:"Folder"}}
                  .view=${p.a.oneline}
                  class="mgt-file-item">
                </mgt-file>
              </div>
            </div>
            ${this.renderFileTemplate(e,"file-upload-folder-tab")}`):c.c`${this.renderFileTemplate(e,"file-upload-folder-tab")}`);return c.c`${n}`}return c.c``}renderFileTemplate(e,t){const n=Object(q.a)({"file-upload-table":!0,upload:e.completed}),a=t+("lastModifiedDateTime"===e.fieldUploadResponse?" file-upload-dialog-success":""),i=Object(q.a)({description:"description"===e.fieldUploadResponse}),r=e.completed?c.c``:this.renderFileUploadTemplate(e);return o.a`
        <div class="${n}">
          <div class="${a}">
            <div class='file-upload-cell'>
              <div class="${i}">
                <div class="file-upload-status">
                  ${e.iconStatus}
                </div>
                <mgt-file
                  .fileDetails=${e.driveItem}
                  .view=${e.view}
                  .line2Property=${e.fieldUploadResponse}
                  part="upload"
                  class="mgt-file-item">
                </mgt-file>
              </div>
            </div>
              ${r}
            </div>
          </div>
        </div>`}renderFileUploadTemplate(e){const t=Object(q.a)({"file-upload-table":!0,upload:e.completed});return c.c`
    <div class='file-upload-cell'>
      <div class='file-upload-table file-upload-name'>
        <div class='file-upload-cell'>
          <div
            title="${e.file.name}"
            class='file-upload-filename'>
            ${e.file.name}
          </div>
        </div>
      </div>
      <div class='file-upload-table'>
        <div class='file-upload-cell'>
          <div class="${t}">
            <fluent-progress class="file-upload-bar" value="${e.percent}"></fluent-progress>
            <div class='file-upload-cell percent-indicator'>
              <span>${e.percent}%</span>
              <span
                class="file-upload-cancel"
                @click=${()=>this.deleteFileUploadSession(e)}>
                ${Object(f.b)(f.a.Cancel)}
              </span>
            </div>
          </div>
        </div>
      </div>
    </div>
    `}deleteFileUploadSession(e){return te(this,void 0,void 0,function*(){try{void 0!==e.uploadUrl?(yield Object(u.b)(this.fileUploadList.graph,e.uploadUrl),e.uploadUrl=void 0,e.completed=!0,this.setUploadFail(e,X.cancelUploadFile)):(e.uploadUrl=void 0,e.completed=!0,this.setUploadFail(e,X.cancelUploadFile))}catch(t){e.uploadUrl=void 0,e.completed=!0,this.setUploadFail(e,X.cancelUploadFile)}})}readUploadedFiles(e,t){return te(this,void 0,void 0,function*(){const n=yield this.getFilesFromUploadArea(e);yield this.getSelectedFiles(n),t()})}getSelectedFiles(e){return te(this,void 0,void 0,function*(){let t=[];const n=[];this._applyAll=!1,this._applyAllConflictBehavior=null,this._maximumFileSize=!1,this._excludedFileType=!1,this.filesToUpload.forEach(e=>{e.completed?n.push(e):t.push(e)});for(const n of e){const e=""===n.fullPath?"/"+n.name:n.fullPath;if(0===t.filter(t=>t.fullPath===e).length){let a=!0;if(void 0!==this.fileUploadList.maxFileSize&&a&&n.size>1024*this.fileUploadList.maxFileSize&&(a=!1,!1===this._maximumFileSize)){const t=yield this.getFileUploadStatus(n,e,"MaxFileSize",this.fileUploadList);null!==t&&1===t[0]&&(this._maximumFileSize=!0)}if(void 0!==this.fileUploadList.excludedFileExtensions&&this.fileUploadList.excludedFileExtensions.length>0&&a&&this.fileUploadList.excludedFileExtensions.filter(e=>n.name.toLowerCase().indexOf(e.toLowerCase())>-1).length>0&&(a=!1,!1===this._excludedFileType)){const t=yield this.getFileUploadStatus(n,e,"ExcludedFileType",this.fileUploadList);null!==t&&1===t[0]&&(this._excludedFileType=!0)}if(a){const a=yield this.getFileUploadStatus(n,e,"Upload",this.fileUploadList);let i=!1;null!==a&&(-1===a[0]?i=!0:(this._applyAll=Boolean(a[0]),this._applyAllConflictBehavior=a[1]?1:0)),t.push({file:n,driveItem:{name:n.name},fullPath:e,conflictBehavior:null!==a?a[1]?1:0:null,iconStatus:null,percent:1,view:p.a.image,completed:i,maxSize:this._maxChunkSize,minSize:0})}}}t=t.sort((e,t)=>e.fullPath.substring(0,e.fullPath.lastIndexOf("/")).localeCompare(t.fullPath.substring(0,t.fullPath.lastIndexOf("/")))),t.forEach(e=>{if(0!==n.filter(t=>t.fullPath===e.fullPath).length){const t=n.findIndex(t=>t.fullPath===e.fullPath);n.splice(t,1)}}),t.push(...n),this.filesToUpload=t;const a=this.filesToUpload.map(e=>this.sendFileItemGraph(e));yield Promise.all(a)})}getFileUploadStatus(e,t,n,a){const i=Object.create(null,{requestStateUpdate:{get:()=>super.requestStateUpdate}});return te(this,void 0,void 0,function*(){const r=this.renderRoot.querySelector("#file-upload-dialog");switch(n){case"Upload":return null!==(yield Object(u.o)(this.fileUploadList.graph,`${this.getGrapQuery(t)}?$select=id`))?!0===this._applyAll?[this._applyAll,this._applyAllConflictBehavior]:(r.classList.add("visible"),this._dialogTitle=X.fileReplaceTitle,this._dialogContent=X.fileReplace.replace("{FileName}",e.name),this._dialogCheckBox=X.checkApplyAll,this._dialogPrimaryButton=X.buttonReplace,this._dialogSecondaryButton=X.buttonKeep,yield i.requestStateUpdate.call(this,!0),new Promise(e=>{const t=this.renderRoot.querySelector(".file-upload-dialog-close"),n=this.renderRoot.querySelector(".file-upload-dialog-ok"),a=this.renderRoot.querySelector(".file-upload-dialog-cancel"),i=this.renderRoot.querySelector("#file-upload-dialog-check");i.checked=!1,i.classList.remove("hide");const o=()=>{r.classList.remove("visible"),e([i.checked?1:0,1])},s=()=>{r.classList.remove("visible"),e([i.checked?1:0,0])},c=()=>{r.classList.remove("visible"),e([-1])};n.removeEventListener("click",o),a.removeEventListener("click",s),t.removeEventListener("click",c),n.addEventListener("click",o),a.addEventListener("click",s),t.addEventListener("click",c)})):null;case"ExcludedFileType":return r.classList.add("visible"),this._dialogTitle=X.fileTypeTitle,this._dialogContent=X.fileType.replace("{FileName}",e.name)+" ("+a.excludedFileExtensions.join(",")+")",this._dialogCheckBox=X.checkAgain,this._dialogPrimaryButton=X.buttonOk,this._dialogSecondaryButton=X.buttonCancel,yield i.requestStateUpdate.call(this,!0),new Promise(e=>{const t=this.renderRoot.querySelector(".file-upload-dialog-ok"),n=this.renderRoot.querySelector(".file-upload-dialog-cancel"),a=this.renderRoot.querySelector(".file-upload-dialog-close"),i=this.renderRoot.querySelector("#file-upload-dialog-check");i.checked=!1,i.classList.remove("hide");const o=()=>{r.classList.remove("visible"),e([i.checked?1:0])},s=()=>{r.classList.remove("visible"),e([0])};t.removeEventListener("click",o),n.removeEventListener("click",s),a.removeEventListener("click",s),t.addEventListener("click",o),n.addEventListener("click",s),a.addEventListener("click",s)});case"MaxFileSize":return r.classList.add("visible"),this._dialogTitle=X.maximumFileSizeTitle,this._dialogContent=X.maximumFileSize.replace("{FileSize}",Object(Y.d)(1024*a.maxFileSize)).replace("{FileName}",e.name)+Object(Y.d)(e.size)+".",this._dialogCheckBox=X.checkAgain,this._dialogPrimaryButton=X.buttonOk,this._dialogSecondaryButton=X.buttonCancel,yield i.requestStateUpdate.call(this,!0),new Promise(e=>{const t=this.renderRoot.querySelector(".file-upload-dialog-ok"),n=this.renderRoot.querySelector(".file-upload-dialog-cancel"),a=this.renderRoot.querySelector(".file-upload-dialog-close"),i=this.renderRoot.querySelector("#file-upload-dialog-check");i.checked=!1,i.classList.remove("hide");const o=()=>{r.classList.remove("visible"),e([i.checked?1:0])},s=()=>{r.classList.remove("visible"),e([0])};t.removeEventListener("click",o),n.removeEventListener("click",s),a.removeEventListener("click",s),t.addEventListener("click",o),n.addEventListener("click",s),a.addEventListener("click",s)})}})}getGrapQuery(e){let t="";return this.fileUploadList.itemPath&&this.fileUploadList.itemPath.length>0&&(t=this.fileUploadList.itemPath.startsWith("/")?this.fileUploadList.itemPath:"/"+this.fileUploadList.itemPath),this.fileUploadList.userId&&this.fileUploadList.itemId?`/users/${this.fileUploadList.userId}/drive/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.userId&&this.fileUploadList.itemPath?`/users/${this.fileUploadList.userId}/drive/root:${t}${e}`:this.fileUploadList.groupId&&this.fileUploadList.itemId?`/groups/${this.fileUploadList.groupId}/drive/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.groupId&&this.fileUploadList.itemPath?`/groups/${this.fileUploadList.groupId}/drive/root:${t}${e}`:this.fileUploadList.driveId&&this.fileUploadList.itemId?`/drives/${this.fileUploadList.driveId}/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.driveId&&this.fileUploadList.itemPath?`/drives/${this.fileUploadList.driveId}/root:${t}${e}`:this.fileUploadList.siteId&&this.fileUploadList.itemId?`/sites/${this.fileUploadList.siteId}/drive/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.siteId&&this.fileUploadList.itemPath?`/sites/${this.fileUploadList.siteId}/drive/root:${t}${e}`:this.fileUploadList.itemId?`/me/drive/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.itemPath?`/me/drive/root:${t}${e}`:`/me/drive/root:${e}`}sendFileItemGraph(e){return te(this,void 0,void 0,function*(){const t=this.fileUploadList.graph;let n="";if(e.file.size<this._maxChunkSize)try{e.completed||(null!==e.conflictBehavior&&1!==e.conflictBehavior||(n=`${this.getGrapQuery(e.fullPath)}:/content`),0===e.conflictBehavior&&(n=`${this.getGrapQuery(e.fullPath)}:/content?@microsoft.graph.conflictBehavior=rename`),e.driveItem=yield Object(u.L)(t,n,e.file),null!==e.driveItem?this.setUploadSuccess(e):(e.driveItem={name:e.file.name},this.setUploadFail(e,X.failUploadFile)))}catch(t){this.setUploadFail(e,X.failUploadFile)}else if(!e.completed&&void 0===e.uploadUrl){const n=yield Object(u.C)(t,`${this.getGrapQuery(e.fullPath)}:/createUploadSession`,e.conflictBehavior);try{if(null!==n){e.uploadUrl=n.uploadUrl;const a=yield this.sendSessionUrlGraph(t,e);null!==a?(e.driveItem=a,this.setUploadSuccess(e)):this.setUploadFail(e,X.failUploadFile)}else this.setUploadFail(e,X.failUploadFile)}catch(e){}}})}sendSessionUrlGraph(e,t){const n=Object.create(null,{requestStateUpdate:{get:()=>super.requestStateUpdate}});return te(this,void 0,void 0,function*(){for(;t.file.size>t.minSize;){void 0===t.mimeStreamString&&(t.mimeStreamString=yield this.readFileContent(t.file));const a=new Blob([t.mimeStreamString.slice(t.minSize,t.maxSize)]);if(t.percent=Math.round(t.maxSize/t.file.size*100),yield n.requestStateUpdate.call(this,!0),void 0===t.uploadUrl)return null;{const n=yield Object(u.K)(e,t.uploadUrl,""+(t.maxSize-t.minSize),`bytes ${t.minSize}-${t.maxSize-1}/${t.file.size}`,a);if(null===n)return null;if(Object(u.J)(n))t.minSize=parseInt(n.nextExpectedRanges[0].split("-")[0],10),t.maxSize=t.minSize+this._maxChunkSize,t.maxSize>t.file.size&&(t.maxSize=t.file.size);else if(void 0!==n.id)return n}}})}setUploadSuccess(e){e.percent=100,super.requestStateUpdate(!0),setTimeout(()=>{e.iconStatus=Object(f.b)(f.a.Success),e.view=p.a.twolines,e.fieldUploadResponse="lastModifiedDateTime",e.completed=!0,super.requestStateUpdate(!0),Object(u.a)()},500)}setUploadFail(e,t){setTimeout(()=>{e.iconStatus=Object(f.b)(f.a.Fail),e.view=p.a.twolines,e.driveItem.description=t,e.fieldUploadResponse="description",e.completed=!0,super.requestStateUpdate(!0)},500)}readFileContent(e){return new Promise((t,n)=>{const a=new FileReader;a.onloadend=()=>{t(a.result)},a.onerror=e=>{n(e)},a.readAsArrayBuffer(e)})}getFilesFromUploadArea(e){return te(this,void 0,void 0,function*(){const t=[];let n;const a=[];for(const r of e)if("getAsFile"in(i=r)&&"function"==typeof i.getAsFile||"webkitGetAsEntry"in i&&"function"==typeof i.webkitGetAsEntry)if(ae(r))if(n=r.getAsEntry(),ne(n))t.push(n);else{const e=r.getAsFile();e&&(this.writeFilePath(e,""),a.push(e))}else if(r.webkitGetAsEntry)if(n=r.webkitGetAsEntry(),ne(n))t.push(n);else{const e=r.getAsFile();e&&(this.writeFilePath(e,""),a.push(e))}else{const e=r.getAsFile();e&&(this.writeFilePath(e,""),a.push(e))}else this.writeFilePath(r,""),a.push(r);var i;if(t.length>0){const e=yield this.getFolderFiles(t);a.push(...e)}return a})}getFolderFiles(e){return new Promise(t=>{let n=0;const a=[];e.forEach(e=>{i(e,"")});const i=(e,i)=>{ne(e)?r(e.createReader()):(e=>e.isFile)(e)&&(n++,e.file(e=>{n--,this.writeFilePath(e,i),a.push(e),0===n&&t(a)}))},r=e=>{n++,e.readEntries(e=>{n--;for(const t of e)i(t,t.fullPath);0===n&&t(a)})}})}writeFilePath(e,t){e.fullPath=t}}$([Object(d.b)({type:Object}),ee("design:type",Array)],ie.prototype,"filesToUpload",void 0),$([Object(d.b)({type:Object}),ee("design:type",Object)],ie.prototype,"fileUploadList",void 0);var re=n("VcPv"),oe=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},se=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)},ce=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const de=()=>{Object(Q.a)(re.a),Object(h.b)(),Object(Q.a)(O,w.a,E.a,K),Object(h.b)(),Object(Z.b)("file-upload",ie),Object(Z.b)("file-list",le)};class le extends a.a{static get styles(){return m}get strings(){return _}get fileListQuery(){return this._fileListQuery}set fileListQuery(e){e!==this._fileListQuery&&(this._fileListQuery=e,this.requestStateUpdate(!0))}get displayName(){return this.strings.filesSectionTitle}get cardTitle(){return this.strings.filesSectionTitle}renderIcon(){return Object(f.b)(f.a.Files)}get fileQueries(){return this._fileQueries}set fileQueries(e){Object(i.a)(this._fileQueries,e)||(this._fileQueries=e,this.requestStateUpdate(!0))}get siteId(){return this._siteId}set siteId(e){e!==this._siteId&&(this._siteId=e,this.requestStateUpdate(!0))}get driveId(){return this._driveId}set driveId(e){e!==this._driveId&&(this._driveId=e,this.requestStateUpdate(!0))}get groupId(){return this._groupId}set groupId(e){e!==this._groupId&&(this._groupId=e,this.requestStateUpdate(!0))}get itemId(){return this._itemId}set itemId(e){e!==this._itemId&&(this._itemId=e,this.requestStateUpdate(!0))}get itemPath(){return this._itemPath}set itemPath(e){e!==this._itemPath&&(this._itemPath=e,this.requestStateUpdate(!0))}get userId(){return this._userId}set userId(e){e!==this._userId&&(this._userId=e,this.requestStateUpdate(!0))}get insightType(){return this._insightType}set insightType(e){e!==this._insightType&&(this._insightType=e,this.requestStateUpdate(!0))}get fileExtensions(){return this._fileExtensions}set fileExtensions(e){Object(i.a)(this._fileExtensions,e)||(this._fileExtensions=e,this.requestStateUpdate(!0))}get pageSize(){return this._pageSize}set pageSize(e){e!==this._pageSize&&(this._pageSize=e,this.requestStateUpdate(!0))}get maxFileSize(){return this._maxFileSize}set maxFileSize(e){e!==this._maxFileSize&&(this._maxFileSize=e,this.requestStateUpdate(!0))}get maxUploadFile(){return this._maxUploadFile}set maxUploadFile(e){e!==this._maxUploadFile&&(this._maxUploadFile=e,this.requestStateUpdate(!0))}get excludedFileExtensions(){return this._excludedFileExtensions}set excludedFileExtensions(e){Object(i.a)(this._excludedFileExtensions,e)||(this._excludedFileExtensions=e,this.requestStateUpdate(!0))}static get requiredScopes(){return[...new Set([...h.a.requiredScopes])]}constructor(){super(),this._isCompact=!1,this.disableOpenOnClick=!1,this._focusedItemIndex=-1,this.onFocusFirstItem=()=>this._focusedItemIndex=0,this.onFileListKeyDown=e=>{const t=this.renderRoot.querySelector(".file-list");let n;if(null==t?void 0:t.children.length){if("ArrowUp"!==e.code&&"ArrowDown"!==e.code||("ArrowUp"===e.code&&(-1===this._focusedItemIndex&&(this._focusedItemIndex=t.children.length),this._focusedItemIndex=(this._focusedItemIndex-1+t.children.length)%t.children.length),"ArrowDown"===e.code&&(this._focusedItemIndex=(this._focusedItemIndex+1)%t.children.length),n=t.children[this._focusedItemIndex],this.updateItemBackgroundColor(t,n,"focused")),"Enter"===e.code||"Space"===e.code){n=t.children[this._focusedItemIndex];const a=n.children[0];e.preventDefault(),this.fireCustomEvent("itemClick",a.fileDetails),this.handleFileClick(a.fileDetails),this.updateItemBackgroundColor(t,n,"selected")}"Tab"===e.code&&(n=t.children[this._focusedItemIndex])}},this.pageSize=10,this.itemView=p.a.twolines,this.maxUploadFile=10,this.enableFileUpload=!1,this._preloadedFiles=[]}requestStateUpdate(e){return this.clearState(),super.requestStateUpdate(e)}clearState(){super.clearState(),this.files=null}asCompactView(){return this._isCompact=!0,this}asFullView(){return this._isCompact=!1,this}render(){return!this.files&&this.isLoadingState?this.renderLoading():this.files&&0!==this.files.length?this._isCompact?this.renderCompactView():this.renderFullView():this.renderNoData()}renderCompactView(){const e=this.files.slice(0,3);return this.renderFiles(e)}renderFullView(){return this.renderTemplate("default",{files:this.files})||this.renderFiles(this.files)}renderLoading(){return this.renderTemplate("loading",null)||c.c``}renderNoData(){return this.renderTemplate("no-data",null)||(!0===this.enableFileUpload&&void 0!==r.a.globalProvider?c.c`
            <div class="file-list-wrapper" dir=${this.direction}>
              ${this.renderFileUpload()}
            </div>`:c.c``)}renderFiles(e){return c.c`
      <div id="file-list-wrapper" class="file-list-wrapper" dir=${this.direction}>
        ${this.enableFileUpload?this.renderFileUpload():null}
        <ul
          id="file-list"
          class="file-list"
        >
          <li
            tabindex="0"
            class="file-item"
            @keydown="${this.onFileListKeyDown}"
            @focus="${this.onFocusFirstItem}"
            @click=${t=>this.handleItemSelect(e[0],t)}>
            ${this.renderFile(e[0])}
          </li>
          ${Object(l.a)(e.slice(1),e=>e.id,e=>c.c`
              <li
                class="file-item"
                @keydown="${this.onFileListKeyDown}"
                @click=${t=>this.handleItemSelect(e,t)}>
                ${this.renderFile(e)}
              </li>
            `)}
        </ul>
        ${this.hideMoreFilesButton||!this.pageIterator||!this.pageIterator.hasNext&&!this._preloadedFiles.length||this._isCompact?null:this.renderMoreFileButton()}
      </div>
    `}renderFile(e){const t=this.itemView;return this.renderTemplate("file",{file:e},e.id)||o.a`
        <mgt-file class="mgt-file-item" .fileDetails=${e} .view=${t}></mgt-file>
      `}renderMoreFileButton(){return this._isLoadingMore?c.c`
        <fluent-progress-ring role="progressbar" viewBox="0 0 8 8" class="progress-ring"></fluent-progress-ring>
      `:c.c`
        <fluent-button
          appearance="stealth"
          id="show-more"
          class="show-more"
          @click=${()=>this.renderNextPage()}
        >
          <span class="show-more-text">${this.strings.showMoreSubtitle}</span>
        </fluent-button>`}renderFileUpload(){const e={graph:r.a.globalProvider.graph.forComponent(this),driveId:this.driveId,excludedFileExtensions:this.excludedFileExtensions,groupId:this.groupId,itemId:this.itemId,itemPath:this.itemPath,userId:this.userId,siteId:this.siteId,maxFileSize:this.maxFileSize,maxUploadFile:this.maxUploadFile};return o.a`
        <mgt-file-upload .fileUploadList=${e} ></mgt-file-upload>
      `}loadState(){var e;return ce(this,void 0,void 0,function*(){const t=r.a.globalProvider;if(!t||t.state===s.c.Loading)return;if(t.state===s.c.SignedOut)return void(this.files=null);const n=t.graph.forComponent(this);let a,i;const o=!(this.driveId||this.siteId||this.groupId||this.userId);if((this.driveId&&!this.itemId&&!this.itemPath||this.groupId&&!this.itemId&&!this.itemPath||this.siteId&&!this.itemId&&!this.itemPath||this.userId&&!this.insightType&&!this.itemId&&!this.itemPath)&&(this.files=null),!this.files){let t;if(this.fileListQuery?i=yield Object(u.k)(n,this.fileListQuery,this.pageSize):this.fileQueries?a=yield Object(u.m)(n,this.fileQueries):o?this.itemId?i=yield Object(u.j)(n,this.itemId,this.pageSize):this.itemPath?i=yield Object(u.l)(n,this.itemPath,this.pageSize):this.insightType?a=yield Object(u.x)(n,this.insightType):i=yield Object(u.n)(n,this.pageSize):this.driveId?this.itemId?i=yield Object(u.e)(n,this.driveId,this.itemId,this.pageSize):this.itemPath&&(i=yield Object(u.f)(n,this.driveId,this.itemPath,this.pageSize)):this.groupId?this.itemId?i=yield Object(u.r)(n,this.groupId,this.itemId,this.pageSize):this.itemPath&&(i=yield Object(u.s)(n,this.groupId,this.itemPath,this.pageSize)):this.siteId?this.itemId?i=yield Object(u.A)(n,this.siteId,this.itemId,this.pageSize):this.itemPath&&(i=yield Object(u.B)(n,this.siteId,this.itemPath,this.pageSize)):this.userId&&(this.itemId?i=yield Object(u.F)(n,this.userId,this.itemId,this.pageSize):this.itemPath?i=yield Object(u.G)(n,this.userId,this.itemPath,this.pageSize):this.insightType&&(a=yield Object(u.I)(n,this.userId,this.insightType))),i&&(this.pageIterator=i,this._preloadedFiles=[...this.pageIterator.value],a=this._preloadedFiles.length>=this.pageSize?this._preloadedFiles.splice(0,this.pageSize):this._preloadedFiles.splice(0,this._preloadedFiles.length)),this.fileExtensions&&null!==this.fileExtensions){if(null===(e=this.pageIterator)||void 0===e?void 0:e.value){for(;this.pageIterator.hasNext;)yield Object(u.c)(this.pageIterator);a=this.pageIterator.value,this._preloadedFiles=[]}t=a.filter(e=>{for(const t of this.fileExtensions)if(t===this.getFileExtension(e.name))return e})}(null==t?void 0:t.length)>=0?(this.files=t,this.pageSize&&(a=this.files.splice(0,this.pageSize),this.files=a)):this.files=a}})}handleItemSelect(e,t){if(this.handleFileClick(e),this.fireCustomEvent("itemClick",e),t){const e=this.renderRoot.querySelector(".file-list"),n=Array.from(e.children),a=t.target.closest("li"),i=n.indexOf(a);this._focusedItemIndex=i;const r=e.children[this._focusedItemIndex];this.updateItemBackgroundColor(e,r,"selected")}}renderNextPage(){return ce(this,void 0,void 0,function*(){if(this._preloadedFiles.length>0)this.files=[...this.files,...this._preloadedFiles.splice(0,Math.min(this.pageSize,this._preloadedFiles.length))];else if(this.pageIterator.hasNext){this._isLoadingMore=!0;const e=this.renderRoot.querySelector("#file-list-wrapper");(null==e?void 0:e.animate)&&e.animate([{height:"auto",transformOrigin:"top left"},{height:"auto",transformOrigin:"top left"}],{duration:1e3,easing:"ease-in-out",fill:"both"}),yield Object(u.c)(this.pageIterator),this._isLoadingMore=!1,this.files=this.pageIterator.value}this.requestUpdate()})}handleFileClick(e){(null==e?void 0:e.webUrl)&&!this.disableOpenOnClick&&window.open(e.webUrl,"_blank","noreferrer")}getFileExtension(e){return/(?:\.([^.]+))?$/.exec(e)[1]||""}updateItemBackgroundColor(e,t,n){for(const t of e.children)t.classList.remove(n),t.removeAttribute("tabindex");t&&(t.classList.add(n),t.scrollIntoView({behavior:"smooth",block:"nearest",inline:"start"}),t.setAttribute("tabindex","0"),t.focus());for(const t of e.children)t.classList.remove("selected")}reload(e=!1){e&&Object(u.a)(),this.requestStateUpdate(!0)}}oe([Object(d.c)(),se("design:type",Object)],le.prototype,"_isCompact",void 0),oe([Object(d.b)({attribute:"file-list-query"}),se("design:type",String),se("design:paramtypes",[String])],le.prototype,"fileListQuery",null),oe([Object(d.b)({attribute:"file-queries",converter:(e,t)=>e?e.split(",").map(e=>e.trim()):null}),se("design:type",Array),se("design:paramtypes",[Array])],le.prototype,"fileQueries",null),oe([Object(d.b)({type:Object}),se("design:type",Array)],le.prototype,"files",void 0),oe([Object(d.b)({attribute:"site-id"}),se("design:type",String),se("design:paramtypes",[String])],le.prototype,"siteId",null),oe([Object(d.b)({attribute:"drive-id"}),se("design:type",String),se("design:paramtypes",[String])],le.prototype,"driveId",null),oe([Object(d.b)({attribute:"group-id"}),se("design:type",String),se("design:paramtypes",[String])],le.prototype,"groupId",null),oe([Object(d.b)({attribute:"item-id"}),se("design:type",String),se("design:paramtypes",[String])],le.prototype,"itemId",null),oe([Object(d.b)({attribute:"item-path"}),se("design:type",String),se("design:paramtypes",[String])],le.prototype,"itemPath",null),oe([Object(d.b)({attribute:"user-id"}),se("design:type",String),se("design:paramtypes",[String])],le.prototype,"userId",null),oe([Object(d.b)({attribute:"insight-type"}),se("design:type",String),se("design:paramtypes",[String])],le.prototype,"insightType",null),oe([Object(d.b)({attribute:"item-view",converter:e=>e&&0!==e.length?(e=e.toLowerCase(),void 0===p.a[e]?p.a.threelines:p.a[e]):p.a.threelines}),se("design:type",Number)],le.prototype,"itemView",void 0),oe([Object(d.b)({attribute:"file-extensions",converter:(e,t)=>e.split(",").map(e=>e.trim())}),se("design:type",Array),se("design:paramtypes",[Array])],le.prototype,"fileExtensions",null),oe([Object(d.b)({attribute:"page-size",type:Number}),se("design:type",Number),se("design:paramtypes",[Number])],le.prototype,"pageSize",null),oe([Object(d.b)({attribute:"disable-open-on-click",type:Boolean}),se("design:type",Object)],le.prototype,"disableOpenOnClick",void 0),oe([Object(d.b)({attribute:"hide-more-files-button",type:Boolean}),se("design:type",Boolean)],le.prototype,"hideMoreFilesButton",void 0),oe([Object(d.b)({attribute:"max-file-size",type:Number}),se("design:type",Number),se("design:paramtypes",[Number])],le.prototype,"maxFileSize",null),oe([Object(d.b)({attribute:"enable-file-upload",type:Boolean}),se("design:type",Boolean)],le.prototype,"enableFileUpload",void 0),oe([Object(d.b)({attribute:"max-upload-file",type:Number}),se("design:type",Number),se("design:paramtypes",[Number])],le.prototype,"maxUploadFile",null),oe([Object(d.b)({attribute:"excluded-file-extensions",converter:(e,t)=>e.split(",").map(e=>e.trim())}),se("design:type",Array),se("design:paramtypes",[Array])],le.prototype,"excludedFileExtensions",null),oe([Object(d.c)(),se("design:type",Boolean)],le.prototype,"_isLoadingMore",void 0)},eftJ:function(e,t,n){"use strict";function a(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o}n.d(t,"a",function(){return a})},fU9z:function(e,t,n){"use strict";n.d(t,"b",function(){return a}),n.d(t,"c",function(){return i}),n.d(t,"a",function(){return r});const a=e=>"groupTypes"in e,i=e=>"personType"in e||"userType"in e,r=e=>"initials"in e},glp4:function(e,t,n){"use strict";n.d(t,"g",function(){return u}),n.d(t,"c",function(){return f}),n.d(t,"e",function(){return p}),n.d(t,"a",function(){return m}),n.d(t,"h",function(){return _}),n.d(t,"i",function(){return h}),n.d(t,"d",function(){return b}),n.d(t,"b",function(){return g}),n.d(t,"f",function(){return v}),n.d(t,"j",function(){return y});var a=n("DNu6"),i=n("RIOo"),r=n("EYXE"),o=n("/4Fm"),s=n("imsm"),c=n("z0DP"),d=n("s9K1"),l=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const u=()=>a.a.config.photos.invalidationPeriod||a.a.config.defaultInvalidationPeriod,f=()=>a.a.config.photos.isEnabled&&a.a.config.isEnabled,p=(e,t,n)=>l(void 0,void 0,void 0,function*(){try{const a=yield e.api(`${t}/photo/$value`).responseType(r.ResponseType.RAW).middlewareOptions(Object(i.a)(...n)).get();return 404===a.status?{eTag:null,photo:null}:a.ok?{eTag:a["@odata.mediaEtag"],photo:yield Object(o.a)(yield a.blob())}:null}catch(e){return null}}),m=(e,t)=>l(void 0,void 0,void 0,function*(){let n,i;return f()&&(n=a.a.getCache(s.a.photos,s.a.photos.stores.contacts),i=yield n.getValue(t),i&&u()>Date.now()-i.timeCached)?i.photo:(i=yield p(e,`me/contacts/${t}`,["contacts.read"]),f()&&i&&(yield n.putValue(t,i)),i?i.photo:null)}),_=(e,t)=>l(void 0,void 0,void 0,function*(){let n,i;if(f()){if(n=a.a.getCache(s.a.photos,s.a.photos.stores.users),i=yield n.getValue(t),i&&u()>Date.now()-i.timeCached)return i.photo;if(i)try{const n=yield e.api(`users/${t}/photo`).get();n&&(n["@odata.mediaEtag"]!==i.eTag||null===n["@odata.mediaEtag"]&&null===i.eTag)&&(i=null)}catch(e){return null}}return i=i||(yield p(e,`users/${t}`,["user.readbasic.all"])),f()&&i&&(yield n.putValue(t,i)),i?i.photo:null}),h=e=>l(void 0,void 0,void 0,function*(){let t,n;if(f()&&(t=a.a.getCache(s.a.photos,s.a.photos.stores.users),n=yield t.getValue("me"),n&&u()>Date.now()-n.timeCached))return n.photo;try{const t=yield e.api("me/photo").get();t&&(t["@odata.mediaEtag"]!==n.eTag||null===t["@odata.mediaEtag"]&&null===n.eTag)&&(n=null)}catch(e){return null}return n=n||(yield p(e,"me",["user.read"])),f()&&(yield t.putValue("me",n||{})),n?n.photo:null}),b=(e,t,n=!0)=>l(void 0,void 0,void 0,function*(){if("personType"in t&&"OrganizationUser"!==t.personType.subclass){if("PersonalContact"===t.personType.subclass&&n){const n=Object(c.e)(t),a=yield Object(c.c)(e,n);if((null==a?void 0:a.length)&&a[0].id)return yield m(e,a[0].id)}return null}if(t.userPrincipalName||t.id){const n=t.userPrincipalName||t.id;return yield _(e,n)}if(t.id){const n=yield _(e,t.id);if(n)return n}const a=Object(c.e)(t);if(a){const t=yield Object(d.b)(e,a,1);if(null==t?void 0:t.length)return yield _(e,t[0].id);if(n){const t=yield Object(c.c)(e,a);if(null==t?void 0:t.length)return yield m(e,t[0].id)}}return null}),g=(e,t)=>l(void 0,void 0,void 0,function*(){let n,i;const r=t.id;if(f()){if(i=a.a.getCache(s.a.photos,s.a.photos.stores.groups),n=yield i.getValue(r),n&&u()>Date.now()-n.timeCached)return n.photo;if(n)try{const t=yield e.api(`groups/${r}/photo`).get();t&&(t["@odata.mediaEtag"]!==n.eTag||null===t["@odata.mediaEtag"]&&null===n.eTag)&&(n=null)}catch(e){return null}}return n=n||(yield p(e,`groups/${r}`,["user.readbasic.all"])),f()&&n&&(yield i.putValue(r,n)),n?n.photo:null}),v=(e,t)=>l(void 0,void 0,void 0,function*(){const n=a.a.getCache(s.a.photos,t);return yield n.getValue(e)}),y=(e,t,n)=>l(void 0,void 0,void 0,function*(){const i=a.a.getCache(s.a.photos,t);yield i.putValue(e,n)})},h2QR:function(e,t,n){"use strict";n.d(t,"a",function(){return r});var a=n("5rba"),i=n("GcG9");const r=Object(i.c)(class extends i.a{constructor(e){var t;if(super(e),e.type!==i.b.ATTRIBUTE||"class"!==e.name||(null===(t=e.strings)||void 0===t?void 0:t.length)>2)throw Error("`classMap()` can only be used in the `class` attribute and must be the only part in the attribute.")}render(e){return" "+Object.keys(e).filter(t=>e[t]).join(" ")+" "}update(e,[t]){if(void 0===this.it){this.it=new Set,void 0!==e.strings&&(this.st=new Set(e.strings.join(" ").split(/\s/).filter(e=>""!==e)));for(const e in t){var n;t[e]&&(null===(n=this.st)||void 0===n||!n.has(e))&&this.it.add(e)}return this.render(t)}const i=e.element.classList;for(const e of this.it)e in t||(i.remove(e),this.it.delete(e));for(const e in t){var r;const n=!!t[e];n===this.it.has(e)||(null===(r=this.st)||void 0===r?void 0:r.has(e))||(n?(i.add(e),this.it.add(e)):(i.remove(e),this.it.delete(e)))}return a.c}})},hgjj:function(e,t,n){"use strict";n.d(t,"a",function(){return d});var a=n("DNu6"),i=n("RIOo"),r=n("glp4"),o=n("s9K1"),s=n("imsm"),c=n("o2vq");const d=(e,t,n)=>{return void 0,void 0,l=function*(){let d,l,u,f=null;const p=t?`users/${t}`:"me",m=p+(n?`?$select=${n.toString()}`:""),_=t?["user.readbasic.all"]:["user.read"];if(Object(o.d)()){const e=a.a.getCache(s.a.users,s.a.users.stores.users);u=yield e.getValue(t||"me"),u&&Object(o.g)()>Date.now()-u.timeCached?(f=u.user?JSON.parse(u.user):null,null!==f&&n&&n.filter(e=>!Object.keys(f).includes(e)).length>=1&&(f=null,u=null)):u=null}if(Object(r.c)())if(l=yield Object(r.f)(t||"me",s.a.photos.stores.users),l&&Object(r.g)()>Date.now()-l.timeCached)d=l.photo;else if(l)try{const n=yield e.api(`${p}/photo`).get();(null==n?void 0:n["@odata.mediaEtag"])&&n["@odata.mediaEtag"]===l.eTag?(yield Object(r.j)(t||"me",s.a.photos.stores.users,l),d=l.photo):l=null}catch(e){Object(c.a)(e)&&("ErrorItemNotFound"!==e.code&&"ImageNotFound"!==e.code||(yield Object(r.j)(t||"me",s.a.photos.stores.users,{eTag:null,photo:null})))}if(l||u)if(l){if(!u)try{const n=yield e.api(m).middlewareOptions(Object(i.a)(..._)).get();if(n){if(Object(o.d)()){const e=a.a.getCache(s.a.users,s.a.users.stores.users);yield e.putValue(t||"me",{user:JSON.stringify(n)})}f=n}}catch(e){}}else try{const n=yield Object(r.e)(e,p,_);n&&(Object(r.c)()&&(yield Object(r.j)(t||"me",s.a.photos.stores.users,{eTag:n.eTag,photo:n.photo})),d=n.photo)}catch(e){}else{let i;const c=e.createBatch();t?(c.get("user",`/users/${t}${n?"?$select="+n.toString():""}`,["user.readbasic.all"]),c.get("photo",`users/${t}/photo/$value`,["user.readbasic.all"])):(c.get("user","me",["user.read"]),c.get("photo","me/photo/$value",["user.read"]));const l=yield c.executeAll(),u=l.get("photo");u&&(i=u.headers.ETag,d=u.content);const p=l.get("user");if(p&&(f=p.content),Object(o.d)()){const e=a.a.getCache(s.a.users,s.a.users.stores.users);yield e.putValue(t||"me",{user:JSON.stringify(f)})}Object(r.c)()&&(yield Object(r.j)(t||"me",s.a.photos.stores.users,{eTag:i,photo:d}))}return f&&(f.personImage=d),f},new((d=void 0)||(d=Promise))(function(e,t){function n(e){try{i(l.next(e))}catch(e){t(e)}}function a(e){try{i(l.throw(e))}catch(e){t(e)}}function i(t){var i;t.done?e(t.value):(i=t.value,i instanceof d?i:new d(function(e){e(i)})).then(n,a)}i((l=l.apply(undefined,[])).next())});var d,l}},"hv+n":function(e,t,n){"use strict";n.d(t,"a",function(){return _});var a,i=n("0q6d"),r=n("swXE"),o=n("PT2o"),s=n("6eT7"),c=n("RN7+");function d(e,t,n){return isNaN(e)||e<=0?t:e>=1?n:new s.a(Object(i.e)(e,t.r,n.r),Object(i.e)(e,t.g,n.g),Object(i.e)(e,t.b,n.b),Object(i.e)(e,t.a,n.a))}n("tq/8"),n("SiT+"),n("xAa8"),function(e){e[e.RGB=0]="RGB",e[e.HSL=1]="HSL",e[e.HSV=2]="HSV",e[e.XYZ=3]="XYZ",e[e.LAB=4]="LAB",e[e.LCH=5]="LCH"}(a||(a={}));var l=n("i2HT");function u(e,t,n=0,a=e.length-1){if(a===n)return e[n];const i=Math.floor((a-n)/2)+n;return t(e[i])?u(e,t,n,i):u(e,t,i+1,a)}var f=n("s2f2"),p=n("THNU");const m={stepContrast:1.03,stepContrastRamp:.03,preserveSource:!1},_=Object.freeze({create:function(e,t,n){return"number"==typeof e?_.from(l.a.create(e,t,n)):_.from(e)},from:function(e,t){return Object(l.b)(e)?h.from(e,t):h.from(l.a.create(e.r,e.g,e.b),t)}});class h{constructor(e,t){this.closestIndexCache=new Map,this.source=e,this.swatches=t,this.reversedSwatches=Object.freeze([...this.swatches].reverse()),this.lastIndex=this.swatches.length-1}colorContrast(e,t,n,a){void 0===n&&(n=this.closestIndexOf(e));let i=this.swatches;const r=this.lastIndex;let o=n;return void 0===a&&(a=Object(f.a)(e)),-1===a&&(i=this.reversedSwatches,o=r-o),u(i,n=>Object(p.a)(e,n)>=t,o,r)}get(e){return this.swatches[e]||this.swatches[Object(i.a)(e,0,this.lastIndex)]}closestIndexOf(e){if(this.closestIndexCache.has(e.relativeLuminance))return this.closestIndexCache.get(e.relativeLuminance);let t=this.swatches.indexOf(e);if(-1!==t)return this.closestIndexCache.set(e.relativeLuminance,t),t;const n=this.swatches.reduce((t,n)=>Math.abs(n.relativeLuminance-e.relativeLuminance)<Math.abs(t.relativeLuminance-e.relativeLuminance)?n:t);return t=this.swatches.indexOf(n),this.closestIndexCache.set(e.relativeLuminance,t),t}static saturationBump(e,t){const n=Object(r.f)(e).s,a=Object(r.f)(t);if(a.s<n){const e=new o.a(a.h,n,a.l);return Object(r.b)(e)}return t}static ramp(e){const t=e/100;return t>.5?(t-.5)/.5:2*t}static createHighResolutionPalette(e){const t=[],n=Object(r.h)(s.a.fromObject(e).roundToPrecision(4)),a=Object(r.d)(new c.a(0,n.a,n.b)).clamp().roundToPrecision(4),i=Object(r.d)(new c.a(50,n.a,n.b)).clamp().roundToPrecision(4),o=Object(r.d)(new c.a(100,n.a,n.b)).clamp().roundToPrecision(4),u=new s.a(0,0,0),f=new s.a(1,1,1),p=o.equalValue(f)?0:14,m=a.equalValue(u)?0:14;for(let e=100+p;e>=0-m;e-=.5){let n;n=e<0?d(e/m+1,u,a):e<=50?d(h.ramp(e),a,i):e<=100?d(h.ramp(e),i,o):d((e-100)/p,o,f),n=h.saturationBump(i,n).roundToPrecision(4),t.push(l.a.from(n))}return new h(e,t)}static adjustEnd(e,t,n,a){const r=-1===a?t.swatches:t.reversedSwatches,o=e=>{const n=t.closestIndexOf(e);return 1===a?t.lastIndex-n:n};1===a&&n.reverse();const s=e(n[n.length-2]);if(Object(i.i)(Object(p.a)(n[n.length-1],n[n.length-2]),2)<s){n.pop();const e=o(t.colorContrast(r[t.lastIndex],s,void 0,a))-o(n[n.length-2]);let i=1;for(let a=n.length-e-1;a<n.length;a++){const e=o(n[a]),s=a===n.length-1?t.lastIndex:e+i;n[a]=r[s],i++}}1===a&&n.reverse()}static createColorPaletteByContrast(e,t){const n=h.createHighResolutionPalette(e),a=e=>{const n=t.stepContrast+t.stepContrast*(1-e.relativeLuminance)*t.stepContrastRamp;return Object(i.i)(n,2)},r=[];let o=t.preserveSource?e:n.swatches[0];r.push(o);do{const e=a(o);o=n.colorContrast(o,e,void 0,1),r.push(o)}while(o.relativeLuminance>0);if(t.preserveSource){o=e;do{const e=a(o);o=n.colorContrast(o,e,void 0,-1),r.unshift(o)}while(o.relativeLuminance<1)}return this.adjustEnd(a,n,r,-1),t.preserveSource&&this.adjustEnd(a,n,r,1),r}static from(e,t){const n=void 0===t?m:Object.assign(Object.assign({},m),t);return new h(e,Object.freeze(h.createColorPaletteByContrast(e,n)))}}},i2HT:function(e,t,n){"use strict";n.d(t,"a",function(){return o}),n.d(t,"b",function(){return s});var a=n("6eT7"),i=n("swXE"),r=n("THNU");const o=Object.freeze({create:(e,t,n)=>new c(e,t,n),from:e=>new c(e.r,e.g,e.b)});function s(e){const t={r:0,g:0,b:0,toColorString:()=>"",contrast:()=>0,relativeLuminance:0};for(const n in t)if(typeof t[n]!=typeof e[n])return!1;return!0}class c extends a.a{constructor(e,t,n){super(e,t,n,1),this.toColorString=this.toStringHexRGB,this.contrast=r.a.bind(null,this),this.createCSS=this.toColorString,this.relativeLuminance=Object(i.j)(this)}static fromObject(e){return new c(e.r,e.g,e.b)}}},icq5:function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("YZrM");const i=()=>{const e=[];return a.a.sections.files&&e.push("Sites.Read.All"),a.a.sections.mailMessages&&(e.push("Mail.Read"),e.push("Mail.ReadBasic")),a.a.sections.organization&&(e.push("User.Read.All"),a.a.sections.organization.showWorksWith&&e.push("People.Read.All")),a.a.sections.profile&&e.push("User.Read.All"),a.a.useContactApis&&e.push("Contacts.Read"),e.indexOf("User.Read.All")<0&&(e.push("User.ReadBasic.All"),e.push("User.Read")),e.indexOf("People.Read.All")<0&&e.push("People.Read"),a.a.isSendMessageVisible&&e.push("Chat.ReadWrite"),[...new Set(e)]}},imsm:function(e,t,n){"use strict";n.d(t,"a",function(){return a});const a={conversation:{name:"conversation",stores:{chats:"chats",subscriptions:"subscriptions",lastRead:"lastRead"},version:3,indexes:{subscriptions:[{name:"lastAccessDateTime",field:"lastAccessDateTime"}]}},presence:{name:"presence",stores:{presence:"presence"},version:2},users:{name:"users",stores:{users:"users",usersQuery:"usersQuery",userFilters:"userFilters"},version:3},photos:{name:"photos",stores:{contacts:"contacts",users:"users",groups:"groups",teams:"teams"},version:2},people:{name:"people",stores:{contacts:"contacts",groupPeople:"groupPeople",peopleQuery:"peopleQuery"},version:3},groups:{name:"groups",stores:{groups:"groups",groupsQuery:"groupsQuery"},version:5},get:{name:"responses",stores:{responses:"responses"},version:2},search:{name:"search",stores:{responses:"responses"},version:2},files:{name:"files",stores:{driveFiles:"driveFiles",groupFiles:"groupFiles",siteFiles:"siteFiles",userFiles:"userFiles",insightFiles:"insightFiles",fileQueries:"fileQueries"},version:2},fileLists:{name:"file-lists",stores:{fileLists:"fileLists",insightfileLists:"insightfileLists"},version:2}}},j9Xn:function(e,t,n){"use strict";var a;n.d(t,"a",function(){return a}),function(e){e.Canvas="Canvas",e.CanvasText="CanvasText",e.LinkText="LinkText",e.VisitedText="VisitedText",e.ActiveText="ActiveText",e.ButtonFace="ButtonFace",e.ButtonText="ButtonText",e.Field="Field",e.FieldText="FieldText",e.Highlight="Highlight",e.HighlightText="HighlightText",e.GrayText="GrayText"}(a||(a={}))},kd7Q:function(e,t,n){"use strict";n.d(t,"a",function(){return x});var a=n("eftJ"),i=n("QBS5"),r=n("oePG"),o=n("Gy7L"),s=n("o87Z"),c=n("TDEi");class d extends c.a{}class l extends(Object(s.a)(d)){constructor(){super(...arguments),this.proxy=document.createElement("input")}}class u extends l{constructor(){super(),this.initialValue="on",this.indeterminate=!1,this.keypressHandler=e=>{this.readOnly||e.key!==o.m||(this.indeterminate&&(this.indeterminate=!1),this.checked=!this.checked)},this.clickHandler=e=>{this.disabled||this.readOnly||(this.indeterminate&&(this.indeterminate=!1),this.checked=!this.checked)},this.proxy.setAttribute("type","checkbox")}readOnlyChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.readOnly=this.readOnly)}}Object(a.a)([Object(i.c)({attribute:"readonly",mode:"boolean"})],u.prototype,"readOnly",void 0),Object(a.a)([r.d],u.prototype,"defaultSlottedNodes",void 0),Object(a.a)([r.d],u.prototype,"indeterminate",void 0);var f=n("6BDD"),p=n("UauI"),m=n("4X57"),_=n("xY0q"),h=n("wHpb"),b=n("2i1/"),g=n("oqLQ"),v=n("j9Xn"),y=n("QkLF"),S=n("8hiW"),D=n("NLdk"),I=n("FVZ7");const x=u.compose({baseName:"checkbox",template:(e,t)=>f.a`
    <template
        role="checkbox"
        aria-checked="${e=>e.checked}"
        aria-required="${e=>e.required}"
        aria-disabled="${e=>e.disabled}"
        aria-readonly="${e=>e.readOnly}"
        tabindex="${e=>e.disabled?null:0}"
        @keypress="${(e,t)=>e.keypressHandler(t.event)}"
        @click="${(e,t)=>e.clickHandler(t.event)}"
        class="${e=>e.readOnly?"readonly":""} ${e=>e.checked?"checked":""} ${e=>e.indeterminate?"indeterminate":""}"
    >
        <div part="control" class="control">
            <slot name="checked-indicator">
                ${t.checkedIndicator||""}
            </slot>
            <slot name="indeterminate-indicator">
                ${t.indeterminateIndicator||""}
            </slot>
        </div>
        <label
            part="label"
            class="${e=>e.defaultSlottedNodes&&e.defaultSlottedNodes.length?"label":"label label__hidden"}"
        >
            <slot ${Object(p.a)("defaultSlottedNodes")}></slot>
        </label>
    </template>
`,styles:(e,t)=>m.a`
    ${Object(_.a)("inline-flex")} :host {
      align-items: center;
      outline: none;
      ${""} user-select: none;
    }

    .control {
      position: relative;
      width: calc((${y.a} / 2 + ${S.s}) * 1px);
      height: calc((${y.a} / 2 + ${S.s}) * 1px);
      box-sizing: border-box;
      border-radius: calc(${S.q} * 1px);
      border: calc(${S.vb} * 1px) solid ${S.ub};
      background: ${S.K};
      cursor: pointer;
    }

    .label__hidden {
      display: none;
      visibility: hidden;
    }

    .label {
      ${D.a}
      color: ${S.fb};
      ${""} padding-inline-start: calc(${S.s} * 2px + 2px);
      margin-inline-end: calc(${S.s} * 2px + 2px);
      cursor: pointer;
    }

    slot[name='checked-indicator'],
    slot[name='indeterminate-indicator'] {
      display: flex;
      align-items: center;
      justify-content: center;
      width: 100%;
      height: 100%;
      fill: ${S.fb};
      opacity: 0;
      pointer-events: none;
    }

    slot[name='indeterminate-indicator'] {
      position: absolute;
      top: 0;
    }

    :host(.checked) slot[name='checked-indicator'],
    :host(.checked) slot[name='indeterminate-indicator'] {
      fill: ${S.C};
    }

    :host(:not(.disabled):hover) .control {
      background: ${S.J};
      border-color: ${S.tb};
    }

    :host(:not(.disabled):active) .control {
      background: ${S.H};
      border-color: ${S.sb};
    }

    :host(:${h.a}) .control {
      background: ${S.I};
      ${I.b}
    }

    :host(.checked) .control {
      background: ${S.e};
      border-color: transparent;
    }

    :host(.checked:not(.disabled):hover) .control {
      background: ${S.d};
      border-color: transparent;
    }

    :host(.checked:not(.disabled):active) .control {
      background: ${S.b};
      border-color: transparent;
    }

    :host(.disabled) .label,
    :host(.readonly) .label,
    :host(.readonly) .control,
    :host(.disabled) .control {
      cursor: ${b.a};
    }

    :host(.checked:not(.indeterminate)) slot[name='checked-indicator'],
    :host(.indeterminate) slot[name='indeterminate-indicator'] {
      opacity: 1;
    }

    :host(.disabled) {
      opacity: ${S.u};
    }
  `.withBehaviors(Object(g.a)(m.a`
        .control {
          border-color: ${v.a.FieldText};
          background: ${v.a.Field};
        }
        :host(:not(.disabled):hover) .control,
        :host(:not(.disabled):active) .control {
          border-color: ${v.a.Highlight};
          background: ${v.a.Field};
        }
        slot[name='checked-indicator'],
        slot[name='indeterminate-indicator'] {
          fill: ${v.a.FieldText};
        }
        :host(:${h.a}) .control {
          forced-color-adjust: none;
          outline-color: ${v.a.FieldText};
          background: ${v.a.Field};
          border-color: ${v.a.Highlight};
        }
        :host(.checked) .control {
          background: ${v.a.Highlight};
          border-color: ${v.a.Highlight};
        }
        :host(.checked:not(.disabled):hover) .control,
        :host(.checked:not(.disabled):active) .control {
          background: ${v.a.HighlightText};
          border-color: ${v.a.Highlight};
        }
        :host(.checked) slot[name='checked-indicator'],
        :host(.checked) slot[name='indeterminate-indicator'] {
          fill: ${v.a.HighlightText};
        }
        :host(.checked:hover ) .control slot[name='checked-indicator'],
        :host(.checked:hover ) .control slot[name='indeterminate-indicator'] {
          fill: ${v.a.Highlight};
        }
        :host(.disabled) {
          opacity: 1;
        }
        :host(.disabled) .control {
          border-color: ${v.a.GrayText};
          background: ${v.a.Field};
        }
        :host(.disabled) slot[name='checked-indicator'],
        :host(.checked.disabled:hover) .control slot[name='checked-indicator'],
        :host(.disabled) slot[name='indeterminate-indicator'],
        :host(.checked.disabled:hover) .control slot[name='indeterminate-indicator'] {
          fill: ${v.a.GrayText};
        }
      `)),checkedIndicator:'\n    <svg width="16" height="16" xmlns="http://www.w3.org/2000/svg">\n      <path d="M13.86 3.66a.5.5 0 01-.02.7l-7.93 7.48a.6.6 0 01-.84-.02L2.4 9.1a.5.5 0 01.72-.7l2.4 2.44 7.65-7.2a.5.5 0 01.7.02z"/>\n    </svg>\n  ',indeterminateIndicator:'\n    <svg width="16" height="16" xmlns="http://www.w3.org/2000/svg">\n      <path d="M3 8c0-.28.22-.5.5-.5h9a.5.5 0 010 1h-9A.5.5 0 013 8z"/>\n    </svg>\n  '})},kpPY:function(e,t,n){"use strict";n.d(t,"a",function(){return i});let a=0;function i(e=""){return`${e}${a++}`}},m1Vi:function(e,t,n){"use strict";n.d(t,"a",function(){return g});var a=n("3rEL"),i=n("QBS5"),r=n("Mjrb"),o=n("6BDD"),s=n("+53S"),c=n("UauI"),d=n("C5kU"),l=n("4X57"),u=n("2i1/"),f=n("81Gc"),p=n("57ob"),m=n("8hiW");const _=":not([disabled])",h="[disabled]";class b extends r.a{appearanceChanged(e,t){e!==t&&(this.classList.add(t),this.classList.remove(e))}connectedCallback(){super.connectedCallback(),this.appearance||(this.appearance="neutral")}defaultSlottedContentChanged(){const e=this.defaultSlottedContent.filter(e=>e.nodeType===Node.ELEMENT_NODE);1===e.length&&e[0]instanceof SVGElement?this.control.classList.add("icon-only"):this.control.classList.remove("icon-only")}}Object(a.a)([i.c],b.prototype,"appearance",void 0);const g=b.compose({baseName:"button",baseClass:r.a,template:(e,t)=>o.a`
    <button
        class="control"
        part="control"
        ?autofocus="${e=>e.autofocus}"
        ?disabled="${e=>e.disabled}"
        form="${e=>e.formId}"
        formaction="${e=>e.formaction}"
        formenctype="${e=>e.formenctype}"
        formmethod="${e=>e.formmethod}"
        formnovalidate="${e=>e.formnovalidate}"
        formtarget="${e=>e.formtarget}"
        name="${e=>e.name}"
        type="${e=>e.type}"
        value="${e=>e.value}"
        aria-atomic="${e=>e.ariaAtomic}"
        aria-busy="${e=>e.ariaBusy}"
        aria-controls="${e=>e.ariaControls}"
        aria-current="${e=>e.ariaCurrent}"
        aria-describedby="${e=>e.ariaDescribedby}"
        aria-details="${e=>e.ariaDetails}"
        aria-disabled="${e=>e.ariaDisabled}"
        aria-errormessage="${e=>e.ariaErrormessage}"
        aria-expanded="${e=>e.ariaExpanded}"
        aria-flowto="${e=>e.ariaFlowto}"
        aria-haspopup="${e=>e.ariaHaspopup}"
        aria-hidden="${e=>e.ariaHidden}"
        aria-invalid="${e=>e.ariaInvalid}"
        aria-keyshortcuts="${e=>e.ariaKeyshortcuts}"
        aria-label="${e=>e.ariaLabel}"
        aria-labelledby="${e=>e.ariaLabelledby}"
        aria-live="${e=>e.ariaLive}"
        aria-owns="${e=>e.ariaOwns}"
        aria-pressed="${e=>e.ariaPressed}"
        aria-relevant="${e=>e.ariaRelevant}"
        aria-roledescription="${e=>e.ariaRoledescription}"
        ${Object(s.a)("control")}
    >
        ${Object(d.d)(e,t)}
        <span class="content" part="content">
            <slot ${Object(c.a)("defaultSlottedContent")}></slot>
        </span>
        ${Object(d.b)(e,t)}
    </button>
`,styles:(e,t)=>l.a`
    :host(${_}) .control {
      cursor: pointer;
    }

    :host(${h}) .control {
      cursor: ${u.a};
    }

    @media (forced-colors: none) {
      :host(${h}) .control {
        opacity: ${m.u};
      }
    }

    ${Object(f.f)(e,t,_,h)}
  `.withBehaviors(Object(p.a)("neutral",Object(f.c)(e,t,_,h)),Object(p.a)("accent",Object(f.a)(e,t,_,h)),Object(p.a)("lightweight",Object(f.b)(e,t,_,h)),Object(p.a)("outline",Object(f.d)(e,t,_,h)),Object(p.a)("stealth",Object(f.e)(e,t,_,h))),shadowOptions:{delegatesFocus:!0}})},mwqp:function(e,t,n){"use strict";n.r(t),n.d(t,"MgtLibrary",function(){return a}),n.d(t,"SharePointProvider",function(){return c}),n.d(t,"registerMgtComponents",function(){return pr}),n.d(t,"registerMgtAgendaComponent",function(){return B}),n.d(t,"MgtAgenda",function(){return j}),n.d(t,"registerMgtFileComponent",function(){return bi.b}),n.d(t,"MgtFile",function(){return bi.a}),n.d(t,"registerMgtFileListComponent",function(){return gi.b}),n.d(t,"MgtFileList",function(){return gi.a}),n.d(t,"registerMgtPickerComponent",function(){return li}),n.d(t,"MgtPicker",function(){return ui}),n.d(t,"registerMgtTaxonomyPickerComponent",function(){return Ii}),n.d(t,"MgtTaxonomyPicker",function(){return xi}),n.d(t,"isCollectionResponse",function(){return Q}),n.d(t,"ResponseType",function(){return Y}),n.d(t,"registerMgtGetComponent",function(){return X}),n.d(t,"MgtGet",function(){return Z}),n.d(t,"registerMgtLoginComponent",function(){return Fe}),n.d(t,"MgtLogin",function(){return He}),n.d(t,"GroupType",function(){return Re}),n.d(t,"PersonType",function(){return I.a}),n.d(t,"UserType",function(){return I.b}),n.d(t,"registerMgtPeoplePickerComponent",function(){return et}),n.d(t,"MgtPeoplePicker",function(){return tt}),n.d(t,"PersonCardInteraction",function(){return w.a}),n.d(t,"registerMgtPeopleComponent",function(){return M}),n.d(t,"MgtPeople",function(){return P}),n.d(t,"MgtPersonCardConfig",function(){return mr.a}),n.d(t,"getMgtPersonCardScopes",function(){return _r.a}),n.d(t,"registerMgtPersonCardComponent",function(){return l.registerMgtPersonCardComponent}),n.d(t,"MgtPersonCard",function(){return l.MgtPersonCard}),n.d(t,"defaultPersonProperties",function(){return d.b}),n.d(t,"registerMgtPersonComponent",function(){return d.c}),n.d(t,"MgtPerson",function(){return d.a}),n.d(t,"PersonViewType",function(){return ae.a}),n.d(t,"avatarType",function(){return ae.b}),n.d(t,"TasksSource",function(){return On}),n.d(t,"registerMgtTasksComponent",function(){return Pn}),n.d(t,"MgtTasks",function(){return Tn}),n.d(t,"registerMgtTeamsChannelPickerComponent",function(){return sa}),n.d(t,"MgtTeamsChannelPicker",function(){return ca}),n.d(t,"registerMgtTodoComponent",function(){return _i}),n.d(t,"MgtTodo",function(){return hi}),n.d(t,"registerMgtContactComponent",function(){return hr.b}),n.d(t,"MgtContact",function(){return hr.a}),n.d(t,"registerMgtMessagesComponent",function(){return br.b}),n.d(t,"MgtMessages",function(){return br.a}),n.d(t,"registerMgtOrganizationComponent",function(){return gr.b}),n.d(t,"MgtOrganization",function(){return gr.a}),n.d(t,"registerMgtProfileComponent",function(){return vr.b}),n.d(t,"MgtProfile",function(){return vr.a}),n.d(t,"registerMgtThemeToggleComponent",function(){return Ni}),n.d(t,"MgtThemeToggle",function(){return Bi}),n.d(t,"registerMgtSpinnerComponent",function(){return Ye.b}),n.d(t,"MgtSpinner",function(){return Ye.a}),n.d(t,"registerMgtSearchBoxComponent",function(){return nr}),n.d(t,"MgtSearchBox",function(){return ar}),n.d(t,"registerMgtSearchResultsComponent",function(){return ur}),n.d(t,"MgtSearchResults",function(){return fr}),n.d(t,"ViewType",function(){return ee.a}),n.d(t,"getUserWithPhoto",function(){return ne.a}),n.d(t,"getPhotoInvalidationTime",function(){return z.g}),n.d(t,"getIsPhotosCacheEnabled",function(){return z.c}),n.d(t,"getPhotoForResource",function(){return z.e}),n.d(t,"getContactPhoto",function(){return z.a}),n.d(t,"getUserPhoto",function(){return z.h}),n.d(t,"myPhoto",function(){return z.i}),n.d(t,"getPersonImage",function(){return z.d}),n.d(t,"getGroupImage",function(){return z.b}),n.d(t,"getPhotoFromCache",function(){return z.f}),n.d(t,"storePhotoInCache",function(){return z.j}),n.d(t,"getUserPresence",function(){return x.a}),n.d(t,"getUsersPresenceByPeople",function(){return x.b}),n.d(t,"schemas",function(){return K.a}),n.d(t,"applyTheme",function(){return Ui}),n.d(t,"isGroup",function(){return yr.b}),n.d(t,"isUser",function(){return yr.c}),n.d(t,"isContact",function(){return yr.a}),n.d(t,"getRelativeDisplayDate",function(){return Ke.l}),n.d(t,"getDateString",function(){return Ke.f}),n.d(t,"getShortDateString",function(){return Ke.n}),n.d(t,"getMonthString",function(){return Ke.j}),n.d(t,"getDayOfWeekString",function(){return Ke.g}),n.d(t,"getDaysInMonth",function(){return Ke.h}),n.d(t,"getDateFromMonthYear",function(){return Ke.e}),n.d(t,"debounce",function(){return Ke.b}),n.d(t,"blobToBase64",function(){return Ke.a}),n.d(t,"extractEmailAddress",function(){return Ke.c}),n.d(t,"isValidEmail",function(){return Ke.o}),n.d(t,"formatBytes",function(){return Ke.d}),n.d(t,"sanitizeSummary",function(){return Ke.p}),n.d(t,"trimFileExtension",function(){return Ke.q}),n.d(t,"getNameFromUrl",function(){return Ke.k}),n.d(t,"getResponseInvalidationTime",function(){return Ke.m}),n.d(t,"getIsResponseCacheEnabled",function(){return Ke.i}),n.d(t,"BatchResponse",function(){return Sr.a}),n.d(t,"MICROSOFT_GRAPH_DEFAULT_ENDPOINT",function(){return Sr.l}),n.d(t,"MICROSOFT_GRAPH_ENDPOINTS",function(){return Sr.m}),n.d(t,"Graph",function(){return Sr.g}),n.d(t,"createFromProvider",function(){return Sr.B}),n.d(t,"BetaGraph",function(){return Sr.b}),n.d(t,"ComponentMediaQuery",function(){return Sr.e}),n.d(t,"MgtBaseComponent",function(){return Sr.n}),n.d(t,"MgtBaseProvider",function(){return Sr.o}),n.d(t,"MgtTemplatedComponent",function(){return Sr.p}),n.d(t,"customElementHelper",function(){return Sr.D}),n.d(t,"IProvider",function(){return Sr.i}),n.d(t,"LoginType",function(){return Sr.k}),n.d(t,"ProviderState",function(){return Sr.s}),n.d(t,"Providers",function(){return Sr.t}),n.d(t,"ProvidersChangedState",function(){return Sr.u}),n.d(t,"SimpleProvider",function(){return Sr.v}),n.d(t,"dbListKey",function(){return Sr.E}),n.d(t,"CacheService",function(){return Sr.c}),n.d(t,"CacheStore",function(){return Sr.d}),n.d(t,"EventDispatcher",function(){return Sr.f}),n.d(t,"equals",function(){return Sr.F}),n.d(t,"arraysAreEqual",function(){return Sr.y}),n.d(t,"chainMiddleware",function(){return Sr.A}),n.d(t,"prepScopes",function(){return Sr.J}),n.d(t,"validateBaseURL",function(){return Sr.L}),n.d(t,"TeamsHelper",function(){return Sr.w}),n.d(t,"TemplateHelper",function(){return Sr.x}),n.d(t,"GraphPageIterator",function(){return Sr.h}),n.d(t,"LocalizationHelper",function(){return Sr.j}),n.d(t,"mgtHtml",function(){return Sr.I}),n.d(t,"customElement",function(){return Sr.C}),n.d(t,"log",function(){return Sr.H}),n.d(t,"warn",function(){return Sr.M}),n.d(t,"error",function(){return Sr.G}),n.d(t,"buildComponentName",function(){return Sr.z}),n.d(t,"registerComponent",function(){return Sr.K}),n.d(t,"PACKAGE_VERSION",function(){return Sr.r}),n.d(t,"MockProvider",function(){return Sr.q});class a{name(){return"MgtLibrary"}}var i=n("/i08"),r=n("/49y"),o=n("SUtl"),s=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};class c extends i.a{get provider(){return this._provider}get isLoggedIn(){return!!this._idToken}get name(){return"MgtSharePointProvider"}constructor(e,t=r.a){super(),e.aadTokenProviderFactory.getTokenProvider().then(e=>{this._provider=e,this.baseURL=t,this.graph=Object(o.b)(this),this.internalLogin()})}getAccessToken(){return s(this,void 0,void 0,function*(){const e=this.baseURL?this.baseURL:r.a;return yield this.provider.getToken(e)})}updateScopes(e){this.scopes=e}internalLogin(){return s(this,void 0,void 0,function*(){this._idToken=yield this.getAccessToken(),this.setState(this._idToken?i.c.SignedIn:i.c.SignedOut)})}}var d=n("7SX6"),l=n("uVfA"),u=n("qqMp"),f=n("ZgG/"),p=n("Y1A4"),m=n("cBsD"),_=n("zFbe");const h=[u.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{--card-height:auto;--card-width:99%;background-color:var(--agenda-background-color,transparent)}:host .header{margin:var(--agenda-header-margin,18px 0 12px 10px);font-size:var(--agenda-header-font-size,24px);font-style:normal;font-weight:400;line-height:32px;color:var(--agenda-header-color,var(--neutral-foreground-rest));opacity:.9}:host .agenda,:host .group{display:flex;flex-direction:column;row-gap:var(--agenda-event-row-gap,14px)}:host .agenda>.group:first-child>.header,:host .group>.group:first-child>.header{margin-top:0}:host .agenda .event,:host .group .event{background:var(--agenda-event-background-color,var(--fill-color));border:var(--agenda-event-border,solid 2px transparent);box-shadow:var(--agenda-event-box-shadow,var(--elevation-shadow-card-rest));padding:var(--agenda-event-padding,12px);position:relative;display:flex;flex:1 1 auto;content-visibility:visible;contain:none}:host .agenda .event-container,:host .group .event-container{border-radius:calc(var(--layer-corner-radius) * 1px);padding:1px}:host .agenda .event.narrow,:host .group .event.narrow{display:flex;flex-direction:column;inset:0}:host .agenda .event-time-container,:host .group .event-time-container{font-style:normal;font-weight:600;font-size:12px;color:var(--agenda-event-time-color,var(--neutral-foreground-rest));width:112px;height:16px}:host .agenda .event-time-container.narrow,:host .group .event-time-container.narrow{margin-bottom:1px;width:100%}:host .agenda .event-time,:host .group .event-time{font-size:var(--agenda-event-time-font-size,12px);color:var(--agenda-event-time-color,var(--neutral-foreground-rest));font-weight:600}:host .agenda .event-details-container,:host .group .event-details-container{display:flex;flex-direction:column;position:relative;bottom:8px;top:0;padding-inline-start:32px}:host .agenda .event-details-container.narrow,:host .group .event-details-container.narrow{position:inherit;left:6px;display:flex;flex-direction:column;padding-inline-start:0}:host .agenda .event-subject,:host .group .event-subject{font-style:normal;font-weight:400;font-size:var(--agenda-event-subject-font-size,20px);line-height:28px;color:var(--agenda-event-subject-color,var(--neutral-foreground-rest));mix-blend-mode:normal;position:inherit;bottom:8px}:host .agenda .event-location-container,:host .group .event-location-container{display:inline-flex;flex-direction:row}:host .agenda .event-location-container .event-location,:host .group .event-location-container .event-location{padding-inline-start:3px;font-style:normal;font-weight:400;font-size:var(--agenda-event-location-font-size,12px);line-height:16px;color:var(--agenda-event-location-color,var(--neutral-foreground-rest))}:host .agenda .event-location-container .event-location-loading,:host .group .event-location-container .event-location-loading{width:90px;height:10px;margin:2px 0 0 4px}:host .agenda .event-location-container .event-location-icon,:host .group .event-location-container .event-location-icon{display:inline-flex}:host .agenda .event-location-container .event-location-icon svg,:host .group .event-location-container .event-location-icon svg{position:relative;top:2px;width:12px;height:12px}:host .agenda .event-location-container .event-location-icon svg path,:host .group .event-location-container .event-location-icon svg path{stroke:var(--agenda-event-location-color,var(--neutral-foreground-rest))}:host .agenda .event-location-container .event-location-icon-loading,:host .group .event-location-container .event-location-icon-loading{width:14px;height:14px}:host .agenda .event-location-container .event-attendee-loading,:host .group .event-location-container .event-attendee-loading{width:20px;height:20px;border-radius:10px;margin:0 2px 0 0}:host .agenda .event-attendees,:host .group .event-attendees{--list-margin:8px 0 0 0;--avatar-size-s:20px}fluent-card.event.event-loading{--card-height:90px}:host .event-attendees{--color:$agenda-event-attendees-color}:host fluent-tooltip{width:auto;contain:inline-size}[dir=rtl] :host{direction:rtl}[dir=rtl] .event-time-container{direction:ltr;display:flex;justify-content:flex-end}@media (forced-colors:active) and (prefers-color-scheme:dark){:host .agenda .event-location-container .event-location-icon svg path{stroke:#fff!important}}@media (forced-colors:active) and (prefers-color-scheme:light){:host .agenda .event-location-container .event-location-icon svg path{stroke:#000!important}}
`];var b=n("RIOo"),g=n("WHff"),v=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const y=(e,t,n="calendars.read")=>v(void 0,void 0,void 0,function*(){const a=e.api(t).middlewareOptions(Object(b.a)(n)).orderby("start/dateTime");return g.a.create(e,a)});var S=n("7Cdu"),D=n("V07F"),I=n("z0DP"),x=n("zlIh"),C=n("s9K1"),O=n("ylMe"),w=n("zCAG");const E=[u.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host .people-list{list-style:none;margin:var(--people-list-margin,8px 4px 8px 8px);padding:unset;display:flex;align-items:center;gap:var(--people-avatar-gap,4px)}:host .overflow span{vertical-align:middle;color:var(--people-overflow-font-color,currentColor);font-size:var(--people-overflow-font-size,12px);font-weight:var(--people-overflow-font-weight,400)}
`];var A=n("0mOt"),L=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},k=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};const M=()=>{Object(d.c)(),Object(A.b)("people",P)};class P extends p.a{static get styles(){return E}get groupId(){return this._groupId}set groupId(e){this._groupId!==e&&(this._groupId=e,this.requestStateUpdate(!0))}get userIds(){return this._userIds}set userIds(e){Object(O.a)(this._userIds,e)||(this._userIds=e,this.requestStateUpdate(!0))}get peopleQueries(){return this._peopleQueries}set peopleQueries(e){Object(O.a)(this._peopleQueries,e)||(this._peopleQueries=e,this.requestStateUpdate(!0))}get resource(){return this._resource}set resource(e){this._resource!==e&&(this._resource=e,this.requestStateUpdate(!0))}get version(){return this._version}set version(e){this._version!==e&&(this._version=e,this.requestStateUpdate(!0))}get fallbackDetails(){return this._fallbackDetails}set fallbackDetails(e){e!==this._fallbackDetails&&(this._fallbackDetails=e,this.requestStateUpdate())}static get requiredScopes(){return[...new Set(["user.read.all","people.read","user.readbasic.all","presence.read.all","contacts.read",...d.a.requiredScopes])]}constructor(){super(),this.personCardInteraction=w.a.hover,this.scopes=[],this._peoplePresence={},this._version="v1.0",this._arrowKeyLocation=-1,this.handleKeyDown=e=>{const t=this.shadowRoot.querySelector(".people-list");let n;const a=null==t?void 0:t.children;for(const e of a){const t=e;t.setAttribute("tabindex","-1"),t.blur()}const i=t.childElementCount,r=e.key;if("ArrowRight"===r)this._arrowKeyLocation=(this._arrowKeyLocation+1+i)%i;else if("ArrowLeft"===r)this._arrowKeyLocation=(this._arrowKeyLocation-1+i)%i;else if("Tab"===r||"Escape"===r)this._arrowKeyLocation=-1,t.blur();else if(["Enter","space"," "].includes(r)&&this.personCardInteraction!==w.a.none){const e=a[this._arrowKeyLocation].querySelector("mgt-person");e&&e.showPersonCard()}this._arrowKeyLocation>-1&&(n=a[this._arrowKeyLocation],n.setAttribute("tabindex","1"),n.focus())},this.showMax=3}clearState(){this.people=null}requestStateUpdate(e){return e&&(this.people=null),super.requestStateUpdate(e)}render(){return this.isLoadingState?this.renderLoading():this.people&&0!==this.people.length?this.renderTemplate("default",{people:this.people,max:this.showMax})||this.renderPeople():this.renderNoData()}renderLoading(){return this.renderTemplate("loading",null)||u.c``}renderPeople(){const e=this.people.slice(0,this.showMax).filter(e=>e);return u.c`
      <ul
        tabindex="0"
        class="people-list"
        aria-label="people"
        @keydown=${this.handleKeyDown}>
        ${Object(D.a)(e,e=>e.id?e.id:e.displayName,e=>u.c`
            <li tabindex="-1" class="people-person">
              ${this.renderPerson(e)}
            </li>
          `)}
        ${this.people.length>this.showMax?this.renderOverflow():null}
      </ul>
    `}renderOverflow(){const e=this.people.length-this.showMax;return this.renderTemplate("overflow",{extra:e,max:this.showMax,people:this.people})||u.c`
        <li tabindex="-1" aria-label="and ${e} more attendees" class="overflow"><span>+${e}</span></li>
      `}renderPerson(e){let t={activity:"Offline",availability:"Offline",id:null};return this.showPresence&&this._peoplePresence&&(t=this._peoplePresence[e.id]),this.renderTemplate("person",{person:e},e.id)||m.a`
        <mgt-person
          .personDetails=${e}
          .fetchImage=${!0}
          .avatarSize=${"small"}
          .personCardInteraction=${this.personCardInteraction}
          .showPresence=${this.showPresence}
          .personPresence=${t}
          .usage=${"people"}
        ></mgt-person>
      `}renderNoData(){return this.renderTemplate("no-data",null)||u.c``}loadState(){return e=this,void 0,n=function*(){if(!this.people){const e=_.a.globalProvider;if(e&&e.state===i.c.SignedIn){const t=e.graph.forComponent(this);this.groupId?this.people=yield Object(C.a)(t,null,this.groupId,this.showMax,I.a.person):this.userIds||this.peopleQueries?this.people=this.userIds?yield Object(C.j)(t,this.userIds,"","",this._fallbackDetails):yield Object(C.i)(t,this.peopleQueries,this._fallbackDetails):this.resource?this.people=yield Object(I.g)(t,this.version,this.resource,this.scopes):this.people=yield Object(I.f)(t),this.showPresence?this._peoplePresence=yield Object(x.b)(t,this.people):this._peoplePresence=null}}},new((t=void 0)||(t=Promise))(function(a,i){function r(e){try{s(n.next(e))}catch(e){i(e)}}function o(e){try{s(n.throw(e))}catch(e){i(e)}}function s(e){var n;e.done?a(e.value):(n=e.value,n instanceof t?n:new t(function(e){e(n)})).then(r,o)}s((n=n.apply(e,[])).next())});var e,t,n}}L([Object(f.b)({attribute:"group-id",type:String}),k("design:type",String),k("design:paramtypes",[Object])],P.prototype,"groupId",null),L([Object(f.b)({attribute:"user-ids",converter:(e,t)=>e.split(",").map(e=>e.trim())}),k("design:type",Array),k("design:paramtypes",[Array])],P.prototype,"userIds",null),L([Object(f.b)({attribute:"people",type:Object}),k("design:type",Array)],P.prototype,"people",void 0),L([Object(f.b)({attribute:"people-queries",converter:(e,t)=>e.split(",").map(e=>e.trim())}),k("design:type",Array),k("design:paramtypes",[Array])],P.prototype,"peopleQueries",null),L([Object(f.b)({attribute:"show-max",type:Number}),k("design:type",Number)],P.prototype,"showMax",void 0),L([Object(f.b)({attribute:"show-presence",type:Boolean}),k("design:type",Boolean)],P.prototype,"showPresence",void 0),L([Object(f.b)({attribute:"person-card",converter:(e,t)=>(e=e.toLowerCase(),void 0===w.a[e]?w.a.hover:w.a[e])}),k("design:type",Number)],P.prototype,"personCardInteraction",void 0),L([Object(f.b)({attribute:"resource",type:String}),k("design:type",String),k("design:paramtypes",[Object])],P.prototype,"resource",null),L([Object(f.b)({attribute:"version",type:String}),k("design:type",String),k("design:paramtypes",[Object])],P.prototype,"version",null),L([Object(f.b)({attribute:"scopes",converter:e=>e?e.toLowerCase().split(","):null,reflect:!0}),k("design:type",Array)],P.prototype,"scopes",void 0),L([Object(f.b)({attribute:"fallback-details",type:Array}),k("design:type",Array),k("design:paramtypes",[Array])],P.prototype,"fallbackDetails",null),L([Object(f.c)(),k("design:type",Object)],P.prototype,"_arrowKeyLocation",void 0);var T=n("VDxL"),U=n("bica"),F=n("h2QR"),H=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},R=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)},N=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const B=()=>{Object(T.a)(U.a),M(),Object(A.b)("agenda",j)};class j extends p.a{constructor(){super(...arguments),this._days=3,this.onResize=()=>{this._isNarrow=this.offsetWidth<600}}static get styles(){return h}get date(){return this._date}set date(e){this._date!==e&&(this._date=e,this.reloadState())}get groupId(){return this._groupId}set groupId(e){this._groupId!==e&&(this._groupId=e,this.reloadState())}get days(){return this._days}set days(e){this._days!==e&&(this._days=e,this.reloadState())}get eventQuery(){return this._eventQuery}set eventQuery(e){this._eventQuery!==e&&(this._eventQuery=e,this.reloadState())}get preferredTimezone(){return this._preferredTimezone}set preferredTimezone(e){this._preferredTimezone!==e&&(this._preferredTimezone=e,this.reloadState())}static get requiredScopes(){return[...new Set(["calendars.read",...P.requiredScopes])]}connectedCallback(){this._isNarrow=this.offsetWidth<600,super.connectedCallback(),window.addEventListener("resize",this.onResize)}disconnectedCallback(){window.removeEventListener("resize",this.onResize),super.disconnectedCallback()}render(){if(!this.events&&this.isLoadingState)return this.renderLoading();if(!this.events||0===this.events.length)return this.renderNoData();const e=this.showMax&&this.showMax>0?this.events.slice(0,this.showMax):this.events,t=this.renderTemplate("default",{events:e});if(t)return t;const n={agenda:!0,grouped:this.groupByDay};return u.c`
      <div dir=${this.direction} class="${Object(F.a)(n)}">
        ${this.groupByDay?this.renderGroups(e):this.renderEvents(e)}
        ${this.isLoadingState?this.renderLoading():u.c``}
      </div>
    `}reload(){return N(this,void 0,void 0,function*(){this.events=yield this.loadEvents()})}renderLoading(){return this.renderTemplate("loading",null)||u.c`
        <fluent-card class="event event-loading">
          <div class="event-time-container">
            <div class="event-time-loading loading-element"></div>
          </div>
          <div class="event-details-container">
            <div class="event-subject-loading loading-element"></div>
            <div class="event-location-container">
              <div class="event-location-icon-loading loading-element"></div>
              <div class="event-location-loading loading-element"></div>
            </div>
            <div class="event-location-container">
              <div class="event-attendee-loading loading-element"></div>
              <div class="event-attendee-loading loading-element"></div>
              <div class="event-attendee-loading loading-element"></div>
            </div>
          </div>
        </fluent-card>`}clearState(){this.events=null}renderNoData(){return this.renderTemplate("no-data",null)||u.c``}renderEvent(e){this._isNarrow=this.offsetWidth<600;const t={narrow:this._isNarrow};return u.c`
      <fluent-card class="${Object(F.a)(Object.assign({event:!0},t))}">
        <div class="${Object(F.a)(Object.assign({"event-time-container":!0},t))}">
          <div class="event-time" aria-label="${this.getEventTimeString(e)}">${this.getEventTimeString(e)}</div>
        </div>
        <div class="${Object(F.a)(Object.assign({"event-details-container":!0},t))}">
          ${this.renderTitle(e)} ${this.renderLocation(e)} ${this.renderAttendees(e)}
        </div>
        <div class="event-other-container">${this.renderOther(e)}</div>
      </fluent-card>
    `}renderHeader(e){return this.renderTemplate("header",{header:e},"header-"+e)||u.c`
        <div class="header" aria-label="${e}">${e}</div>
      `}renderTitle(e){return u.c`
      <div
        aria-label=${e.subject}
        class="${Object(F.a)({"event-subject":!0,narrow:this._isNarrow})}"
      >
        ${e.subject}
      </div>`}renderLocation(e){return e.location.displayName?u.c`
      <div class="event-location-container">
        <div class="event-location-icon">${Object(S.b)(S.a.OfficeLocation)}</div>
        <div class="event-location" aria-label="${e.location.displayName}">${e.location.displayName}</div>
      </div>
    `:null}renderAttendees(e){return e.attendees.length?m.a`
      <mgt-people
        show-max="5"
        show-presence
        class="event-attendees"
        .peopleQueries=${e.attendees.map(e=>e.emailAddress.address)}
      ></mgt-people>
    `:null}renderOther(e){return this.hasTemplate("event-other")?u.c`
          ${this.renderTemplate("event-other",{event:e},e.id+"-other")}
        `:null}renderGroups(e){const t={};return e.forEach(e=>{var n;let a=null===(n=null==e?void 0:e.start)||void 0===n?void 0:n.dateTime;"UTC"===e.end.timeZone&&(a+="Z");const i=this.getDateHeaderFromDateTimeString(a);t[i]=t[i]||[],t[i].push(e)}),u.c`
      ${Object.keys(t).map(e=>u.c`
            <div class="group">${this.renderHeader(e)} ${this.renderEvents(t[e])}</div>
          `)}
    `}renderEvents(e){return u.c`
        ${e.map(e=>u.c`
              <div
                class="event-container"
                tabindex="0"
                @focus=${()=>this.eventClicked(e)}>
                ${this.renderTemplate("event",{event:e},e.id)||this.renderEvent(e)}
              </div>`)}`}loadState(){return N(this,void 0,void 0,function*(){if(this.events)return;const e=yield this.loadEvents();(null==e?void 0:e.length)>0&&(this.events=e)})}reloadState(){return N(this,void 0,void 0,function*(){this.events=null,yield this.requestStateUpdate(!0)})}eventClicked(e){this.fireCustomEvent("eventClick",e)}getEventTimeString(e){if(e.isAllDay)return"ALL DAY";let t=e.start.dateTime;"UTC"===e.start.timeZone&&(t+="Z");let n=e.end.dateTime;return"UTC"===e.end.timeZone&&(n+="Z"),`${this.prettyPrintTimeFromDateTime(new Date(t))} - ${this.prettyPrintTimeFromDateTime(new Date(n))}`}loadEvents(){return N(this,void 0,void 0,function*(){const e=_.a.globalProvider;let t=[];if((null==e?void 0:e.state)===i.c.SignedIn){const n=e.graph.forComponent(this);if(this.eventQuery)try{const e=this.eventQuery.split("|");let a,i;e.length>1?(i=e[0].trim(),a=e[1].trim()):i=this.eventQuery;const r=yield y(n,i,a);if(null==r?void 0:r.value)for(t=r.value;r.hasNext;)yield r.next(),t=r.value}catch(e){}else{const e=this.date?new Date(this.date):new Date,a=new Date(e.getTime());a.setDate(e.getDate()+this.days);try{const i=yield((e,t,n,a)=>v(void 0,void 0,void 0,function*(){const i=`startdatetime=${t.toISOString()}`,r=`enddatetime=${n.toISOString()}`;let o;return o=a?`groups/${a}/calendar`:"me",o+=`/calendarview?${i}&${r}`,y(e,o)}))(n,e,a,this.groupId);if(null==i?void 0:i.value)for(t=i.value;i.hasNext;)yield i.next(),t=i.value}catch(e){}}}return t})}prettyPrintTimeFromDateTime(e){return e.toLocaleTimeString(navigator.language,{timeStyle:"short",timeZone:this.preferredTimezone})}getDateHeaderFromDateTimeString(e){return new Date(e).toLocaleDateString(navigator.language,{dateStyle:"full",timeZone:this.preferredTimezone})}}H([Object(f.b)({attribute:"date",type:String}),R("design:type",String),R("design:paramtypes",[Object])],j.prototype,"date",null),H([Object(f.b)({attribute:"group-id",type:String}),R("design:type",String),R("design:paramtypes",[Object])],j.prototype,"groupId",null),H([Object(f.b)({attribute:"days",type:Number}),R("design:type",Number),R("design:paramtypes",[Object])],j.prototype,"days",null),H([Object(f.b)({attribute:"event-query",type:String}),R("design:type",String),R("design:paramtypes",[Object])],j.prototype,"eventQuery",null),H([Object(f.b)({attribute:"events",type:Array}),R("design:type",Array)],j.prototype,"events",void 0),H([Object(f.b)({attribute:"show-max",type:Number}),R("design:type",Number)],j.prototype,"showMax",void 0),H([Object(f.b)({attribute:"group-by-day",type:Boolean}),R("design:type",Boolean)],j.prototype,"groupByDay",void 0),H([Object(f.b)({attribute:"preferred-timezone",type:String}),R("design:type",String),R("design:paramtypes",[Object])],j.prototype,"preferredTimezone",null),H([Object(f.b)({attribute:!1}),R("design:type",Boolean)],j.prototype,"_isNarrow",void 0);var V=n("DNu6"),z=n("glp4"),G=n("0Uyf"),K=n("imsm"),W=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},q=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};const Q=e=>Array.isArray(null==e?void 0:e.value);var Y;!function(e){e.json="json",e.image="image"}(Y||(Y={}));const J=()=>V.a.config.response.isEnabled&&V.a.config.isEnabled,X=()=>Object(A.b)("get",Z);class Z extends p.a{constructor(){super(...arguments),this.scopes=[],this.version="v1.0",this.type=Y.json,this.maxPages=3,this.pollingRate=0,this.cacheEnabled=!1,this.cacheInvalidationPeriod=0,this.isPolling=!1,this.isRefreshing=!1}attributeChangedCallback(e,t,n){super.attributeChangedCallback(e,t,n),this.requestStateUpdate()}refresh(e=!1){this.isRefreshing=!0,e&&this.clearState(),this.requestStateUpdate(e)}clearState(){this.response=null}render(){if(this.isLoadingState&&!this.response)return this.renderTemplate("loading",null);if(this.error)return this.renderTemplate("error",this.error);if(this.hasTemplate("value")&&Q(this.response)){let e;if(Q(this.response)){let t=null;this.isLoadingState&&!this.isPolling&&(t=this.renderTemplate("loading",null)),e=u.c`
          ${this.response.value.map(e=>this.renderTemplate("value",e,e.id))} ${t}
        `}else e=this.renderTemplate("value",this.response);if(this.hasTemplate("default")){const t=this.renderTemplate("default",this.response);return this.templates.value.templateOrder>this.templates.default.templateOrder?u.c`
            ${t}${e}
          `:u.c`
            ${e}${t}
          `}return e}return this.response?this.renderTemplate("default",this.response)||u.c``:this.hasTemplate("no-data")?this.renderTemplate("no-data",null):u.c``}loadState(){var e,t,n,a,r,o;return a=this,void 0,o=function*(){const a=_.a.globalProvider;if(this.error=null,a&&a.state===i.c.SignedIn){if(this.resource){try{let i;const r=`${this.version}${this.resource}`;let o=null;if(this.shouldRetrieveCache()){i=V.a.getCache(K.a.get,K.a.get.stores.responses);const e=J()?yield i.getValue(r):null;e&&(this.cacheInvalidationPeriod||V.a.config.response.invalidationPeriod||V.a.config.defaultInvalidationPeriod)>Date.now()-e.timeCached&&(o=JSON.parse(e.response))}if(!o){let s=this.resource,c=!1;(null===(e=this.response)||void 0===e?void 0:e["@odata.deltaLink"])?(s=this.response["@odata.deltaLink"],c=!0):c=new URL(s,"https://graph.microsoft.com").pathname.endsWith("delta");const d=a.graph.forComponent(this);let l=d.api(s).version(this.version);if((null===(t=this.scopes)||void 0===t?void 0:t.length)&&(l=l.middlewareOptions(Object(b.a)(...this.scopes))),this.type===Y.json){if(o=yield l.get(),c&&Q(this.response)&&Q(o)){const e=o.value;o.value=this.response.value.concat(e)}if(this.isPolling||Object(O.b)(this.response,o)||(this.response=o),Q(o)&&o["@odata.nextLink"]){let e=1,t=o;for(;(e<this.maxPages||this.maxPages<=0||c&&this.pollingRate)&&(null==t?void 0:t["@odata.nextLink"]);){e++;const a=t["@odata.nextLink"].split(this.version)[1];t=yield d.client.api(a).version(this.version).get(),(null===(n=null==t?void 0:t.value)||void 0===n?void 0:n.length)&&(t.value=o.value.concat(t.value),o=t,this.isPolling||(this.response=o))}}}else{if(-1===this.resource.indexOf("/photo/$value")&&-1===this.resource.indexOf("/thumbnails/"))throw new Error("Only /photo/$value and /thumbnails/ endpoints support the image type");let e;if(this.resource.indexOf("/photo/$value")>-1){const t=this.resource.replace("/photo/$value",""),n=yield Object(z.e)(d,t,this.scopes);n&&(e=n.photo)}else if(this.resource.indexOf("/thumbnails/")>-1){const t=yield Object(G.d)(d,this.resource,this.scopes);t&&(e=t.thumbnail)}e&&(o={image:e})}this.shouldUpdateCache()&&o&&(i=V.a.getCache(K.a.get,K.a.get.stores.responses),yield i.putValue(r,{response:JSON.stringify(o)}))}Object(O.b)(this.response,o)||(this.response=o)}catch(e){this.error=e}this.response&&(this.error=null,this.pollingRate&&setTimeout(()=>{this.isPolling=!0,this.loadState().finally(()=>{this.isPolling=!1})},this.pollingRate))}else this.response=null;this.isRefreshing=!1,this.fireCustomEvent("dataChange",{response:this.response,error:this.error})}},new((r=void 0)||(r=Promise))(function(e,t){function n(e){try{s(o.next(e))}catch(e){t(e)}}function i(e){try{s(o.throw(e))}catch(e){t(e)}}function s(t){var a;t.done?e(t.value):(a=t.value,a instanceof r?a:new r(function(e){e(a)})).then(n,i)}s((o=o.apply(a,[])).next())})}shouldRetrieveCache(){return J()&&this.cacheEnabled&&!(this.isRefreshing||this.isPolling)}shouldUpdateCache(){return J()&&this.cacheEnabled}}W([Object(f.b)({attribute:"resource",reflect:!0,type:String}),q("design:type",String)],Z.prototype,"resource",void 0),W([Object(f.b)({attribute:"scopes",converter:(e,t)=>e?e.toLowerCase().split(","):null,reflect:!0}),q("design:type",Array)],Z.prototype,"scopes",void 0),W([Object(f.b)({attribute:"version",reflect:!0,type:String}),q("design:type",Object)],Z.prototype,"version",void 0),W([Object(f.b)({attribute:"type",reflect:!0,type:Y}),q("design:type",String)],Z.prototype,"type",void 0),W([Object(f.b)({attribute:"max-pages",reflect:!0,type:Number}),q("design:type",Object)],Z.prototype,"maxPages",void 0),W([Object(f.b)({attribute:"polling-rate",reflect:!0,type:Number}),q("design:type",Object)],Z.prototype,"pollingRate",void 0),W([Object(f.b)({attribute:"cache-enabled",reflect:!0,type:Boolean}),q("design:type",Object)],Z.prototype,"cacheEnabled",void 0),W([Object(f.b)({attribute:"cache-invalidation-period",type:Number}),q("design:type",Object)],Z.prototype,"cacheInvalidationPeriod",void 0),W([Object(f.b)({attribute:!1}),q("design:type",Object)],Z.prototype,"response",void 0),W([Object(f.b)({attribute:!1}),q("design:type",Object)],Z.prototype,"error",void 0);var $=n("ox5k"),ee=n("ZzBS"),te=n("c1DA"),ne=n("hgjj"),ae=n("dP6N");const ie=[u.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host .signed-in-person{--person-background-color:$signed-in-background-color;padding:var(--login-button-padding,10px 16px)}:host .account{padding:0;margin:7px 0}:host fluent-button.signed-in{height:auto;min-width:auto}:host fluent-button.signed-in::part(control){width:100%;height:100%;padding:var(--login-button-padding,0);background:var(--login-signed-in-background,transparent);word-spacing:inherit;text-indent:inherit;letter-spacing:inherit}:host fluent-button.signed-in::part(control):focus-within,:host fluent-button.signed-in::part(control):hover{background:var(--login-signed-in-hover-background,var(--neutral-fill-stealth-hover));--secondary-text-color:var(--secondary-text-hover-color)}:host fluent-button.signed-out::part(control){color:var(--login-signed-out-button-text-color,var(--neutral-fill-foreground-rest));background:var(--login-signed-out-button-background,padding-box linear-gradient(var(--neutral-fill-rest),var(--neutral-fill-rest)),border-box var(--neutral-stroke-control-rest))}:host fluent-button.signed-out::part(control):focus-within,:host fluent-button.signed-out::part(control):hover{background:var(--login-signed-out-button-hover-background,var(--neutral-fill-stealth-hover))}:host fluent-button.small::part(control):hover{background:0 0}:host fluent-button:focus-visible{outline-style:var(--focus-ring-style,auto)}:host fluent-card{--fill-color:var(--login-popup-background-color, var(--neutral-layer-1));padding:var(--login-popup-padding,16px)}:host .login-root .small .signed-in-person{padding:0;background:0 0}:host .login-root .small .signed-in-person:focus-within,:host .login-root .small .signed-in-person:hover{background:0 0}:host .login-root .account-list{padding:calc(var(--design-unit) * 1px) 0;margin:0}:host .login-root .account-item{height:auto;min-width:auto;margin-top:4px;background:var(--login-popup-background-color,var(--neutral-layer-1));list-style-type:none;cursor:pointer}:host .login-root .account-item:hover{background:var(--login-account-item-hover-bg-color,var(--neutral-fill-stealth-hover));--person-background-color:$login-account-item-hover-bg-color}:host .login-root .flyout .flyout-command{color:var(--login-flyout-command-text-color,var(--accent-foreground-rest))}:host .login-root .flyout .popup-content .commands{display:flex;align-items:flex-end;justify-content:flex-end}:host .login-root .flyout .popup-content .commands fluent-button::part(control){color:var(--login-command-button-text-color,var(--neutral-fill-foreground-rest));background:var(--login-command-button-background-color,var(--neutral-fill-stealth-rest));word-spacing:inherit;text-indent:inherit;letter-spacing:inherit}:host .login-root .flyout .popup-content .commands fluent-button::part(control):hover{background:var(--login-command-button-hover-background-color,var(--neutral-fill-stealth-hover))}:host .login-root .flyout .popup-content .content .main-profile{margin-bottom:56px;margin-top:27px}:host .login-root .flyout .popup-content .add-account{padding-top:16px}:host .login-root .flyout .popup-content .add-account fluent-button::part(control){color:var(--login-add-account-button-text-color,var(--neutral-fill-foreground-rest));background:var(--login-add-account-button-background-color,var(--neutral-fill-stealth-rest));word-spacing:inherit;text-indent:inherit;letter-spacing:inherit}:host .login-root .flyout .popup-content .add-account fluent-button::part(control):hover{background:var(--login-add-account-button-hover-background-color,var(--neutral-fill-stealth-hover))}
`],re={signInLinkSubtitle:"Sign In",signOutLinkSubtitle:"Sign Out",signInWithADifferentAccount:"Sign in with a different account"};var oe=n("eftJ"),se=n("oePG"),ce=n("QBS5"),de=n("Gy7L"),le=n("kpPY"),ue=n("TDEi"),fe=n("FGLN"),pe=n("uXNP"),me=n("C5kU"),_e=n("6fxl");function he(e){return Object(fe.c)(e)&&("option"===e.getAttribute("role")||e instanceof HTMLOptionElement)}class be extends ue.a{constructor(e,t,n,a){super(),this.defaultSelected=!1,this.dirtySelected=!1,this.selected=this.defaultSelected,this.dirtyValue=!1,e&&(this.textContent=e),t&&(this.initialValue=t),n&&(this.defaultSelected=n),a&&(this.selected=a),this.proxy=new Option(`${this.textContent}`,this.initialValue,this.defaultSelected,this.selected),this.proxy.disabled=this.disabled}checkedChanged(e,t){this.ariaChecked="boolean"!=typeof t?null:t?"true":"false"}contentChanged(e,t){this.proxy instanceof HTMLOptionElement&&(this.proxy.textContent=this.textContent),this.$emit("contentchange",null,{bubbles:!0})}defaultSelectedChanged(){this.dirtySelected||(this.selected=this.defaultSelected,this.proxy instanceof HTMLOptionElement&&(this.proxy.selected=this.defaultSelected))}disabledChanged(e,t){this.ariaDisabled=this.disabled?"true":"false",this.proxy instanceof HTMLOptionElement&&(this.proxy.disabled=this.disabled)}selectedAttributeChanged(){this.defaultSelected=this.selectedAttribute,this.proxy instanceof HTMLOptionElement&&(this.proxy.defaultSelected=this.defaultSelected)}selectedChanged(){this.ariaSelected=this.selected?"true":"false",this.dirtySelected||(this.dirtySelected=!0),this.proxy instanceof HTMLOptionElement&&(this.proxy.selected=this.selected)}initialValueChanged(e,t){this.dirtyValue||(this.value=this.initialValue,this.dirtyValue=!1)}get label(){var e;return null!==(e=this.value)&&void 0!==e?e:this.text}get text(){var e,t;return null!==(t=null===(e=this.textContent)||void 0===e?void 0:e.replace(/\s+/g," ").trim())&&void 0!==t?t:""}set value(e){const t=`${null!=e?e:""}`;this._value=t,this.dirtyValue=!0,this.proxy instanceof HTMLOptionElement&&(this.proxy.value=t),se.b.notify(this,"value")}get value(){var e;return se.b.track(this,"value"),null!==(e=this._value)&&void 0!==e?e:this.text}get form(){return this.proxy?this.proxy.form:null}}Object(oe.a)([se.d],be.prototype,"checked",void 0),Object(oe.a)([se.d],be.prototype,"content",void 0),Object(oe.a)([se.d],be.prototype,"defaultSelected",void 0),Object(oe.a)([Object(ce.c)({mode:"boolean"})],be.prototype,"disabled",void 0),Object(oe.a)([Object(ce.c)({attribute:"selected",mode:"boolean"})],be.prototype,"selectedAttribute",void 0),Object(oe.a)([se.d],be.prototype,"selected",void 0),Object(oe.a)([Object(ce.c)({attribute:"value",mode:"fromView"})],be.prototype,"initialValue",void 0);class ge{}Object(oe.a)([se.d],ge.prototype,"ariaChecked",void 0),Object(oe.a)([se.d],ge.prototype,"ariaPosInSet",void 0),Object(oe.a)([se.d],ge.prototype,"ariaSelected",void 0),Object(oe.a)([se.d],ge.prototype,"ariaSetSize",void 0),Object(_e.a)(ge,pe.a),Object(_e.a)(be,me.a,ge);class ve extends ue.a{constructor(){super(...arguments),this._options=[],this.selectedIndex=-1,this.selectedOptions=[],this.shouldSkipFocus=!1,this.typeaheadBuffer="",this.typeaheadExpired=!0,this.typeaheadTimeout=-1}get firstSelectedOption(){var e;return null!==(e=this.selectedOptions[0])&&void 0!==e?e:null}get hasSelectableOptions(){return this.options.length>0&&!this.options.every(e=>e.disabled)}get length(){var e,t;return null!==(t=null===(e=this.options)||void 0===e?void 0:e.length)&&void 0!==t?t:0}get options(){return se.b.track(this,"options"),this._options}set options(e){this._options=e,se.b.notify(this,"options")}get typeAheadExpired(){return this.typeaheadExpired}set typeAheadExpired(e){this.typeaheadExpired=e}clickHandler(e){const t=e.target.closest("option,[role=option]");if(t&&!t.disabled)return this.selectedIndex=this.options.indexOf(t),!0}focusAndScrollOptionIntoView(e=this.firstSelectedOption){this.contains(document.activeElement)&&null!==e&&(e.focus(),requestAnimationFrame(()=>{e.scrollIntoView({block:"nearest"})}))}focusinHandler(e){this.shouldSkipFocus||e.target!==e.currentTarget||(this.setSelectedOptions(),this.focusAndScrollOptionIntoView()),this.shouldSkipFocus=!1}getTypeaheadMatches(){const e=this.typeaheadBuffer.replace(/[.*+\-?^${}()|[\]\\]/g,"\\$&"),t=new RegExp(`^${e}`,"gi");return this.options.filter(e=>e.text.trim().match(t))}getSelectableIndex(e=this.selectedIndex,t){const n=e>t?-1:e<t?1:0,a=e+n;let i=null;switch(n){case-1:i=this.options.reduceRight((e,t,n)=>!e&&!t.disabled&&n<a?t:e,i);break;case 1:i=this.options.reduce((e,t,n)=>!e&&!t.disabled&&n>a?t:e,i)}return this.options.indexOf(i)}handleChange(e,t){"selected"===t&&(ve.slottedOptionFilter(e)&&(this.selectedIndex=this.options.indexOf(e)),this.setSelectedOptions())}handleTypeAhead(e){this.typeaheadTimeout&&window.clearTimeout(this.typeaheadTimeout),this.typeaheadTimeout=window.setTimeout(()=>this.typeaheadExpired=!0,ve.TYPE_AHEAD_TIMEOUT_MS),e.length>1||(this.typeaheadBuffer=`${this.typeaheadExpired?"":this.typeaheadBuffer}${e}`)}keydownHandler(e){if(this.disabled)return!0;this.shouldSkipFocus=!1;const t=e.key;switch(t){case de.j:e.shiftKey||(e.preventDefault(),this.selectFirstOption());break;case de.b:e.shiftKey||(e.preventDefault(),this.selectNextOption());break;case de.e:e.shiftKey||(e.preventDefault(),this.selectPreviousOption());break;case de.f:e.preventDefault(),this.selectLastOption();break;case de.n:return this.focusAndScrollOptionIntoView(),!0;case de.g:case de.h:return!0;case de.m:if(this.typeaheadExpired)return!0;default:return 1===t.length&&this.handleTypeAhead(`${t}`),!0}}mousedownHandler(e){return this.shouldSkipFocus=!this.contains(document.activeElement),!0}multipleChanged(e,t){this.ariaMultiSelectable=t?"true":null}selectedIndexChanged(e,t){var n;if(this.hasSelectableOptions){if((null===(n=this.options[this.selectedIndex])||void 0===n?void 0:n.disabled)&&"number"==typeof e){const n=this.getSelectableIndex(e,t),a=n>-1?n:e;return this.selectedIndex=a,void(t===a&&this.selectedIndexChanged(t,a))}this.setSelectedOptions()}else this.selectedIndex=-1}selectedOptionsChanged(e,t){var n;const a=t.filter(ve.slottedOptionFilter);null===(n=this.options)||void 0===n||n.forEach(e=>{const t=se.b.getNotifier(e);t.unsubscribe(this,"selected"),e.selected=a.includes(e),t.subscribe(this,"selected")})}selectFirstOption(){var e,t;this.disabled||(this.selectedIndex=null!==(t=null===(e=this.options)||void 0===e?void 0:e.findIndex(e=>!e.disabled))&&void 0!==t?t:-1)}selectLastOption(){this.disabled||(this.selectedIndex=function(e,t){let n=e.length;for(;n--;)if(!e[n].disabled)return n;return-1}(this.options))}selectNextOption(){!this.disabled&&this.selectedIndex<this.options.length-1&&(this.selectedIndex+=1)}selectPreviousOption(){!this.disabled&&this.selectedIndex>0&&(this.selectedIndex=this.selectedIndex-1)}setDefaultSelectedOption(){var e,t;this.selectedIndex=null!==(t=null===(e=this.options)||void 0===e?void 0:e.findIndex(e=>e.defaultSelected))&&void 0!==t?t:-1}setSelectedOptions(){var e,t,n;(null===(e=this.options)||void 0===e?void 0:e.length)&&(this.selectedOptions=[this.options[this.selectedIndex]],this.ariaActiveDescendant=null!==(n=null===(t=this.firstSelectedOption)||void 0===t?void 0:t.id)&&void 0!==n?n:"",this.focusAndScrollOptionIntoView())}slottedOptionsChanged(e,t){this.options=t.reduce((e,t)=>(he(t)&&e.push(t),e),[]);const n=`${this.options.length}`;this.options.forEach((e,t)=>{e.id||(e.id=Object(le.a)("option-")),e.ariaPosInSet=`${t+1}`,e.ariaSetSize=n}),this.$fastController.isConnected&&(this.setSelectedOptions(),this.setDefaultSelectedOption())}typeaheadBufferChanged(e,t){if(this.$fastController.isConnected){const e=this.getTypeaheadMatches();if(e.length){const t=this.options.indexOf(e[0]);t>-1&&(this.selectedIndex=t)}this.typeaheadExpired=!1}}}ve.slottedOptionFilter=e=>he(e)&&!e.hidden,ve.TYPE_AHEAD_TIMEOUT_MS=1e3,Object(oe.a)([Object(ce.c)({mode:"boolean"})],ve.prototype,"disabled",void 0),Object(oe.a)([se.d],ve.prototype,"selectedIndex",void 0),Object(oe.a)([se.d],ve.prototype,"selectedOptions",void 0),Object(oe.a)([se.d],ve.prototype,"slottedOptions",void 0),Object(oe.a)([se.d],ve.prototype,"typeaheadBuffer",void 0);class ye{}Object(oe.a)([se.d],ye.prototype,"ariaActiveDescendant",void 0),Object(oe.a)([se.d],ye.prototype,"ariaDisabled",void 0),Object(oe.a)([se.d],ye.prototype,"ariaExpanded",void 0),Object(oe.a)([se.d],ye.prototype,"ariaMultiSelectable",void 0),Object(_e.a)(ye,pe.a),Object(_e.a)(ve,ye);var Se=n("6BDD"),De=n("UauI"),Ie=n("5ZAu"),xe=n("VRJB");class Ce extends ve{constructor(){super(...arguments),this.activeIndex=-1,this.rangeStartIndex=-1}get activeOption(){return this.options[this.activeIndex]}get checkedOptions(){var e;return null===(e=this.options)||void 0===e?void 0:e.filter(e=>e.checked)}get firstSelectedOptionIndex(){return this.options.indexOf(this.firstSelectedOption)}activeIndexChanged(e,t){var n,a;this.ariaActiveDescendant=null!==(a=null===(n=this.options[t])||void 0===n?void 0:n.id)&&void 0!==a?a:"",this.focusAndScrollOptionIntoView()}checkActiveIndex(){if(!this.multiple)return;const e=this.activeOption;e&&(e.checked=!0)}checkFirstOption(e=!1){e?(-1===this.rangeStartIndex&&(this.rangeStartIndex=this.activeIndex+1),this.options.forEach((e,t)=>{e.checked=Object(xe.a)(t,this.rangeStartIndex)})):this.uncheckAllOptions(),this.activeIndex=0,this.checkActiveIndex()}checkLastOption(e=!1){e?(-1===this.rangeStartIndex&&(this.rangeStartIndex=this.activeIndex),this.options.forEach((e,t)=>{e.checked=Object(xe.a)(t,this.rangeStartIndex,this.options.length)})):this.uncheckAllOptions(),this.activeIndex=this.options.length-1,this.checkActiveIndex()}connectedCallback(){super.connectedCallback(),this.addEventListener("focusout",this.focusoutHandler)}disconnectedCallback(){this.removeEventListener("focusout",this.focusoutHandler),super.disconnectedCallback()}checkNextOption(e=!1){e?(-1===this.rangeStartIndex&&(this.rangeStartIndex=this.activeIndex),this.options.forEach((e,t)=>{e.checked=Object(xe.a)(t,this.rangeStartIndex,this.activeIndex+1)})):this.uncheckAllOptions(),this.activeIndex+=this.activeIndex<this.options.length-1?1:0,this.checkActiveIndex()}checkPreviousOption(e=!1){e?(-1===this.rangeStartIndex&&(this.rangeStartIndex=this.activeIndex),1===this.checkedOptions.length&&(this.rangeStartIndex+=1),this.options.forEach((e,t)=>{e.checked=Object(xe.a)(t,this.activeIndex,this.rangeStartIndex)})):this.uncheckAllOptions(),this.activeIndex-=this.activeIndex>0?1:0,this.checkActiveIndex()}clickHandler(e){var t;if(!this.multiple)return super.clickHandler(e);const n=null===(t=e.target)||void 0===t?void 0:t.closest("[role=option]");return n&&!n.disabled?(this.uncheckAllOptions(),this.activeIndex=this.options.indexOf(n),this.checkActiveIndex(),this.toggleSelectedForAllCheckedOptions(),!0):void 0}focusAndScrollOptionIntoView(){super.focusAndScrollOptionIntoView(this.activeOption)}focusinHandler(e){if(!this.multiple)return super.focusinHandler(e);this.shouldSkipFocus||e.target!==e.currentTarget||(this.uncheckAllOptions(),-1===this.activeIndex&&(this.activeIndex=-1!==this.firstSelectedOptionIndex?this.firstSelectedOptionIndex:0),this.checkActiveIndex(),this.setSelectedOptions(),this.focusAndScrollOptionIntoView()),this.shouldSkipFocus=!1}focusoutHandler(e){this.multiple&&this.uncheckAllOptions()}keydownHandler(e){if(!this.multiple)return super.keydownHandler(e);if(this.disabled)return!0;const{key:t,shiftKey:n}=e;switch(this.shouldSkipFocus=!1,t){case de.j:return void this.checkFirstOption(n);case de.b:return void this.checkNextOption(n);case de.e:return void this.checkPreviousOption(n);case de.f:return void this.checkLastOption(n);case de.n:return this.focusAndScrollOptionIntoView(),!0;case de.h:return this.uncheckAllOptions(),this.checkActiveIndex(),!0;case de.m:if(e.preventDefault(),this.typeAheadExpired)return void this.toggleSelectedForAllCheckedOptions();default:return 1===t.length&&this.handleTypeAhead(`${t}`),!0}}mousedownHandler(e){if(e.offsetX>=0&&e.offsetX<=this.scrollWidth)return super.mousedownHandler(e)}multipleChanged(e,t){var n;this.ariaMultiSelectable=t?"true":null,null===(n=this.options)||void 0===n||n.forEach(e=>{e.checked=!t&&void 0}),this.setSelectedOptions()}setSelectedOptions(){this.multiple?this.$fastController.isConnected&&this.options&&(this.selectedOptions=this.options.filter(e=>e.selected),this.focusAndScrollOptionIntoView()):super.setSelectedOptions()}sizeChanged(e,t){var n;const a=Math.max(0,parseInt(null!==(n=null==t?void 0:t.toFixed())&&void 0!==n?n:"",10));a!==t&&Ie.a.queueUpdate(()=>{this.size=a})}toggleSelectedForAllCheckedOptions(){const e=this.checkedOptions.filter(e=>!e.disabled),t=!e.every(e=>e.selected);e.forEach(e=>e.selected=t),this.selectedIndex=this.options.indexOf(e[e.length-1]),this.setSelectedOptions()}typeaheadBufferChanged(e,t){if(this.multiple){if(this.$fastController.isConnected){const e=this.getTypeaheadMatches(),t=this.options.indexOf(e[0]);t>-1&&(this.activeIndex=t,this.uncheckAllOptions(),this.checkActiveIndex()),this.typeAheadExpired=!1}}else super.typeaheadBufferChanged(e,t)}uncheckAllOptions(e=!1){this.options.forEach(e=>e.checked=!this.multiple&&void 0),e||(this.rangeStartIndex=-1)}}Object(oe.a)([se.d],Ce.prototype,"activeIndex",void 0),Object(oe.a)([Object(ce.c)({mode:"boolean"})],Ce.prototype,"multiple",void 0),Object(oe.a)([Object(ce.c)({converter:ce.e})],Ce.prototype,"size",void 0);var Oe=n("4X57"),we=n("xY0q"),Ee=n("8hiW"),Ae=n("FVZ7");const Le=class extends ve{}.compose({baseName:"listbox",template:(e,t)=>Se.a`
    <template
        aria-activedescendant="${e=>e.ariaActiveDescendant}"
        aria-multiselectable="${e=>e.ariaMultiSelectable}"
        class="listbox"
        role="listbox"
        tabindex="${e=>e.disabled?null:"0"}"
        @click="${(e,t)=>e.clickHandler(t.event)}"
        @focusin="${(e,t)=>e.focusinHandler(t.event)}"
        @keydown="${(e,t)=>e.keydownHandler(t.event)}"
        @mousedown="${(e,t)=>e.mousedownHandler(t.event)}"
    >
        <slot
            ${Object(De.a)({filter:Ce.slottedOptionFilter,flatten:!0,property:"slottedOptions"})}
        ></slot>
    </template>
`,styles:(e,t)=>Oe.a`
    ${Object(we.a)("inline-flex")} :host {
      border: calc(${Ee.vb} * 1px) solid ${Ee.rb};
      border-radius: calc(${Ee.q} * 1px);
      box-sizing: border-box;
      flex-direction: column;
      padding: calc(${Ee.s} * 1px) 0;
    }

    ::slotted(${e.tagFor(be)}) {
      margin: 0 calc(${Ee.s} * 1px);
    }

    :host(:focus-within:not([disabled])) {
      ${Ae.a}
    }
  `});var ke=n("VcPv"),Me=n("m1Vi"),Pe=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},Te=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)},Ue=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const Fe=()=>{Object(T.a)(Le,ke.a,Me.a,U.a),Object(te.a)(),Object(d.c)(),Object(A.b)("login",He)};class He extends p.a{static get styles(){return ie}get strings(){return re}get flyout(){return this.renderRoot.querySelector(".flyout")}static get requiredScopes(){return[...new Set(["user.read",...d.a.requiredScopes])]}get _userDetailsKey(){return"-userDetails"}constructor(){super(),this.showPresence=!1,this.loginView="full",this._arrowKeyLocation=-1,this.logout=()=>Ue(this,void 0,void 0,function*(){if(!this.fireCustomEvent("logoutInitiated"))return;const e=_.a.globalProvider;(null==e?void 0:e.isMultiAccountSupportedAndEnabled)&&localStorage.removeItem(e.getActiveAccount().id+this._userDetailsKey),(null==e?void 0:e.logout)&&(yield e.logout(),this.userDetails=null,e.isMultiAccountSupportedAndEnabled&&localStorage.removeItem(e.getActiveAccount().id+this._userDetailsKey),this.hideFlyout(),this.fireCustomEvent("logoutCompleted"))}),this.flyoutOpened=()=>{this._isFlyoutOpen=!0},this.flyoutClosed=()=>{this._isFlyoutOpen=!1},this.onUserKeyDown=e=>{if(!this.flyout.isOpen)return;const t=this.renderRoot.querySelector(".popup-content"),n=t.querySelectorAll("ul, fluent-button"),a=t.querySelector("#signout-button")||n[0],i=t.querySelector("#signin-different-account-button")||n[n.length-1];if("Tab"===e.key&&e.shiftKey&&a===e.target&&(e.preventDefault(),null==i||i.focus()),"Tab"!==e.key||e.shiftKey||i!==e.target||(e.preventDefault(),null==a||a.focus()),"Escape"===e.key){const e=this.renderRoot.querySelector("#login-button");null==e||e.focus()}const r=this.renderRoot.querySelector("fluent-card");e.shiftKey&&"Tab"===e.key&&e.target===r&&this.hideFlyout()},this.handleAccountListKeyDown=e=>{const t=this.renderRoot.querySelector("ul.account-list");let n;const a=null==t?void 0:t.children;for(const e of a){const t=e;t.setAttribute("tabindex","-1"),t.blur()}const i=t.childElementCount,r=e.key;if("ArrowDown"===r)this._arrowKeyLocation=(this._arrowKeyLocation+1+i)%i;else if("ArrowUp"===r)this._arrowKeyLocation=(this._arrowKeyLocation-1+i)%i;else if("Tab"===r||"Escape"===r)return this._arrowKeyLocation=-1,t.blur(),void("Escape"===r&&(e.preventDefault(),e.stopPropagation()));this._arrowKeyLocation>-1&&(n=a[this._arrowKeyLocation],n.setAttribute("tabindex","1"),n.focus())},this.onClick=()=>{this.userDetails&&this._isFlyoutOpen?this.hideFlyout():this.userDetails?this.showFlyout():this.login()},this._isFlyoutOpen=!1}connectedCallback(){super.connectedCallback(),this.addEventListener("click",e=>e.stopPropagation())}login(){return Ue(this,void 0,void 0,function*(){const e=_.a.globalProvider;(e.isMultiAccountSupportedAndEnabled||!this.userDetails&&this.fireCustomEvent("loginInitiated"))&&(null==e?void 0:e.login)&&(yield e.login(),e.state===i.c.SignedIn?this.fireCustomEvent("loginCompleted"):this.fireCustomEvent("loginFailed"))})}render(){return u.c`
      <div class="login-root">
        ${this.renderButton()}
        ${this.renderFlyout()}
      </div>
    `}loadState(){return Ue(this,void 0,void 0,function*(){const e=_.a.globalProvider;e&&!this.userDetails&&(e.state===i.c.SignedIn?(this.userDetails=yield Object(ne.a)(e.graph.forComponent(this)),this.userDetails.personImage&&(this._image=this.userDetails.personImage),e.isMultiAccountSupportedAndEnabled&&localStorage.setItem(_.a.globalProvider.getActiveAccount().id+this._userDetailsKey,JSON.stringify(this.userDetails)),this.fireCustomEvent("loginCompleted")):this.userDetails=null)})}renderButton(){var e;const t=(null===(e=_.a.globalProvider)||void 0===e?void 0:e.state)===i.c.SignedIn,n=Object(F.a)({"signed-in":t&&Boolean(this.userDetails),"signed-out":!t,small:"avatar"===this.loginView}),a=t?"stealth":"neutral",r=t&&this.userDetails,o=r?this.renderSignedInButtonContent(this.userDetails,this._image):this.renderSignedOutButtonContent(),s=r?this._isFlyoutOpen:void 0;return u.c`
      <fluent-button
        id="login-button"
        aria-expanded="${Object($.a)(s)}"
        appearance=${a}
        aria-label="${Object($.a)(t?void 0:this.strings.signInLinkSubtitle)}"
        ?disabled=${this.isLoadingState}
        @click=${this.onClick}
        class=${n}>
          ${o}
      </fluent-button>`}renderFlyout(){return m.a`
      <mgt-flyout
        class="flyout"
        light-dismiss
        @opened=${this.flyoutOpened}
        @closed=${this.flyoutClosed}>
        <fluent-card 
          slot="flyout" 
          tabindex="0" 
          class="flyout-card"
          @keydown=${this.onUserKeyDown}
          >
          ${this.renderFlyoutContent()}
        </fluent-card>
      </mgt-flyout>`}renderFlyoutContent(){if(this.userDetails)return u.c`
       <div class="popup">
         <div class="popup-content">
           <div class="commands">
             ${this.renderFlyoutCommands()}
           </div>
           <div class="content">
             <div class="main-profile">
               ${this.renderFlyoutPersonDetails(this.userDetails,this._image)}
             </div>
             ${this.renderAccounts()}
           </div>
           ${this.renderAddAccountContent()}
         </div>
       </div>
     `}get hasMultipleAccounts(){var e,t,n,a;return(null===(e=_.a.globalProvider)||void 0===e?void 0:e.isMultiAccountSupportedAndEnabled)&&(null===(a=null===(n=null===(t=_.a.globalProvider)||void 0===t?void 0:t.getAllAccounts)||void 0===n?void 0:n.call(t))||void 0===a?void 0:a.length)>1}get usesVerticalPersonCard(){return"full"===this.loginView||this.hasMultipleAccounts}renderFlyoutPersonDetails(e,t){return this.renderTemplate("flyout-person-details",{personDetails:e,personImage:t})||m.a`
        <mgt-person
          .personDetails=${e}
          .personImage=${t}
          .view=${ee.a.twolines}
          .line2Property=${"email"}
          ?vertical-layout=${this.usesVerticalPersonCard}
          class="person">
        </mgt-person>`}renderFlyoutCommands(){return this.renderTemplate("flyout-commands",{handleSignOut:()=>this.logout()})||u.c`
        <fluent-button
          id="signout-button"
          appearance="stealth"
          size="medium"
          class="flyout-command"
          @click=${this.logout}
          aria-label=${this.strings.signOutLinkSubtitle}>
            ${this.strings.signOutLinkSubtitle}
        </fluent-button>`}renderButtonContent(){return this.userDetails?this.renderSignedInButtonContent(this.userDetails,this._image):this.renderSignedOutButtonContent()}renderAddAccountContent(){if(_.a.globalProvider.isMultiAccountSupportedAndEnabled)return u.c`
        <div class="add-account">
          <fluent-button
            id="signin-different-account-button"
            appearance="stealth"
            aria-label="${this.strings.signInWithADifferentAccount}"
            @click=${()=>{this.login()}}>
            <span slot="start"><i>${Object(S.b)(S.a.SelectAccount,"currentColor")}</i></span>
            ${this.strings.signInWithADifferentAccount}
          </fluent-button>
        </div>`}parsePersonDisplayConfiguration(){const e={view:ee.a.twolines,avatarSize:"small"};switch(this.loginView){case"avatar":e.view=ee.a.image,e.avatarSize="small";break;case"compact":e.view=ee.a.oneline,e.avatarSize="small";break;default:e.view=ee.a.twolines,e.avatarSize="auto"}return e}renderSignedInButtonContent(e,t){const n=this.renderTemplate("signed-in-button-content",{personDetails:e,personImage:t}),a=this.parsePersonDisplayConfiguration();return n||m.a`
        <mgt-person
          .personDetails=${this.userDetails}
          .personImage=${this._image}
          .view=${a.view}
          .showPresence=${this.showPresence}
          .avatarSize=${a.avatarSize}
          line2-property="email"
          class="signed-in-person"
        ></mgt-person>`}renderAccounts(){if(_.a.globalProvider.state===i.c.SignedIn&&_.a.globalProvider.isMultiAccountSupportedAndEnabled){const e=_.a.globalProvider,t=e.getAllAccounts();if((null==t?void 0:t.length)>1)return u.c`
          <div id="accounts">
            <ul
              tabindex="0"
              class="account-list"
              part="account-list"
              aria-label="${this.ariaLabel}"
              @keydown=${this.handleAccountListKeyDown}
            >
              ${t.filter(t=>t.id!==e.getActiveAccount().id).map(e=>{const t=localStorage.getItem(e.id+this._userDetailsKey);return m.a`
                    <li
                      tabindex="-1"
                      part="account-item"
                      class="account-item"
                      @click=${()=>this.setActiveAccount(e)}
                      @keyup=${t=>{"Enter"===t.key&&this.setActiveAccount(e)}}
                    >
                      <mgt-person
                        .personDetails=${t?JSON.parse(t):null}
                        .fallbackDetails=${{displayName:e.name,mail:e.mail}}
                        .view=${ae.a.twolines}
                        class="account"
                      ></mgt-person>
                    </li>`})}
            </ul>
          </div>
       `}}setActiveAccount(e){_.a.globalProvider.setActiveAccount(e)}clearState(){this.userDetails=null,this._image=null}renderSignedOutButtonContent(){return this.renderTemplate("signed-out-button-content",null)||u.c`
        <span>${this.strings.signInLinkSubtitle}</span>`}showFlyout(){const e=this.flyout;e&&e.open()}hideFlyout(){const e=this.flyout;e&&e.close()}}Pe([Object(f.b)({attribute:"user-details",type:Object}),Te("design:type",Object)],He.prototype,"userDetails",void 0),Pe([Object(f.b)({attribute:"show-presence",type:Boolean}),Te("design:type",Object)],He.prototype,"showPresence",void 0),Pe([Object(f.b)({attribute:"login-view",type:String}),Te("design:type",String)],He.prototype,"loginView",void 0),Pe([Object(f.c)(),Te("design:type",Boolean)],He.prototype,"_isFlyoutOpen",void 0),Pe([Object(f.c)(),Te("design:type",Object)],He.prototype,"_arrowKeyLocation",void 0);var Re,Ne=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};!function(e){e[e.any=0]="any",e[e.unified=1]="unified",e[e.security=2]="security",e[e.mailenabledsecurity=4]="mailenabledsecurity",e[e.distribution=8]="distribution"}(Re||(Re={}));const Be=()=>V.a.config.groups.invalidationPeriod||V.a.config.defaultInvalidationPeriod,je=()=>V.a.config.groups.isEnabled&&V.a.config.isEnabled,Ve=(e,t,n=10,a=Re.any,i="")=>Ne(void 0,void 0,void 0,function*(){const r="Group.Read.All";let o;const s=`${t||"*"}*${a}*${i}:${n}`;if(je()){o=V.a.getCache(K.a.groups,K.a.groups.stores.groupsQuery);const e=yield o.getValue(s);if(e&&Be()>Date.now()-e.timeCached&&e.top>=n)return e.groups.map(e=>JSON.parse(e)).slice(0,n+1)}let c,d="";const l=[];if(""!==t&&(d=`(startswith(displayName,'${t}') or startswith(mailNickname,'${t}') or startswith(mail,'${t}'))`),i&&(d+=`${t?" and ":""}${i}`),a!==Re.any){const t=e.createBatch(),i=[];Re.unified===(a&Re.unified)&&i.push("groupTypes/any(c:c+eq+'Unified')"),Re.security===(a&Re.security)&&i.push("(mailEnabled eq false and securityEnabled eq true)"),Re.mailenabledsecurity===(a&Re.mailenabledsecurity)&&i.push("(mailEnabled eq true and securityEnabled eq true)"),Re.distribution===(a&Re.distribution)&&i.push("(mailEnabled eq true and securityEnabled eq false)"),d=d?`${d} and `:"";for(const e of i)t.get(e,`/groups?$filter=${d+e}`,["Group.Read.All"]);try{c=yield t.executeAll();for(const e of i)if(c.get(e).content.value)for(const t of c.get(e).content.value)l.find(e=>e.id===t.id)||l.push(t)}catch(t){try{const t=[];for(const a of i)t.push(e.api("groups").filter(`${d} and ${a}`).top(n).count(!0).header("ConsistencyLevel","eventual").middlewareOptions(Object(b.a)(r)).get());return(yield Promise.all(t)).map(e=>e.value).reduce((e,t)=>e.concat(t),[])}catch(e){return[]}}}else if(0===l.length){const t=yield e.api("groups").filter(d).top(n).count(!0).header("ConsistencyLevel","eventual").middlewareOptions(Object(b.a)(r)).get();return je()&&t&&(yield o.putValue(s,{groups:t.value.map(e=>JSON.stringify(e)),top:n})),t?t.value:null}return l}),ze=(e,t,n)=>Ne(void 0,void 0,void 0,function*(){let a;if(je()){a=V.a.getCache(K.a.groups,K.a.groups.stores.groups);const e=yield a.getValue(t);if(e&&Be()>Date.now()-e.timeCached){const t=e.group?JSON.parse(e.group):null,a=n&&t?n.filter(e=>!Object.keys(t).includes(e)):null;if(!a||a.length<=1)return t}}let i=`/groups/${t}`;n&&(i=i+"?$select="+n.toString());const r=yield e.api(i).middlewareOptions(Object(b.a)("Group.Read.All")).get();return je()&&(yield a.putValue(t,{group:JSON.stringify(r)})),r}),Ge=(e,t,n="")=>Ne(void 0,void 0,void 0,function*(){if(!t||0===t.length)return[];const a=e.createBatch(),i={},r=[];let o;je()&&(o=V.a.getCache(K.a.groups,K.a.groups.stores.groups));for(const e of t){let t;if(i[e]=null,je()&&(t=yield o.getValue(e)),t&&Be()>Date.now()-t.timeCached)i[e]=t.group?JSON.parse(t.group):null;else if(""!==e){let t=`/groups/${e}`;n&&(t=`${t}?$filters=${n}`),a.get(e,t,["Group.Read.All"]),r.push(e)}}try{const e=yield a.executeAll();for(const n of t){const t=e.get(n);(null==t?void 0:t.content)&&(i[n]=t.content,je()&&(yield o.putValue(n,{group:JSON.stringify(t.content)})))}return Promise.all(Object.values(i))}catch(n){try{return t.filter(e=>r.includes(e)).forEach(t=>{i[t]=ze(e,t)}),je()&&(yield Promise.all(t.filter(e=>r.includes(e)).map(e=>Ne(void 0,void 0,void 0,function*(){return yield o.putValue(e,{group:JSON.stringify(yield i[e])})})))),Promise.all(Object.values(i))}catch(e){return[]}}});var Ke=n("/4Fm");const We=[u.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{--person-details-wrapper-width:0;width:auto}:host fluent-text-field{width:100%}:host fluent-text-field::part(root){display:flex;flex-wrap:wrap;height:auto;background:padding-box linear-gradient(var(--people-picker-input-background,var(--neutral-fill-input-rest)),var(--people-picker-input-background,var(--neutral-fill-input-rest))),border-box var(--people-picker-input-border-color,var(--neutral-stroke-input-rest))}:host fluent-text-field::part(root):hover{background:padding-box linear-gradient(var(--people-picker-input-hover-background,var(--neutral-fill-input-hover)),var(--people-picker-input-hover-background,var(--neutral-fill-input-hover))),border-box var(--people-picker-input-hover-border-color,var(--neutral-stroke-input-hover))}:host fluent-text-field::part(root):focus,:host fluent-text-field::part(root):focus-within{background:padding-box linear-gradient(var(--people-picker-input-focus-background,var(--neutral-fill-input-focus)),var(--people-picker-input-focus-background,var(--neutral-fill-input-focus))),border-box var(--people-picker-input-focus-border-color,var(--neutral-stroke-input-focus))}:host fluent-text-field::part(start){margin:unset}:host fluent-text-field::part(control){width:min-content;height:calc((var(--base-height-multiplier) + var(--density)) * var(--design-unit) * 1px);word-spacing:inherit;text-indent:inherit;letter-spacing:inherit}:host fluent-text-field::part(control)::placeholder{color:var(--people-picker-input-placeholder-text-color,var(--input-placeholder-rest))}:host fluent-text-field::part(control):hover::placeholder{color:var(--people-picker-input-placeholder-hover-text-color,var(--input-placeholder-hover))}:host fluent-text-field::part(control):focus-within::placeholder,:host fluent-text-field::part(control):focus::placeholder{color:var(--people-picker-input-placeholder-focus-text-color,var(--input-placeholder-filled))}:host fluent-text-field .search-icon{display:flex;padding-top:10px;padding-inline-start:10px}:host fluent-text-field .search-icon svg path{fill:var(--people-picker-search-icon-color,currentColor)}:host .selected-list{display:flex;flex-wrap:wrap;column-gap:5px;padding:unset;margin:0 5px}:host .selected-list-item{display:flex;column-gap:5px;border-radius:100px;margin-top:3px;background-color:var(--people-picker-selected-option-background-color,var(--person-background-color,var(--neutral-layer-3)))}:host .selected-list-item.highlighted{background-color:var(--people-picker-selected-option-highlight-background-color,var(--accent-fill-rest))}:host .selected-list-item-close-icon{display:flex;justify-content:center;align-items:center;padding-inline-end:8px;cursor:pointer}:host .selected-list-item-close-icon svg path{fill:var(--people-picker-remove-selected-close-icon-color,currentColor)}:host .selected-list-item-person{width:max-content}:host fluent-card{margin-top:4px;background-color:var(--people-picker-dropdown-background-color,var(--neutral-layer-card-container))}:host .searched-people-list{list-style:none;padding:4px;margin:auto}:host .searched-people-list-result{padding:4px;border-radius:4px;background:var(--people-picker-dropdown-result-background-color,var(--person-background-color,transparent))}:host .searched-people-list-result:hover{background:var(--people-picker-dropdown-result-hover-background-color,var(--person-background-color,var(--neutral-fill-rest)))}:host .searched-people-list-result:focus,:host .searched-people-list-result:focus-within{background:var(--people-picker-dropdown-result-focus-background-color,var(--person-background-color,var(--neutral-fill-rest)))}:host .message-parent{display:flex;place-content:center;flex-direction:row;padding:10px 15px;column-gap:5px}:host .message-parent .loading-text,:host .message-parent .search-error-text{font-style:normal;font-weight:600;font-size:14px;line-height:19px;text-align:center;color:var(--people-picker-no-results-text-color,var(--neutral-foreground-hint))}
`];var qe=n("P2Ap");const Qe={inputPlaceholderText:"Search for a name",maxSelectionsPlaceHolder:"Max contacts added",maxSelectionsAriaLabel:"Maximum contact selections reached",noResultsFound:"We didn't find any matches.",loadingMessage:"Loading...",selected:"selected",removeSelectedUser:"Remove ",selectContact:"select a contact",suggestionsTitle:"Suggested contacts"};var Ye=n("D4ZN"),Je=n("o2vq"),Xe=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},Ze=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)},$e=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const et=()=>{Object(T.a)(qe.a,U.a),Object(te.a)(),Object(d.c)(),Object(Ye.b)(),Object(A.b)("people-picker",tt)};class tt extends p.a{static get styles(){return We}get strings(){return Qe}get flyout(){return this.renderRoot.querySelector(".flyout")}get input(){return this.renderRoot.querySelector("fluent-text-field")}get groupId(){return this._groupId}set groupId(e){this._groupId!==e&&(this._groupId=e,this.requestStateUpdate(!0))}get groupIds(){return this._groupIds}set groupIds(e){Object(O.a)(this._groupIds,e)||(this._groupIds=e,this.requestStateUpdate(!0))}get type(){return this._type}set type(e){this._type!==e&&(this._type=e,this.requestStateUpdate(!0))}get groupType(){return this._groupType}set groupType(e){this._groupType!==e&&(this._groupType=e,this.requestStateUpdate(!0))}get userType(){return this._userType}set userType(e){this._userType!==e&&(this._userType=e,this.requestStateUpdate(!0))}get transitiveSearch(){return this._transitiveSearch}set transitiveSearch(e){this.transitiveSearch!==e&&(this._transitiveSearch=e,this.requestStateUpdate(!0))}get people(){return this._people}set people(e){Object(O.a)(this._people,e)||(this._people=e,this.requestStateUpdate(!0))}get showMax(){return this._showMax}set showMax(e){e!==this._showMax&&(this._showMax=e,this.requestStateUpdate(!0))}get selectedPeople(){return this._selectedPeople}set selectedPeople(e){e||(e=[]),Object(O.a)(this._selectedPeople,e)||(this._selectedPeople=e,this.fireCustomEvent("selectionChanged",this._selectedPeople),this.requestUpdate())}get defaultSelectedUserIds(){return this._defaultSelectedUserIds}set defaultSelectedUserIds(e){Object(O.a)(this._defaultSelectedUserIds,e)||(this._defaultSelectedUserIds=e,this.requestStateUpdate(!0))}get defaultSelectedGroupIds(){return this._defaultSelectedGroupIds}set defaultSelectedGroupIds(e){Object(O.a)(this._defaultSelectedGroupIds,e)||(this._defaultSelectedGroupIds=e,this.requestStateUpdate(!0))}get userIds(){return this._userIds}set userIds(e){Object(O.a)(this._userIds,e)||(this._userIds=e,this.requestStateUpdate(!0))}get userFilters(){return this._userFilters}set userFilters(e){this._userFilters=e,this.requestStateUpdate(!0)}get peopleFilters(){return this._peopleFilters}set peopleFilters(e){this._peopleFilters=e,this.requestStateUpdate(!0)}get groupFilters(){return this._groupFilters}set groupFilters(e){this._groupFilters=e,this.requestStateUpdate(!0)}static get requiredScopes(){return[...new Set(["user.read.all","people.read","group.read.all","user.readbasic.all",...d.a.requiredScopes])]}constructor(){super(),this._type=I.a.person,this._groupType=Re.any,this._userType=I.b.any,this._selectedPeople=[],this._arrowSelectionCount=-1,this.defaultSelectedUsers=[],this.defaultSelectedGroups=[],this._highlightedUsers=[],this._currentHighlightedUserPos=0,this._isFocused=!1,this._setAnyEmail=!1,this.handleSelectionChanged=()=>{0!==this.selectedPeople.length||this.disabled||this.enableTextInput()},this.handleInputClick=()=>{this.flyout.isOpen||this.handleUserSearch()},this.gainedFocus=()=>{this.clearHighlighted(),this._isFocused=!0,this.loadState(),this.showFlyout()},this.lostFocus=()=>{this._isFocused=!1,this.input&&this.input.setAttribute("aria-activedescendant","");const e=this.renderRoot.querySelector(".people-list");if(e)for(const t of e.children)t.classList.remove("focused"),t.setAttribute("aria-selected","false");this.requestUpdate()},this.onUserKeyUp=e=>{const t=e.key,n=e.getModifierState("Control")||e.getModifierState("Meta"),a=n&&"v"===t,i=["ArrowDown","ArrowRight","ArrowUp","ArrowLeft"].includes(t);return!a&&n||i?((n||["ArrowLeft","ArrowRight"].includes(t))&&this.hideFlyout(),void("ArrowDown"===t&&!this.flyout.isOpen&&this._isFocused&&this.handleUserSearch())):["Tab","Enter","Shift"].includes(t)?void 0:"Escape"===t?(this.clearInput(),this._foundPeople=[],void(this._arrowSelectionCount=-1)):"Backspace"===t&&0===this.userInput.length&&this.selectedPeople.length>0?(this.clearHighlighted(),this.selectedPeople=this.selectedPeople.splice(0,this.selectedPeople.length-1),this.loadState(),void this.hideFlyout()):void([";",","].includes(t)&&this.allowAnyEmail&&(this._setAnyEmail=!0,e.preventDefault(),e.stopPropagation()))},this.onUserInput=e=>{const t=e.target;this.userInput=t.value,this.userInput&&(Object(Ke.o)(this.userInput)&&this.allowAnyEmail?this._setAnyEmail&&this.handleAnyEmail():this.handleUserSearch(),this._setAnyEmail=!1)},this.onUserKeyDown=e=>{const t=e.key,n=this.renderRoot.querySelector(".selected-list"),a=e.getModifierState("Control")||e.getModifierState("Meta");if(a&&n){const e=n.querySelectorAll("mgt-person.selected-list-item-person");if(this.hideFlyout(),a&&"ArrowLeft"===t)this._currentHighlightedUserPos=(this._currentHighlightedUserPos-1+e.length)%e.length,this._currentHighlightedUserPos>=0&&!Number.isNaN(this._currentHighlightedUserPos)?this._highlightedUsers.push(e[this._currentHighlightedUserPos]):this._currentHighlightedUserPos=0;else if(a&&"ArrowRight"===t){const e=this._highlightedUsers.pop();if(e){const t=e.parentElement;t&&(this.clearHighlighted(t),this._currentHighlightedUserPos++)}}else a&&"a"===t&&(this._highlightedUsers=[],e.forEach(e=>this._highlightedUsers.push(e)));this._highlightedUsers&&this.highlightSelectedPeople(this._highlightedUsers)}else if(this.clearHighlighted(),this.flyout.isOpen){if("ArrowUp"!==t&&"ArrowDown"!==t||(this.handleArrowSelection(e),e.preventDefault()),"Enter"===t)if(!e.shiftKey&&this._foundPeople){e.preventDefault(),e.stopPropagation();const t=this._foundPeople[this._arrowSelectionCount];t&&(this.addPerson(t),this.hideFlyout(),this.input.value="",this.hasMaxSelections&&this.disableTextInput())}else this.allowAnyEmail?this.handleAnyEmail():this.showFlyout();"Escape"===t&&e.stopPropagation(),"Tab"===t&&this.hideFlyout(),[";",","].includes(t)&&this.allowAnyEmail&&(e.preventDefault(),e.stopPropagation(),this.userInput=this.input.value,this.handleAnyEmail())}},this.handleCut=()=>{this.writeHighlightedText().then(()=>{this.removeHighlightedOnCut()},()=>{})},this.handleCopy=()=>{this.writeHighlightedText()},this.handlePaste=()=>{navigator.clipboard.readText().then(e=>{if(e)try{const t=JSON.parse(e);if(t&&t.length>0)for(const e of t)this.addPerson(e)}catch(t){if(t instanceof SyntaxError){const t=[",",";"];let n;try{for(const a of t)if(n=e.split(a),n.length>1){this.hideFlyout(),this.selectUsersById(n);break}}catch(e){}}}},e=>{})},this.clearState(),this._showLoading=!0,this.showMax=6,this.disableImages=!1,this.disabled=!1,this.allowAnyEmail=!1,this.addEventListener("copy",this.handleCopy),this.addEventListener("cut",this.handleCut),this.addEventListener("paste",this.handlePaste),this.addEventListener("selectionChanged",this.handleSelectionChanged)}disableTextInput(){const e=this.input.shadowRoot.querySelector("input");e&&(e.setAttribute("disabled","true"),e.value="")}enableTextInput(){const e=this.input.shadowRoot.querySelector("input");e&&(e.removeAttribute("disabled"),e.focus())}get hasMaxSelections(){return"single"===this.selectionMode&&this.selectedPeople.length>=1}focus(e){this.input&&(this.input.focus(e),this.input.select())}selectUsersById(e){var t;return $e(this,void 0,void 0,function*(){const n=_.a.globalProvider,a=_.a.globalProvider.graph;if(n&&n.state===i.c.SignedIn)for(const n in e){const i=e[n];try{const e=yield Object(C.f)(a,i,d.b);this.addPerson(e)}catch(e){if(Object(Je.a)(e)&&(null===(t=e.message)||void 0===t?void 0:t.includes("does not exist"))&&this.allowAnyEmail&&Object(Ke.o)(i)){const e={mail:i,displayName:i};this.addPerson(e)}}}})}selectGroupsById(e){return $e(this,void 0,void 0,function*(){const t=_.a.globalProvider,n=_.a.globalProvider.graph;if(t&&t.state===i.c.SignedIn)for(const t in e)try{const a=yield ze(n,e[t]);this.addPerson(a)}catch(e){}})}render(){const e=this.renderTemplate("default",{people:this._foundPeople});if(e)return e;const t=this.renderSelectedPeople(this.selectedPeople),n=this.renderInput(t),a=this.renderFlyout(n);return u.c`
      <div>
        ${a}
      </div>
    `}clearState(){this.selectedPeople=[],this.userInput="",this._highlightedUsers=[],this._currentHighlightedUserPos=0}requestStateUpdate(e){return e&&(this._groupPeople=null,this._foundPeople=null,this.selectedPeople=[],this.defaultPeople=null),super.requestStateUpdate(e)}renderInput(e){var t,n,a;const i=this.disabled?"":this.placeholder||this.strings.inputPlaceholderText,r=this.hasMaxSelections?this.strings.maxSelectionsAriaLabel:"",o=u.c`<span class="search-icon">${Object(S.b)(S.a.Search)}</span>`,s=(null===(t=this.selectedPeople)||void 0===t?void 0:t.length)>0?e:o;return u.c`
      <fluent-text-field
        autocomplete="off"
        appearance="outline"
        slot="anchor"
        id="people-picker-input"
        role="combobox"
        placeholder=${this.hasMaxSelections?this.strings.maxSelectionsPlaceHolder:i}
        aria-label=${this.ariaLabel||r||i||this.strings.selectContact}
        aria-expanded=${null!==(a=null===(n=this.flyout)||void 0===n?void 0:n.isOpen)&&void 0!==a&&a}
        @click="${this.hasMaxSelections?void 0:this.handleInputClick}"
        @focus="${this.hasMaxSelections?void 0:this.gainedFocus}"
        @keydown="${this.hasMaxSelections?void 0:this.onUserKeyDown}"
        @input="${this.hasMaxSelections?void 0:this.onUserInput}"
        @blur="${this.lostFocus}"
        ?disabled=${this.disabled}
      >
        <span slot="start">${s}</span>
      </fluent-text-field>
    `}renderSelectedPeople(e){return(null==e?void 0:e.length)?u.c`
       <ul
        id="selected-list"
        aria-label="${this.strings.selected}"
        class="selected-list">
          ${Object(D.a)(e,e=>null==e?void 0:e.id,e=>{var t;return u.c`
            <li class="selected-list-item">
              ${this.renderTemplate("selected-person",{person:e},`selected-${(null==e?void 0:e.id)?e.id:e.displayName}`)||this.renderSelectedPerson(e)}

              <div
                role="button"
                tabindex="0"
                class="selected-list-item-close-icon"
                aria-label="${this.strings.removeSelectedUser}${null!==(t=null==e?void 0:e.displayName)&&void 0!==t?t:""}"
                @click="${t=>this.removePerson(e,t)}"
                @keydown="${t=>this.handleRemovePersonKeyDown(e,t)}">
                  ${Object(S.b)(S.a.Close)}
              </div>
          </li>`})}
      </ul>`:u.c``}renderFlyout(e){return m.a`
       <mgt-flyout light-dismiss class="flyout">
         ${e}
         <fluent-card
          tabindex="0"
          slot="flyout"
          class="flyout-root"
          @wheel=${e=>this.handleSectionScroll(e)}
          @keydown=${e=>this.onUserKeyDown(e)}
          class="custom">
           ${this.renderFlyoutContent()}
         </fluent-card>
       </mgt-flyout>
     `}renderFlyoutContent(){if(this.isLoadingState||this._showLoading)return this.renderLoading();const e=this._foundPeople;return e&&0!==e.length&&0!==this.showMax?this.renderSearchResults(e):this.renderNoData()}renderLoading(){return this.renderTemplate("loading",null)||m.a`
         <div class="message-parent">
           <mgt-spinner></mgt-spinner>
           <div aria-label="${this.strings.loadingMessage}" class="loading-text">
             ${this.strings.loadingMessage}
           </div>
         </div>
       `}renderNoData(){if(this._isFocused)return this.renderTemplate("error",null)||this.renderTemplate("no-data",null)||u.c`
         <div class="message-parent">
           <div aria-label=${this.strings.noResultsFound} class="search-error-text">
             ${this.strings.noResultsFound}
           </div>
         </div>
       `}renderSearchResults(e){const t=e.filter(e=>e.id);return u.c`
      <ul
        id="suggestions-list"
        class="searched-people-list"
        role="listbox"
        aria-live="polite"
        title=${this.strings.suggestionsTitle}
      >
         ${Object(D.a)(t,e=>e.id,e=>u.c`
            <li
              id="${e.id}"
              class="searched-people-list-result"
              role="option"
              @click="${()=>this.handleSuggestionClick(e)}">
                ${this.renderPersonResult(e)}
            </li>
          `)}
       </ul>
     `}renderPersonResult(e){return this.renderTemplate("person",{person:e},e.id)||m.a`
         <mgt-person
          class="person"
          ?show-presence=${this.showPresence}
          view="twoLines"
          line2-property="jobTitle,mail"
          .personDetails=${e}
          .fetchImage=${!this.disableImages}>
          .personCardInteraction=${w.a.none}
        </mgt-person>`}renderSelectedPerson(e){return m.a`
       <mgt-person
         tabindex="-1"
         class="selected-list-item-person"
         .personDetails=${e}
         .fetchImage=${!this.disableImages}
         .view=${ee.a.oneline}
         .personCardInteraction=${w.a.none}>
        </mgt-person>
     `}loadState(){var e,t,n,a;return $e(this,void 0,void 0,function*(){let r=this.people;const o=this.userInput.toLowerCase(),s=_.a.globalProvider;if(r){if(o){const e=r.filter(e=>null==e?void 0:e.displayName.toLowerCase().includes(o));r=e}this._showLoading=!1}else if(!r&&s&&s.state===i.c.SignedIn){const i=s.graph.forComponent(this);if(!o.length){if(this.defaultPeople)r=this.defaultPeople;else{if(this.groupId||this.groupIds){if(null===this._groupPeople)if(this.groupId)try{this.type===I.a.group?this._groupPeople=yield Object(C.a)(i,null,this.groupId,this.showMax,this.type,this.transitiveSearch):this._groupPeople=yield Object(C.a)(i,null,this.groupId,this.showMax,this.type,this.transitiveSearch,this.userFilters,this.peopleFilters)}catch(e){this._groupPeople=[]}else if(this.groupIds)if(this.type===I.a.group)try{this._groupPeople=yield Ge(i,this.groupIds,this.groupFilters)}catch(e){this._groupPeople=[]}else try{const e=yield Object(C.c)(i,"",this.groupIds,this.showMax,this.type,this.transitiveSearch,this.userFilters);this._groupPeople=e}catch(e){this._groupPeople=[]}r=this._groupPeople||[]}else if(this.type===I.a.person||this.type===I.a.any)if(this.userIds)r=yield Object(C.j)(i,this.userIds,"",this.userFilters);else{const e=this.userType===I.b.user||this.userType===I.b.contact;r=this._userFilters&&e?yield Object(C.h)(i,this._userFilters,this.showMax):yield Object(I.f)(i,this.userType,this._peopleFilters,this.showMax)}else if(this.type===I.a.group)if(this.groupIds)try{r=yield this.getGroupsForGroupIds(i,r)}catch(e){}else{let e=(yield Ve(i,"",this.showMax,this.groupType,this._groupFilters))||[];e.length>0&&e[0].value&&(e=e[0].value),r=e}this.defaultPeople=r}this._isFocused&&(this._showLoading=!1)}if(!((null===(e=this.defaultSelectedUserIds)||void 0===e?void 0:e.length)>0||(null===(t=this.defaultSelectedGroupIds)||void 0===t?void 0:t.length)>0)||this.selectedPeople.length||this.defaultSelectedUsers.length||this.defaultSelectedGroups.length||(this.defaultSelectedUsers=yield Object(C.j)(i,this.defaultSelectedUserIds,"",this.userFilters),this.defaultSelectedGroups=yield Ge(i,this.defaultSelectedGroupIds,this.peopleFilters),this.defaultSelectedGroups=this.defaultSelectedGroups.filter(e=>null!==e),this.defaultSelectedUsers=this.defaultSelectedUsers.filter(e=>null!==e),this.selectedPeople=[...this.defaultSelectedUsers,...this.defaultSelectedGroups],this.requestUpdate()),o)if(r=[],this.groupId)r=(yield Object(C.a)(i,o,this.groupId,this.showMax,this.type,this.transitiveSearch,this.userFilters,this.peopleFilters))||[];else{if(this.type===I.a.person||this.type===I.a.any){try{if(this.userType===I.b.contact||this.userType===I.b.user)r=(null===(n=this.userIds)||void 0===n?void 0:n.length)?yield Object(C.j)(i,this.userIds,o,this._userFilters):yield Object(C.b)(i,o,this.showMax,this._userFilters);else if(this.groupIds)try{r=yield Object(C.c)(i,o,this.groupIds,this.showMax,this.type,this.transitiveSearch,this.userFilters)}catch(e){}else r=(null===(a=this.userIds)||void 0===a?void 0:a.length)?yield Object(C.j)(i,this.userIds,o,this._userFilters):(yield Object(I.d)(i,o,this.showMax,this.userType,this._peopleFilters))||[]}catch(e){}if(r&&r.length<this.showMax&&this.userType!==I.b.contact&&this.type!==I.a.person)try{const e=(yield Object(C.b)(i,o,this.showMax,this._userFilters))||[],t=new Set(r.map(e=>e.id));for(const n of e)t.has(n.id)||r.push(n)}catch(e){}}if((this.type===I.a.group||this.type===I.a.any)&&r.length<this.showMax)if(this.groupIds)try{r=yield((e,t,n,a=10,i=Re.any,r="")=>Ne(void 0,void 0,void 0,function*(){const o=[],s=yield Ve(e,t,a,i,r);if(s)for(const e of s)e.id&&n.includes(e.id)&&o.push(e);return o}))(i,o,this.groupIds,this.showMax,this.groupType,this.userFilters)}catch(e){}else{let e=[];try{e=(yield Ve(i,o,this.showMax,this.groupType,this._groupFilters))||[],r=r.concat(e)}catch(e){}}}}this._foundPeople=this.filterPeople(r)})}getGroupsForGroupIds(e,t){return $e(this,void 0,void 0,function*(){const n=yield Ge(e,this.groupIds,this.groupFilters);for(const e of n)t=t.concat(e);return t=t.filter(e=>e)})}hideFlyout(){const e=this.flyout;e&&e.close(),this.input&&this.input.setAttribute("aria-activedescendant",""),this._arrowSelectionCount=-1}showFlyout(){const e=this.flyout;e&&e.open(),this._arrowSelectionCount=-1}removePerson(e,t){t.stopPropagation();const n=this.selectedPeople.filter(t=>!e.id&&t.displayName?t.displayName!==e.displayName:t.id!==e.id);this.hasMaxSelections&&this.enableTextInput(),this.selectedPeople=n,this.loadState()}handleRemovePersonKeyDown(e,t){"Enter"===t.key&&this.removePerson(e,t)}addPerson(e){e&&(setTimeout(()=>{this.clearInput()},50),0===this.selectedPeople.filter(t=>!e.id&&t.displayName?t.displayName===e.displayName:t.id===e.id).length&&(this.selectedPeople=[...this.selectedPeople,e],this.loadState(),this._foundPeople=[],this._arrowSelectionCount=-1))}clearInput(){this.clearHighlighted(),"single"!==this.selectionMode&&(this.input.value=""),this.userInput=""}handleAnyEmail(){if(this._showLoading=!1,this._arrowSelectionCount=-1,Object(Ke.o)(this.userInput)){const e={mail:this.userInput,displayName:this.userInput};this.addPerson(e)}this.hideFlyout(),this.input&&(this.input.focus(),this._isFocused=!0)}handleSuggestionClick(e){this.addPerson(e),this.hasMaxSelections&&(this.disableTextInput(),this.input.value=""),this.hideFlyout()}handleUserSearch(){this._debouncedSearch||(this._debouncedSearch=Object(Ke.b)(()=>$e(this,void 0,void 0,function*(){const e=setTimeout(()=>{this._showLoading=!0},50);yield this.loadState(),clearTimeout(e),this._showLoading=!1,this._arrowSelectionCount=-1,this.showFlyout()}),400)),this._debouncedSearch()}writeHighlightedText(){return $e(this,void 0,void 0,function*(){const e=[];for(const t of this._highlightedUsers){const{id:n,displayName:a,mail:i,userPrincipalName:r,scoredEmailAddresses:o}=t._personDetails;let s;s=o&&o.length>0?o.pop().address:r||i,e.push({id:n,displayName:a,email:s})}let t="";e.length>0&&(t=JSON.stringify(e)),yield navigator.clipboard.writeText(t)})}removeHighlightedOnCut(){this.selectedPeople=this.selectedPeople.splice(0,this.selectedPeople.length-this._highlightedUsers.length),this._highlightedUsers=[],this._currentHighlightedUserPos=0,this.loadState(),this.hideFlyout()}highlightSelectedPeople(e){for(const t of e)(null==t?void 0:t.parentElement).classList.add("highlighted")}clearHighlighted(e){if(e)e.classList.remove("highlighted");else{for(const e of this._highlightedUsers){const t=e.parentElement;t&&t.classList.remove("highlighted")}this._highlightedUsers=[],this._currentHighlightedUserPos=0}}handleArrowSelection(e){var t,n;const a=this.renderRoot.querySelector(".searched-people-list");if(null===(t=null==a?void 0:a.children)||void 0===t?void 0:t.length){e&&("ArrowUp"===e.key&&(-1===this._arrowSelectionCount?this._arrowSelectionCount=0:this._arrowSelectionCount=(this._arrowSelectionCount-1+a.children.length)%a.children.length),"ArrowDown"===e.key&&(-1===this._arrowSelectionCount?this._arrowSelectionCount=0:this._arrowSelectionCount=(this._arrowSelectionCount+1+a.children.length)%a.children.length));for(const e of null!==(n=null==a?void 0:a.children)&&void 0!==n?n:[]){const t=e;t.setAttribute("aria-selected","false"),t.blur(),t.removeAttribute("tabindex")}const t=a.children[this._arrowSelectionCount];t&&(t.setAttribute("tabindex","0"),t.focus(),t.scrollIntoView({behavior:"smooth",block:"nearest",inline:"nearest"}),t.setAttribute("aria-selected","true"),this.input.setAttribute("aria-activedescendant",null==t?void 0:t.id))}}filterPeople(e){const t=[];if(e&&e.length>0){e=e.filter(e=>e);const n=this.selectedPeople.map(e=>e.id?e.id:e.displayName),a=e.filter(e=>(null==e?void 0:e.id)?-1===n.indexOf(e.id):-1===n.indexOf(null==e?void 0:e.displayName)),i=new Set;for(const e of a){const t=JSON.stringify(e);i.add(t)}i.forEach(e=>{const n=JSON.parse(e);t.push(n)})}return t}handleSectionScroll(e){const t=this.renderRoot.querySelector(".flyout-root");t&&(e.deltaY<0&&0===t.scrollTop||e.deltaY>0&&t.clientHeight+t.scrollTop>=t.scrollHeight-1||e.stopPropagation())}}Xe([Object(f.b)({attribute:"group-id",converter:e=>e.trim()}),Ze("design:type",String),Ze("design:paramtypes",[Object])],tt.prototype,"groupId",null),Xe([Object(f.b)({attribute:"group-ids",converter:e=>e.split(",").map(e=>e.trim())}),Ze("design:type",Array),Ze("design:paramtypes",[Object])],tt.prototype,"groupIds",null),Xe([Object(f.b)({attribute:"type",converter:e=>(e=e.toLowerCase())&&0!==e.length?void 0===I.a[e]?I.a.any:I.a[e]:I.a.any}),Ze("design:type",Object),Ze("design:paramtypes",[Object])],tt.prototype,"type",null),Xe([Object(f.b)({attribute:"group-type",converter:e=>{if(!e||0===e.length)return Re.any;const t=e.split(","),n=[];for(let e of t)e=e.trim(),void 0!==Re[e]&&n.push(Re[e]);return 0===n.length?Re.any:n.reduce((e,t)=>e|t)}}),Ze("design:type",Number),Ze("design:paramtypes",[Object])],tt.prototype,"groupType",null),Xe([Object(f.b)({attribute:"user-type",converter:e=>(e=e.toLowerCase())&&void 0!==I.b[e]?I.b[e]:I.b.any}),Ze("design:type",String),Ze("design:paramtypes",[Object])],tt.prototype,"userType",null),Xe([Object(f.b)({attribute:"transitive-search",type:Boolean}),Ze("design:type",Boolean),Ze("design:paramtypes",[Boolean])],tt.prototype,"transitiveSearch",null),Xe([Object(f.b)({attribute:"people",type:Object}),Ze("design:type",Array),Ze("design:paramtypes",[Array])],tt.prototype,"people",null),Xe([Object(f.b)({attribute:"show-max",type:Number}),Ze("design:type",Number),Ze("design:paramtypes",[Number])],tt.prototype,"showMax",null),Xe([Object(f.b)({attribute:"disable-images",type:Boolean}),Ze("design:type",Boolean)],tt.prototype,"disableImages",void 0),Xe([Object(f.b)({attribute:"show-presence",type:Boolean}),Ze("design:type",Boolean)],tt.prototype,"showPresence",void 0),Xe([Object(f.b)({attribute:"selected-people",type:Array}),Ze("design:type",Array),Ze("design:paramtypes",[Array])],tt.prototype,"selectedPeople",null),Xe([Object(f.b)({attribute:"default-selected-user-ids",converter:e=>e.split(",").map(e=>e.trim()),type:String}),Ze("design:type",Array),Ze("design:paramtypes",[Object])],tt.prototype,"defaultSelectedUserIds",null),Xe([Object(f.b)({attribute:"default-selected-group-ids",converter:e=>e.split(",").map(e=>e.trim()),type:String}),Ze("design:type",Array),Ze("design:paramtypes",[Object])],tt.prototype,"defaultSelectedGroupIds",null),Xe([Object(f.b)({attribute:"placeholder",type:String}),Ze("design:type",String)],tt.prototype,"placeholder",void 0),Xe([Object(f.b)({attribute:"disabled",type:Boolean}),Ze("design:type",Boolean)],tt.prototype,"disabled",void 0),Xe([Object(f.b)({attribute:"allow-any-email",type:Boolean}),Ze("design:type",Boolean)],tt.prototype,"allowAnyEmail",void 0),Xe([Object(f.b)({attribute:"selection-mode",type:String}),Ze("design:type",String)],tt.prototype,"selectionMode",void 0),Xe([Object(f.b)({attribute:"user-ids",converter:e=>e.split(",").map(e=>e.trim()),type:String}),Ze("design:type",Array),Ze("design:paramtypes",[Array])],tt.prototype,"userIds",null),Xe([Object(f.b)({attribute:"user-filters"}),Ze("design:type",String),Ze("design:paramtypes",[String])],tt.prototype,"userFilters",null),Xe([Object(f.b)({attribute:"people-filters"}),Ze("design:type",String),Ze("design:paramtypes",[String])],tt.prototype,"peopleFilters",null),Xe([Object(f.b)({attribute:"group-filters"}),Ze("design:type",String),Ze("design:paramtypes",[String])],tt.prototype,"groupFilters",null),Xe([Object(f.b)({attribute:"aria-label",type:String}),Ze("design:type",String)],tt.prototype,"ariaLabel",void 0),Xe([Object(f.c)(),Ze("design:type",Boolean)],tt.prototype,"_showLoading",void 0),Xe([Object(f.c)(),Ze("design:type",Object)],tt.prototype,"_arrowSelectionCount",void 0),Xe([Object(f.c)(),Ze("design:type",Array)],tt.prototype,"_highlightedUsers",void 0),Xe([Object(f.c)(),Ze("design:type",Object)],tt.prototype,"_isFocused",void 0),Xe([Object(f.c)(),Ze("design:type",Object)],tt.prototype,"_setAnyEmail",void 0),Xe([Object(f.c)(),Ze("design:type",Array)],tt.prototype,"_foundPeople",void 0);const nt=[u.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{background-color:var(--tasks-background-color,transparent);padding:var(--tasks-padding,12px);border:var(--tasks-border,0);border-radius:var(--tasks-border-radius,0);font-family:var(--default-font-family);--skeleton-fill:var(--neutral-foreground-active)}:host .header{display:flex;align-items:center;justify-content:space-between;padding:var(--tasks-header-padding,0 0 14px 0);margin:var(--tasks-header-padding,0 0 14px 0);font-size:var(--tasks-header-font-size,16px);font-weight:var(--tasks-header-font-weight,600);color:var(--tasks-header-text-color,var(--neutral-foreground-hint))}:host .header:hover{color:var(--tasks-header-text-hover-color,var(--neutral-foreground-hover))}:host .header .title{justify-content:left;display:flex;align-items:center}:host .header .title .shimmer{width:80px;height:20px}:host .header .title svg{margin-top:3px;padding:0 4px;width:16px;height:16px}:host .header .new-task-button{justify-content:right}:host .header .new-task-button .shimmer{width:40px;height:24px}:host .header .new-task-button::part(control){font-weight:var(--tasks-new-button-text-font-weight,700);color:var(--tasks-new-button-text-color,var(--foreground-on-accent-rest));width:var(--tasks-new-button-width,none);height:var(--tasks-new-button-height,none);background:var(--tasks-new-button-background,padding-box linear-gradient(var(--accent-fill-rest),var(--accent-fill-rest)),border-box var(--accent-stroke-control-rest))}:host .header .new-task-button::part(control):hover{background:var(--tasks-new-button-background-hover,var(--accent-fill-hover))}:host .header .new-task-button::part(control):active{background:var(--tasks-new-button-background-active,var(--accent-fill-active))}:host .tasks{display:flex;flex-direction:column;row-gap:var(--tasks-gap,20px)}:host .tasks .task{display:flex;column-gap:4px;align-items:flex-start;justify-content:space-between}:host .tasks .task.updating{background:var(--neutral-stroke-rest)!important;pointer-events:none;opacity:.4}:host .tasks .task.complete{border:var(--task-complete-border,2px dotted var(--neutral-foreground-active));border-radius:var(--task-complete-border-radius,4px);background:var(--task-complete-background-color,transparent);padding:var(--task-complete-padding,10px)}:host .tasks .task.incomplete{border:var(--task-incomplete-border,1px solid var(--neutral-foreground-active));border-radius:var(--task-complete-border-radius,4px);background:var(--task-incomplete-background-color,var(--neutral-layer-1));padding:var(--task-incomplete-padding,10px)}:host .tasks .task .task-details-container{display:flex;flex-direction:column;row-gap:12px;width:100%}:host .tasks .task .task-details-container .top{display:flex;justify-content:space-between;column-gap:4px}:host .tasks .task .task-details-container .top.add-new-task{flex-direction:row}:host .tasks .task .task-details-container .top .check-and-title{display:flex;align-items:flex-start;flex-direction:column;width:100%;row-gap:12px}:host .tasks .task .task-details-container .top .check-and-title.shimmer{display:flex;flex-direction:inherit;gap:10px}:host .tasks .task .task-details-container .top .check-and-title.shimmer .checkbox{height:10px;width:10px}:host .tasks .task .task-details-container .top .check-and-title.shimmer .title{height:10px;width:100%}:host .tasks .task .task-details-container .top .check-and-title .task-content{display:grid;grid-template-columns:repeat(auto-fit,250px);gap:12px;justify-content:start;width:100%}:host .tasks .task .task-details-container .top .check-and-title .task-content .picker{max-width:250px}:host .tasks .task .task-details-container .top .check-and-title .task-content .task-due fluent-text-field.dark::part(control){color-scheme:dark}:host .tasks .task .task-details-container .top .task-options .options{height:10px;width:20px}:host .tasks .task .task-details-container .top .task-options.new-task-action-buttons{display:flex;flex-direction:column;row-gap:12px}:host .tasks .task .task-details-container .bottom{display:grid;grid-auto-flow:column;grid-auto-columns:1fr;grid-column-gap:6px}:host .tasks .task .task-details-container .bottom .task-bucket,:host .tasks .task .task-details-container .bottom .task-due,:host .tasks .task .task-details-container .bottom .task-group{display:flex;align-items:center;gap:6px}:host .tasks .task .task-details-container .bottom .task-bucket .task-icon,:host .tasks .task .task-details-container .bottom .task-due .task-icon,:host .tasks .task .task-details-container .bottom .task-group .task-icon{display:flex;place-content:center;width:var(--task-icons-width,20px);height:var(--task-icons-height,20px)}:host .tasks .task .task-details-container .bottom .task-bucket .task-icon .shimmer.icon,:host .tasks .task .task-details-container .bottom .task-due .task-icon .shimmer.icon,:host .tasks .task .task-details-container .bottom .task-group .task-icon .shimmer.icon{width:20px;height:20px}:host .tasks .task .task-details-container .bottom .task-bucket .task-icon svg,:host .tasks .task .task-details-container .bottom .task-due .task-icon svg,:host .tasks .task .task-details-container .bottom .task-group .task-icon svg{width:var(--task-icons-width,20px);height:var(--task-icons-height,20px)}:host .tasks .task .task-details-container .bottom .task-bucket .task-icon svg path,:host .tasks .task .task-details-container .bottom .task-due .task-icon svg path,:host .tasks .task .task-details-container .bottom .task-group .task-icon svg path{fill:var(--task-icons-background-color,var(--neutral-foreground-hint))}:host .tasks .task .task-details-container .bottom .task-bucket .task-icon-text,:host .tasks .task .task-details-container .bottom .task-due .task-icon-text,:host .tasks .task .task-details-container .bottom .task-group .task-icon-text{color:var(--task-icons-text-font-color,var(--neutral-foreground-hint));font-size:var(--task-icons-text-font-size,12px);font-weight:var(--task-icons-text-font-weight,600);white-space:normal;width:100%}:host .tasks .task .task-details-container .bottom .task-bucket .task-icon-text .shimmer.text,:host .tasks .task .task-details-container .bottom .task-due .task-icon-text .shimmer.text,:host .tasks .task .task-details-container .bottom .task-group .task-icon-text .shimmer.text{width:100%;height:10px}:host .tasks .task .task-details-container .task-details{display:flex;flex-direction:column;row-gap:8px}:host .tasks .task .task-details-container .task-details.shimmer{flex-direction:row;place-items:center;column-gap:6px}:host .tasks .task .task-details-container .task-details.shimmer .shimmer.icon{width:24px;height:24px}:host .tasks .task .task-details-container .task-details.shimmer .shimmer.text{width:100%;height:10px}:host .tasks .task .task-details-container .task-details .task-title{color:var(--foreground-on-neutral-rest)}:host .tasks .task .task-details-container .task-details .task-body{display:flex}:host fluent-button.add-task::part(control){font-weight:var(--task-add-new-button-text-font-weight,initial);color:var(--task-add-new-button-text-color,var(--neutral-foreground-rest));width:var(--task-add-new-button-width,none);height:var(--task-add-new-button-height,none);background:var(--task-add-new-button-background,padding-box linear-gradient(var(--neutral-fill-active),var(--neutral-fill-active)),border-box var(--neutral-stroke-control-active));border:var(--task-add-new-button-border,calc(var(--stroke-width) * 1px) solid transparent)}:host fluent-button.add-task::part(control):hover{background:var(--task-add-new-button-background-hover,padding-box linear-gradient(var(--neutral-fill-hover),var(--neutral-fill-hover)),border-box var(--neutral-stroke-control-hover))}:host fluent-button.add-task::part(control):active{background:var(--task-add-new-button-background-active,padding-box linear-gradient(var(--neutral-fill-active),var(--neutral-fill-active)),border-box var(--neutral-stroke-control-active))}:host fluent-button.cancel-task::part(control){font-weight:var(--task-cancel-new-button-text-font-weight,initial);color:var(--task-cancel-new-button-text-color,var(--neutral-foreground-rest));width:var(--task-cancel-new-button-width,none);height:var(--task-cancel-new-button-height,none);background:var(--task-cancel-new-button-background,padding-box linear-gradient(var(--neutral-fill-active),var(--neutral-fill-active)),border-box var(--neutral-stroke-control-active));border:var(--task-cancel-new-button-border,calc(var(--stroke-width) * 1px) solid transparent)}:host fluent-button.cancel-task::part(control):hover{background:var(--task-cancel-new-button-background-hover,padding-box linear-gradient(var(--neutral-fill-hover),var(--neutral-fill-hover)),border-box var(--neutral-stroke-control-hover))}:host fluent-button.cancel-task::part(control):active{background:var(--task-cancel-new-button-background-active,padding-box linear-gradient(var(--neutral-fill-active),var(--neutral-fill-active)),border-box var(--neutral-stroke-control-active))}:host fluent-option{background:var(--task-new-dropdown-list-background-color,var(--fill-color))}:host fluent-option:hover{background:var(--task-new-dropdown-option-hover-background-color,var(--neutral-fill-stealth-hover))}:host fluent-option::part(content){color:var(--task-new-dropdown-option-text-color,var(--neutral-foreground-rest))}:host fluent-select::part(listbox){background:var(--task-new-dropdown-list-background-color,var(--fill-color))}:host fluent-select::part(control){border:var(--task-new-dropdown-border,calc(var(--stroke-width) * 1px) solid transparent);border-radius:var(--task-new-dropdown-border-radius,calc(var(--control-corner-radius) * 1px));background:var(--task-new-dropdown-background-color,padding-box linear-gradient(var(--neutral-fill-input-rest),var(--neutral-fill-input-rest)),border-box var(--neutral-stroke-input-rest))}:host fluent-select::part(control):hover{background:var(--task-new-dropdown-hover-background-color,padding-box linear-gradient(var(--neutral-fill-input-hover),var(--neutral-fill-input-hover)),border-box var(--neutral-stroke-input-hover))}:host fluent-select::part(control)::placeholder{color:var(--task-new-dropdown-placeholder-color,var(--input-placeholder-rest))}:host fluent-checkbox{padding-top:1px}:host fluent-checkbox::part(control){border-radius:50%;background:var(--task-incomplete-checkbox-background-color,var(--neutral-fill-input-alt-rest))}:host fluent-checkbox::part(control):hover{background:var(--task-incomplete-checkbox-background-hover-color,var(--neutral-fill-input-alt-hover))}:host fluent-checkbox::part(label){font-size:var(--task-title-text-font-size,medium);font-weight:var(--task-title-text-font-weight,600);color:var(--task-incomplete-title-text-color,var(var(--neutral-foreground-rest)))}:host fluent-checkbox.checked::part(control){background:var(--task-complete-checkbox-background-color,var(--accent-fill-rest));color:var(--task-complete-checkbox-text-color,currentColor)}:host fluent-checkbox.checked::part(label){text-decoration:line-through;color:var(--task-complete-title-text-color,var(--neutral-foreground-hint))}:host fluent-text-field.new-task{width:100%}:host fluent-text-field.new-task::part(root){border:var(--task-new-input-border,calc(var(--stroke-width) * 1px) solid transparent);border-radius:var(--task-new-input-border-radius,calc(var(--control-corner-radius) * 1px));background:var(--task-new-input-background-color,padding-box linear-gradient(var(--neutral-fill-input-rest),var(--neutral-fill-input-rest)),border-box var(--neutral-stroke-input-rest))}:host fluent-text-field.new-task::part(root):hover{background:var(--task-new-input-hover-background-color,padding-box linear-gradient(var(--neutral-fill-input-hover),var(--neutral-fill-input-hover)),border-box var(--neutral-stroke-input-hover))}:host fluent-text-field.new-task::part(root)::placeholder{color:var(--task-new-input-placeholder-color,var(--input-placeholder-rest))}:host .people [slot=no-data] fluent-button::part(control){color:var(--task-new-person-icon-text-color,var(--neutral-foreground-hint))}:host .people [slot=no-data] svg{fill:var(--task-new-person-icon-color,var(--neutral-foreground-hint))}@media only screen and (width <= 600px){:host fluent-select{width:100%}:host .tasks .task .task-details-container .bottom{display:grid;grid-auto-flow:row;grid-auto-columns:unset;row-gap:4px}:host .tasks .task .task-details-container .bottom .ask-bucket,:host .tasks .task .task-details-container .bottom .ask-due,:host .tasks .task .task-details-container .bottom .ask-group{margin-inline-start:8px}:host .tasks .task .task-details-container .top.add-new-task{flex-direction:column;row-gap:12px}:host .tasks .task .task-details-container .top.add-new-task .check-and-title .task-content{display:grid;grid-auto-flow:row;row-gap:12px;width:100%}:host .tasks .task .task-details-container .top.add-new-task .task-options.new-task-action-buttons{flex-direction:row;column-gap:4px}}
`],at={removeTaskSubtitle:"Delete Task",cancelNewTaskSubtitle:"Cancel",newTaskPlaceholder:"Adding a task...",addTaskButtonSubtitle:"Add",due:"Due ",addTaskDate:"Add the task date",assign:"Assign"};var it=n("pnUA"),rt=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const ot=(e,t,n)=>rt(void 0,void 0,void 0,function*(){const a={assignments:n,appliedCategories:{category4:!0}};yield st(e,t,a)}),st=(e,t,n)=>rt(void 0,void 0,void 0,function*(){let a;try{a=yield e.api(`/planner/tasks/${t.id}`).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Group.ReadWrite.All")).header("Prefer","return=representation").header("If-Match",t.eTag).update(n)}catch(e){}return a});var ct=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const dt=(e,t,n,a)=>ct(void 0,void 0,void 0,function*(){return yield e.api(`/me/outlook/tasks/${t}`).header("Cache-Control","no-store").header("If-Match",a).middlewareOptions(Object(b.a)("Tasks.ReadWrite")).patch(n)});var lt=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};class ut{constructor(e){this.graph=it.a.fromGraph(e)}}class ft extends ut{getTaskGroups(){return lt(this,void 0,void 0,function*(){var e;return(yield(e=this.graph,rt(void 0,void 0,void 0,function*(){const t=yield e.api("/me/planner/plans").header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Group.Read.All")).get();return null==t?void 0:t.value}))).map(e=>{var t;return{id:e.id,title:e.title,containerId:null===(t=null==e?void 0:e.container)||void 0===t?void 0:t.containerId}})})}getTaskGroupsForGroup(e){return lt(this,void 0,void 0,function*(){var t,n;return(yield(t=this.graph,n=e,rt(void 0,void 0,void 0,function*(){const e=`/groups/${n}/planner/plans`,a=yield t.api(e).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Group.Read.All")).get();return null==a?void 0:a.value}))).map(e=>({id:e.id,title:e.title}))})}getTaskGroup(e){return lt(this,void 0,void 0,function*(){const t=yield(n=this.graph,a=e,rt(void 0,void 0,void 0,function*(){return yield n.api(`/planner/plans/${a}`).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Group.Read.All")).get()}));var n,a;return{id:t.id,title:t.title,_raw:t}})}getTaskFoldersForTaskGroup(e){return lt(this,void 0,void 0,function*(){var t,n;return(yield(t=this.graph,n=e,rt(void 0,void 0,void 0,function*(){const e=yield t.api(`/planner/plans/${n}/buckets`).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Group.Read.All")).get();return null==e?void 0:e.value}))).map(e=>({_raw:e,id:e.id,name:e.name,parentId:e.planId}))})}getTasksForTaskFolder(e){return lt(this,void 0,void 0,function*(){var t,n;return(yield(t=this.graph,n=e,rt(void 0,void 0,void 0,function*(){const e=yield t.api(`/planner/buckets/${n}/tasks`).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Group.Read.All")).get();return null==e?void 0:e.value}))).map(e=>({_raw:e,assignments:e.assignments,completed:100===e.percentComplete,dueDate:e.dueDateTime&&new Date(e.dueDateTime),eTag:e["@odata.etag"],id:e.id,immediateParentId:e.bucketId,name:e.title,topParentId:e.planId}))})}setTaskComplete(e){return lt(this,void 0,void 0,function*(){return yield((e,t)=>rt(void 0,void 0,void 0,function*(){yield st(e,t,{percentComplete:100})}))(this.graph,e)})}setTaskIncomplete(e){return lt(this,void 0,void 0,function*(){return((e,t)=>rt(void 0,void 0,void 0,function*(){yield st(e,t,{percentComplete:0})}))(this.graph,e)})}addTask(e){var t;return lt(this,void 0,void 0,function*(){return yield((e,t)=>rt(void 0,void 0,void 0,function*(){return yield e.api("/planner/tasks").header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Group.ReadWrite.All")).post(t)}))(this.graph,{assignments:e.assignments,bucketId:e.immediateParentId,dueDateTime:null===(t=e.dueDate)||void 0===t?void 0:t.toISOString(),planId:e.topParentId,title:e.name})})}assignPeopleToTask(e,t){return lt(this,void 0,void 0,function*(){return ot(this.graph,e,t)})}removeTask(e){return lt(this,void 0,void 0,function*(){return yield((e,t)=>rt(void 0,void 0,void 0,function*(){yield e.api(`/planner/tasks/${t.id}`).header("Cache-Control","no-store").header("If-Match",t.eTag).middlewareOptions(Object(b.a)("Group.ReadWrite.All")).delete()}))(this.graph,e)})}isAssignedToMe(e,t){return Object.keys(e.assignments).includes(t)}}class pt extends ut{getTaskGroups(){return lt(this,void 0,void 0,function*(){var e;return(yield(e=this.graph,ct(void 0,void 0,void 0,function*(){const t=yield e.api("/me/outlook/taskGroups").header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Tasks.Read")).get();return null==t?void 0:t.value}))).map(e=>({_raw:e,id:e.id,secondaryId:e.groupKey,title:e.name}))})}getTaskGroup(e){return lt(this,void 0,void 0,function*(){const t=yield(n=this.graph,a=e,ct(void 0,void 0,void 0,function*(){return yield n.api(`/me/outlook/taskGroups/${a}`).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Tasks.Read")).get()}));var n,a;return{id:t.id,secondaryId:t.groupKey,title:t.name,_raw:t}})}getTaskFoldersForTaskGroup(e){return lt(this,void 0,void 0,function*(){var t,n;return(yield(t=this.graph,n=e,ct(void 0,void 0,void 0,function*(){const e=yield t.api(`/me/outlook/taskGroups/${n}/taskFolders`).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Tasks.Read")).get();return null==e?void 0:e.value}))).map(t=>({_raw:t,id:t.id,name:t.name,parentId:e}))})}getTasksForTaskFolder(e,t){return lt(this,void 0,void 0,function*(){var n,a;return(yield(n=this.graph,a=e,ct(void 0,void 0,void 0,function*(){const e=yield n.api(`/me/outlook/taskFolders/${a}/tasks`).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Tasks.Read")).get();return null==e?void 0:e.value}))).map(n=>({_raw:n,assignments:{},completed:!!n.completedDateTime,dueDate:n.dueDateTime&&new Date(n.dueDateTime.dateTime+"Z"),eTag:n["@odata.etag"],id:n.id,immediateParentId:e,name:n.subject,topParentId:t}))})}setTaskComplete(e){return lt(this,void 0,void 0,function*(){var t,n,a;yield(t=this.graph,n=e.id,a=e.eTag,ct(void 0,void 0,void 0,function*(){return yield dt(t,n,{isReminderOn:!1,status:"completed"},a)}))})}assignPeopleToTask(e,t){return lt(this,void 0,void 0,function*(){return yield ot(this.graph,e,t)})}setTaskIncomplete(e){return lt(this,void 0,void 0,function*(){var t,n,a;yield(t=this.graph,n=e.id,a=e.eTag,ct(void 0,void 0,void 0,function*(){return yield dt(t,n,{isReminderOn:!0,status:"notStarted"},a)}))})}addTask(e){return lt(this,void 0,void 0,function*(){const t={parentFolderId:e.immediateParentId,subject:e.name};return e.dueDate&&(t.dueDateTime={dateTime:e.dueDate.toISOString(),timeZone:"UTC"}),yield((e,t)=>ct(void 0,void 0,void 0,function*(){const{parentFolderId:n=null}=t;return n?yield e.api(`/me/outlook/taskFolders/${n}/tasks`).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Tasks.ReadWrite")).post(t):yield e.api("/me/outlook/tasks").header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Tasks.ReadWrite")).post(t)}))(this.graph,t)})}removeTask(e){return lt(this,void 0,void 0,function*(){return yield(t=this.graph,n=e.id,a=e.eTag,ct(void 0,void 0,void 0,function*(){yield t.api(`/me/outlook/tasks/${n}`).header("Cache-Control","no-store").header("If-Match",a).middlewareOptions(Object(b.a)("Tasks.ReadWrite")).delete()}));var t,n,a})}isAssignedToMe(e,t){return Object.keys(e.assignments).includes(t)}getTaskGroupsForGroup(e){return lt(this,void 0,void 0,function*(){return Promise.resolve(void 0)})}}var mt=n("Kct4"),_t=n("i2HT");const ht=e=>{const t=Ee.o.getValueFor(e);return Object(mt.a)(_t.a.create(t,t,t))};var bt=n("3rEL"),gt=n("o87Z");class vt extends Ce{}class yt extends(Object(gt.b)(vt)){constructor(){super(...arguments),this.proxy=document.createElement("select")}}const St="above",Dt="below";class It extends yt{constructor(){super(...arguments),this.open=!1,this.forcedPosition=!1,this.listboxId=Object(le.a)("listbox-"),this.maxHeight=0}openChanged(e,t){if(this.collapsible){if(this.open)return this.ariaControls=this.listboxId,this.ariaExpanded="true",this.setPositioning(),this.focusAndScrollOptionIntoView(),this.indexWhenOpened=this.selectedIndex,void Ie.a.queueUpdate(()=>this.focus());this.ariaControls="",this.ariaExpanded="false"}}get collapsible(){return!(this.multiple||"number"==typeof this.size)}get value(){return se.b.track(this,"value"),this._value}set value(e){var t,n,a,i,r,o,s;const c=`${this._value}`;if(null===(t=this._options)||void 0===t?void 0:t.length){const t=this._options.findIndex(t=>t.value===e),c=null!==(a=null===(n=this._options[this.selectedIndex])||void 0===n?void 0:n.value)&&void 0!==a?a:null,d=null!==(r=null===(i=this._options[t])||void 0===i?void 0:i.value)&&void 0!==r?r:null;-1!==t&&c===d||(e="",this.selectedIndex=t),e=null!==(s=null===(o=this.firstSelectedOption)||void 0===o?void 0:o.value)&&void 0!==s?s:e}c!==e&&(this._value=e,super.valueChanged(c,e),se.b.notify(this,"value"),this.updateDisplayValue())}updateValue(e){var t,n;this.$fastController.isConnected&&(this.value=null!==(n=null===(t=this.firstSelectedOption)||void 0===t?void 0:t.value)&&void 0!==n?n:""),e&&(this.$emit("input"),this.$emit("change",this,{bubbles:!0,composed:void 0}))}selectedIndexChanged(e,t){super.selectedIndexChanged(e,t),this.updateValue()}positionChanged(e,t){this.positionAttribute=t,this.setPositioning()}setPositioning(){const e=this.getBoundingClientRect(),t=window.innerHeight-e.bottom;this.position=this.forcedPosition?this.positionAttribute:e.top>t?St:Dt,this.positionAttribute=this.forcedPosition?this.positionAttribute:this.position,this.maxHeight=this.position===St?~~e.top:~~t}get displayValue(){var e,t;return se.b.track(this,"displayValue"),null!==(t=null===(e=this.firstSelectedOption)||void 0===e?void 0:e.text)&&void 0!==t?t:""}disabledChanged(e,t){super.disabledChanged&&super.disabledChanged(e,t),this.ariaDisabled=this.disabled?"true":"false"}formResetCallback(){this.setProxyOptions(),super.setDefaultSelectedOption(),-1===this.selectedIndex&&(this.selectedIndex=0)}clickHandler(e){if(!this.disabled){if(this.open){const t=e.target.closest("option,[role=option]");if(t&&t.disabled)return}return super.clickHandler(e),this.open=this.collapsible&&!this.open,this.open||this.indexWhenOpened===this.selectedIndex||this.updateValue(!0),!0}}focusoutHandler(e){var t;if(super.focusoutHandler(e),!this.open)return!0;const n=e.relatedTarget;this.isSameNode(n)?this.focus():(null===(t=this.options)||void 0===t?void 0:t.includes(n))||(this.open=!1,this.indexWhenOpened!==this.selectedIndex&&this.updateValue(!0))}handleChange(e,t){super.handleChange(e,t),"value"===t&&this.updateValue()}slottedOptionsChanged(e,t){this.options.forEach(e=>{se.b.getNotifier(e).unsubscribe(this,"value")}),super.slottedOptionsChanged(e,t),this.options.forEach(e=>{se.b.getNotifier(e).subscribe(this,"value")}),this.setProxyOptions(),this.updateValue()}mousedownHandler(e){var t;return e.offsetX>=0&&e.offsetX<=(null===(t=this.listbox)||void 0===t?void 0:t.scrollWidth)?super.mousedownHandler(e):this.collapsible}multipleChanged(e,t){super.multipleChanged(e,t),this.proxy&&(this.proxy.multiple=t)}selectedOptionsChanged(e,t){var n;super.selectedOptionsChanged(e,t),null===(n=this.options)||void 0===n||n.forEach((e,t)=>{var n;const a=null===(n=this.proxy)||void 0===n?void 0:n.options.item(t);a&&(a.selected=e.selected)})}setDefaultSelectedOption(){var e;const t=null!==(e=this.options)&&void 0!==e?e:Array.from(this.children).filter(ve.slottedOptionFilter),n=null==t?void 0:t.findIndex(e=>e.hasAttribute("selected")||e.selected||e.value===this.value);this.selectedIndex=-1===n?0:n}setProxyOptions(){this.proxy instanceof HTMLSelectElement&&this.options&&(this.proxy.options.length=0,this.options.forEach(e=>{const t=e.proxy||(e instanceof HTMLOptionElement?e.cloneNode():null);t&&this.proxy.options.add(t)}))}keydownHandler(e){super.keydownHandler(e);const t=e.key||e.key.charCodeAt(0);switch(t){case de.m:e.preventDefault(),this.collapsible&&this.typeAheadExpired&&(this.open=!this.open);break;case de.j:case de.f:e.preventDefault();break;case de.g:e.preventDefault(),this.open=!this.open;break;case de.h:this.collapsible&&this.open&&(e.preventDefault(),this.open=!1);break;case de.n:return this.collapsible&&this.open&&(e.preventDefault(),this.open=!1),!0}return this.open||this.indexWhenOpened===this.selectedIndex||(this.updateValue(!0),this.indexWhenOpened=this.selectedIndex),!(t===de.b||t===de.e)}connectedCallback(){super.connectedCallback(),this.forcedPosition=!!this.positionAttribute,this.addEventListener("contentchange",this.updateDisplayValue)}disconnectedCallback(){this.removeEventListener("contentchange",this.updateDisplayValue),super.disconnectedCallback()}sizeChanged(e,t){super.sizeChanged(e,t),this.proxy&&(this.proxy.size=t)}updateDisplayValue(){this.collapsible&&se.b.notify(this,"displayValue")}}Object(oe.a)([Object(ce.c)({attribute:"open",mode:"boolean"})],It.prototype,"open",void 0),Object(oe.a)([se.e],It.prototype,"collapsible",null),Object(oe.a)([se.d],It.prototype,"control",void 0),Object(oe.a)([Object(ce.c)({attribute:"position"})],It.prototype,"positionAttribute",void 0),Object(oe.a)([se.d],It.prototype,"position",void 0),Object(oe.a)([se.d],It.prototype,"maxHeight",void 0);class xt{}Object(oe.a)([se.d],xt.prototype,"ariaControls",void 0),Object(_e.a)(xt,ye),Object(_e.a)(It,me.a,xt);var Ct=n("6vBc"),Ot=n("+53S"),wt=n("wHpb"),Et=n("2i1/"),At=n("oqLQ"),Lt=n("j9Xn"),kt=n("cQsl"),Mt=n("QkLF"),Pt=n("57ob"),Tt=n("NLdk"),Ut=n("81Gc"),Ft=n("KKsr");const Ht=".control",Rt=":not([disabled]):not([open])",Nt="[disabled]",Bt=(e,t)=>Oe.a`
    ${Object(we.a)("inline-flex")}
    
    :host {
      border-radius: calc(${Ee.q} * 1px);
      box-sizing: border-box;
      color: ${Ee.fb};
      fill: currentcolor;
      font-family: ${Ee.p};
      position: relative;
      user-select: none;
      min-width: 250px;
      vertical-align: top;
    }

    .listbox {
      box-shadow: ${kt.c};
      background: ${Ee.v};
      border-radius: calc(${Ee.D} * 1px);
      box-sizing: border-box;
      display: inline-flex;
      flex-direction: column;
      left: 0;
      max-height: calc(var(--max-height) - (${Mt.a} * 1px));
      padding: calc((${Ee.s} - ${Ee.vb} ) * 1px);
      overflow-y: auto;
      position: absolute;
      width: 100%;
      z-index: 1;
      margin: 1px 0;
      border: calc(${Ee.vb} * 1px) solid transparent;
    }

    .listbox[hidden] {
      display: none;
    }

    .control {
      border: calc(${Ee.vb} * 1px) solid transparent;
      border-radius: calc(${Ee.q} * 1px);
      height: calc(${Mt.a} * 1px);
      align-items: center;
      box-sizing: border-box;
      cursor: pointer;
      display: flex;
      ${Tt.a}
      min-height: 100%;
      padding: 0 calc(${Ee.s} * 2.25px);
      width: 100%;
    }

    :host(:${wt.a}) {
      ${Ae.a}
    }

    :host([disabled]) .control {
      cursor: ${Et.a};
      opacity: ${Ee.u};
      user-select: none;
    }

    :host([open][position='above']) .listbox {
      bottom: calc((${Mt.a} + ${Ee.s} * 2) * 1px);
    }

    :host([open][position='below']) .listbox {
      top: calc((${Mt.a} + ${Ee.s} * 2) * 1px);
    }

    .selected-value {
      font-family: inherit;
      flex: 1 1 auto;
      text-align: start;
    }

    .indicator {
      flex: 0 0 auto;
      margin-inline-start: 1em;
    }

    slot[name='listbox'] {
      display: none;
      width: 100%;
    }

    :host([open]) slot[name='listbox'] {
      display: flex;
      position: absolute;
    }

    .start {
      margin-inline-end: 11px;
    }

    .end {
      margin-inline-start: 11px;
    }

    .start,
    .end,
    .indicator,
    ::slotted(svg) {
      display: flex;
    }

    ::slotted([role='option']) {
      flex: 0 0 auto;
    }
  `;class jt extends It{appearanceChanged(e,t){e!==t&&(this.classList.add(t),this.classList.remove(e))}connectedCallback(){super.connectedCallback(),this.appearance||(this.appearance="outline"),this.listbox&&Ee.v.setValueFor(this.listbox,Ee.gb)}}Object(bt.a)([Object(ce.c)({mode:"fromView"})],jt.prototype,"appearance",void 0);const Vt=jt.compose({baseName:"select",baseClass:It,template:(e,t)=>Se.a`
    <template
        class="${e=>[e.collapsible&&"collapsible",e.collapsible&&e.open&&"open",e.disabled&&"disabled",e.collapsible&&e.position].filter(Boolean).join(" ")}"
        aria-activedescendant="${e=>e.ariaActiveDescendant}"
        aria-controls="${e=>e.ariaControls}"
        aria-disabled="${e=>e.ariaDisabled}"
        aria-expanded="${e=>e.ariaExpanded}"
        aria-haspopup="${e=>e.collapsible?"listbox":null}"
        aria-multiselectable="${e=>e.ariaMultiSelectable}"
        ?open="${e=>e.open}"
        role="combobox"
        tabindex="${e=>e.disabled?null:"0"}"
        @click="${(e,t)=>e.clickHandler(t.event)}"
        @focusin="${(e,t)=>e.focusinHandler(t.event)}"
        @focusout="${(e,t)=>e.focusoutHandler(t.event)}"
        @keydown="${(e,t)=>e.keydownHandler(t.event)}"
        @mousedown="${(e,t)=>e.mousedownHandler(t.event)}"
    >
        ${Object(Ct.a)(e=>e.collapsible,Se.a`
                <div
                    class="control"
                    part="control"
                    ?disabled="${e=>e.disabled}"
                    ${Object(Ot.a)("control")}
                >
                    ${Object(me.d)(e,t)}
                    <slot name="button-container">
                        <div class="selected-value" part="selected-value">
                            <slot name="selected-value">${e=>e.displayValue}</slot>
                        </div>
                        <div aria-hidden="true" class="indicator" part="indicator">
                            <slot name="indicator">
                                ${t.indicator||""}
                            </slot>
                        </div>
                    </slot>
                    ${Object(me.b)(e,t)}
                </div>
            `)}
        <div
            class="listbox"
            id="${e=>e.listboxId}"
            part="listbox"
            role="listbox"
            ?disabled="${e=>e.disabled}"
            ?hidden="${e=>!!e.collapsible&&!e.open}"
            ${Object(Ot.a)("listbox")}
        >
            <slot
                ${Object(De.a)({filter:ve.slottedOptionFilter,flatten:!0,property:"slottedOptions"})}
            ></slot>
        </div>
    </template>
`,styles:(e,t)=>Bt().withBehaviors(Object(Pt.a)("outline",Object(Ut.c)(e,t,Rt,Nt)),Object(Pt.a)("filled",Object(Ft.b)(e,t,Ht,Rt).withBehaviors(Object(At.a)(Object(Ft.c)(e,t,Ht,Rt)))),Object(Pt.a)("stealth",Object(Ut.e)(e,t,Rt,Nt)),Object(At.a)(Oe.a`
    :host([open]) .listbox {
      background: ${Lt.a.ButtonFace};
      border-color: ${Lt.a.CanvasText};
    }
  `)),indicator:'\n    <svg width="12" height="12" xmlns="http://www.w3.org/2000/svg">\n      <path d="M2.15 4.65c.2-.2.5-.2.7 0L6 7.79l3.15-3.14a.5.5 0 11.7.7l-3.5 3.5a.5.5 0 01-.7 0l-3.5-3.5a.5.5 0 010-.7z"/>\n    </svg>\n  '});class zt{constructor(e,t){this.cache=new WeakMap,this.ltr=e,this.rtl=t}bind(e){this.attach(e)}unbind(e){const t=this.cache.get(e);t&&Ee.t.unsubscribe(t)}attach(e){const t=this.cache.get(e)||new Gt(this.ltr,this.rtl,e),n=Ee.t.getValueFor(e);Ee.t.subscribe(t),t.attach(n),this.cache.set(e,t)}}class Gt{constructor(e,t,n){this.ltr=e,this.rtl=t,this.source=n,this.attached=null}handleChange({target:e,token:t}){this.attach(t.getValueFor(this.source))}attach(e){this.attached!==this[e]&&(null!==this.attached&&this.source.$fastController.removeStyles(this.attached),this.attached=this[e],null!==this.attached&&this.source.$fastController.addStyles(this.attached))}}const Kt=be.compose({baseName:"option",template:(e,t)=>Se.a`
    <template
        aria-checked="${e=>e.ariaChecked}"
        aria-disabled="${e=>e.ariaDisabled}"
        aria-posinset="${e=>e.ariaPosInSet}"
        aria-selected="${e=>e.ariaSelected}"
        aria-setsize="${e=>e.ariaSetSize}"
        class="${e=>[e.checked&&"checked",e.selected&&"selected",e.disabled&&"disabled"].filter(Boolean).join(" ")}"
        role="option"
    >
        ${Object(me.d)(e,t)}
        <span class="content" part="content">
            <slot ${Object(De.a)("content")}></slot>
        </span>
        ${Object(me.b)(e,t)}
    </template>
`,styles:(e,t)=>Oe.a`
    ${Object(we.a)("inline-flex")} :host {
      position: relative;
      ${Tt.a}
      background: ${Ee.ab};
      border-radius: calc(${Ee.q} * 1px);
      border: calc(${Ee.vb} * 1px) solid transparent;
      box-sizing: border-box;
      color: ${Ee.fb};
      cursor: pointer;
      fill: currentcolor;
      height: calc(${Mt.a} * 1px);
      overflow: hidden;
      align-items: center;
      padding: 0 calc(((${Ee.s} * 3) - ${Ee.vb} - 1) * 1px);
      user-select: none;
      white-space: nowrap;
    }

    :host::before {
      content: '';
      display: block;
      position: absolute;
      left: calc((${Ee.y} - ${Ee.vb}) * 1px);
      top: calc((${Mt.a} / 4) - ${Ee.y} * 1px);
      width: 3px;
      height: calc((${Mt.a} / 2) * 1px);
      background: transparent;
      border-radius: calc(${Ee.q} * 1px);
    }

    :host(:not([disabled]):hover) {
      background: ${Ee.Y};
    }

    :host(:not([disabled]):active) {
      background: ${Ee.W};
    }

    :host(:not([disabled]):active)::before {
      background: ${Ee.e};
      height: calc(((${Mt.a} / 2) - 6) * 1px);
    }

    :host([aria-selected='true'])::before {
      background: ${Ee.e};
    }

    :host(:${wt.a}) {
      ${Ae.a}
      background: ${Ee.X};
    }

    :host([aria-selected='true']) {
      background: ${Ee.V};
    }

    :host(:not([disabled])[aria-selected='true']:hover) {
      background: ${Ee.T};
    }

    :host(:not([disabled])[aria-selected='true']:active) {
      background: ${Ee.R};
    }

    :host(:not([disabled]):not([aria-selected='true']):hover) {
      background: ${Ee.Y};
    }

    :host(:not([disabled]):not([aria-selected='true']):active) {
      background: ${Ee.W};
    }

    :host([disabled]) {
      cursor: ${Et.a};
      opacity: ${Ee.u};
    }

    .content {
      grid-column-start: 2;
      justify-self: start;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .start,
    .end,
    ::slotted(svg) {
      display: flex;
    }

    ::slotted([slot='end']) {
      margin-inline-start: 1ch;
    }

    ::slotted([slot='start']) {
      margin-inline-end: 1ch;
    }
  `.withBehaviors(new zt(null,Oe.a`
      :host::before {
        right: calc((${Ee.y} - ${Ee.vb}) * 1px);
      }
    `),Object(At.a)(Oe.a`
        :host {
          background: ${Lt.a.ButtonFace};
          border-color: ${Lt.a.ButtonFace};
          color: ${Lt.a.ButtonText};
        }
        :host(:not([disabled]):not([aria-selected="true"]):hover),
        :host(:not([disabled])[aria-selected="true"]:hover),
        :host([aria-selected="true"]) {
          forced-color-adjust: none;
          background: ${Lt.a.Highlight};
          color: ${Lt.a.HighlightText};
        }
        :host(:not([disabled]):active)::before,
        :host([aria-selected='true'])::before {
          background: ${Lt.a.HighlightText};
        }
        :host([disabled]),
        :host([disabled]:not([aria-selected='true']):hover) {
          background: ${Lt.a.Canvas};
          color: ${Lt.a.GrayText};
          fill: currentcolor;
          opacity: 1;
        }
        :host(:${wt.a}) {
          outline-color: ${Lt.a.CanvasText};
        }
      `))});var Wt=n("kd7Q");class qt extends ue.a{constructor(){super(...arguments),this.shape="rect"}}Object(oe.a)([ce.c],qt.prototype,"fill",void 0),Object(oe.a)([ce.c],qt.prototype,"shape",void 0),Object(oe.a)([ce.c],qt.prototype,"pattern",void 0),Object(oe.a)([Object(ce.c)({mode:"boolean"})],qt.prototype,"shimmer",void 0);const Qt=qt.compose({baseName:"skeleton",template:(e,t)=>Se.a`
    <template
        class="${e=>"circle"===e.shape?"circle":"rect"}"
        pattern="${e=>e.pattern}"
        ?shimmer="${e=>e.shimmer}"
    >
        ${Object(Ct.a)(e=>!0===e.shimmer,Se.a`
                <span class="shimmer"></span>
            `)}
        <object type="image/svg+xml" data="${e=>e.pattern}" role="presentation">
            <img class="pattern" src="${e=>e.pattern}" />
        </object>
        <slot></slot>
    </template>
`,styles:(e,t)=>Oe.a`
    ${Object(we.a)("block")} :host {
      --skeleton-fill-default: ${Ee.V};
      overflow: hidden;
      width: 100%;
      position: relative;
      background-color: var(--skeleton-fill, var(--skeleton-fill-default));
      --skeleton-animation-gradient-default: linear-gradient(
        270deg,
        var(--skeleton-fill, var(--skeleton-fill-default)) 0%,
        ${Ee.T} 51%,
        var(--skeleton-fill, var(--skeleton-fill-default)) 100%
      );
      --skeleton-animation-timing-default: ease-in-out;
    }

    :host(.rect) {
      border-radius: calc(${Ee.q} * 1px);
    }

    :host(.circle) {
      border-radius: 100%;
      overflow: hidden;
    }

    object {
      position: absolute;
      width: 100%;
      height: auto;
      z-index: 2;
    }

    object img {
      width: 100%;
      height: auto;
    }

    ${Object(we.a)("block")} span.shimmer {
      position: absolute;
      width: 100%;
      height: 100%;
      background-image: var(--skeleton-animation-gradient, var(--skeleton-animation-gradient-default));
      background-size: 0px 0px / 90% 100%;
      background-repeat: no-repeat;
      background-color: var(--skeleton-animation-fill, ${Ee.V});
      animation: shimmer 2s infinite;
      animation-timing-function: var(--skeleton-animation-timing, var(--skeleton-timing-default));
      animation-direction: normal;
      z-index: 1;
    }

    ::slotted(svg) {
      z-index: 2;
    }

    ::slotted(.pattern) {
      width: 100%;
      height: 100%;
    }

    @keyframes shimmer {
      0% {
        transform: translateX(-100%);
      }
      100% {
        transform: translateX(100%);
      }
    }
  `.withBehaviors(Object(At.a)(Oe.a`
        :host{
          background-color: ${Lt.a.CanvasText};
        }
      `))}),Yt="menuitem",Jt="menuitemcheckbox",Xt="menuitemradio",Zt={[Yt]:"menuitem",[Jt]:"menuitemcheckbox",[Xt]:"menuitemradio"};var $t=n("qibw");const en=e=>{const t=e.closest("[dir]");return null!==t&&"rtl"===t.dir?$t.a.rtl:$t.a.ltr};class tn extends ue.a{constructor(){super(...arguments),this.role=Yt,this.hasSubmenu=!1,this.currentDirection=$t.a.ltr,this.focusSubmenuOnLoad=!1,this.handleMenuItemKeyDown=e=>{if(e.defaultPrevented)return!1;switch(e.key){case de.g:case de.m:return this.invoke(),!1;case de.d:return this.expandAndFocus(),!1;case de.c:if(this.expanded)return this.expanded=!1,this.focus(),!1}return!0},this.handleMenuItemClick=e=>(e.defaultPrevented||this.disabled||this.invoke(),!1),this.submenuLoaded=()=>{this.focusSubmenuOnLoad&&(this.focusSubmenuOnLoad=!1,this.hasSubmenu&&(this.submenu.focus(),this.setAttribute("tabindex","-1")))},this.handleMouseOver=e=>(this.disabled||!this.hasSubmenu||this.expanded||(this.expanded=!0),!1),this.handleMouseOut=e=>(!this.expanded||this.contains(document.activeElement)||(this.expanded=!1),!1),this.expandAndFocus=()=>{this.hasSubmenu&&(this.focusSubmenuOnLoad=!0,this.expanded=!0)},this.invoke=()=>{if(!this.disabled)switch(this.role){case Jt:this.checked=!this.checked;break;case Yt:this.updateSubmenu(),this.hasSubmenu?this.expandAndFocus():this.$emit("change");break;case Xt:this.checked||(this.checked=!0)}},this.updateSubmenu=()=>{this.submenu=this.domChildren().find(e=>"menu"===e.getAttribute("role")),this.hasSubmenu=void 0!==this.submenu}}expandedChanged(e){if(this.$fastController.isConnected){if(void 0===this.submenu)return;!1===this.expanded?this.submenu.collapseExpandedItem():this.currentDirection=en(this),this.$emit("expanded-change",this,{bubbles:!1})}}checkedChanged(e,t){this.$fastController.isConnected&&this.$emit("change")}connectedCallback(){super.connectedCallback(),Ie.a.queueUpdate(()=>{this.updateSubmenu()}),this.startColumnCount||(this.startColumnCount=1),this.observer=new MutationObserver(this.updateSubmenu)}disconnectedCallback(){super.disconnectedCallback(),this.submenu=void 0,void 0!==this.observer&&(this.observer.disconnect(),this.observer=void 0)}domChildren(){return Array.from(this.children).filter(e=>!e.hasAttribute("hidden"))}}Object(oe.a)([Object(ce.c)({mode:"boolean"})],tn.prototype,"disabled",void 0),Object(oe.a)([Object(ce.c)({mode:"boolean"})],tn.prototype,"expanded",void 0),Object(oe.a)([se.d],tn.prototype,"startColumnCount",void 0),Object(oe.a)([ce.c],tn.prototype,"role",void 0),Object(oe.a)([Object(ce.c)({mode:"boolean"})],tn.prototype,"checked",void 0),Object(oe.a)([se.d],tn.prototype,"submenuRegion",void 0),Object(oe.a)([se.d],tn.prototype,"hasSubmenu",void 0),Object(oe.a)([se.d],tn.prototype,"currentDirection",void 0),Object(oe.a)([se.d],tn.prototype,"submenu",void 0),Object(_e.a)(tn,me.a);class nn extends ue.a{constructor(){super(...arguments),this.expandedItem=null,this.focusIndex=-1,this.isNestedMenu=()=>null!==this.parentElement&&Object(fe.c)(this.parentElement)&&"menuitem"===this.parentElement.getAttribute("role"),this.handleFocusOut=e=>{if(!this.contains(e.relatedTarget)&&void 0!==this.menuItems){this.collapseExpandedItem();const e=this.menuItems.findIndex(this.isFocusableElement);this.menuItems[this.focusIndex].setAttribute("tabindex","-1"),this.menuItems[e].setAttribute("tabindex","0"),this.focusIndex=e}},this.handleItemFocus=e=>{const t=e.target;void 0!==this.menuItems&&t!==this.menuItems[this.focusIndex]&&(this.menuItems[this.focusIndex].setAttribute("tabindex","-1"),this.focusIndex=this.menuItems.indexOf(t),t.setAttribute("tabindex","0"))},this.handleExpandedChanged=e=>{if(e.defaultPrevented||null===e.target||void 0===this.menuItems||this.menuItems.indexOf(e.target)<0)return;e.preventDefault();const t=e.target;null===this.expandedItem||t!==this.expandedItem||!1!==t.expanded?t.expanded&&(null!==this.expandedItem&&this.expandedItem!==t&&(this.expandedItem.expanded=!1),this.menuItems[this.focusIndex].setAttribute("tabindex","-1"),this.expandedItem=t,this.focusIndex=this.menuItems.indexOf(t),t.setAttribute("tabindex","0")):this.expandedItem=null},this.removeItemListeners=()=>{void 0!==this.menuItems&&this.menuItems.forEach(e=>{e.removeEventListener("expanded-change",this.handleExpandedChanged),e.removeEventListener("focus",this.handleItemFocus)})},this.setItems=()=>{const e=this.domChildren();this.removeItemListeners(),this.menuItems=e;const t=this.menuItems.filter(this.isMenuItemElement);t.length&&(this.focusIndex=0);const n=t.reduce((e,t)=>{const n=function(e){const t=e.getAttribute("role"),n=e.querySelector("[slot=start]");return t!==Yt&&null===n||t===Yt&&null!==n?1:t!==Yt&&null!==n?2:0}(t);return e>n?e:n},0);t.forEach((e,t)=>{e.setAttribute("tabindex",0===t?"0":"-1"),e.addEventListener("expanded-change",this.handleExpandedChanged),e.addEventListener("focus",this.handleItemFocus),e instanceof tn&&(e.startColumnCount=n)})},this.changeHandler=e=>{if(void 0===this.menuItems)return;const t=e.target,n=this.menuItems.indexOf(t);if(-1!==n&&"menuitemradio"===t.role&&!0===t.checked){for(let e=n-1;e>=0;--e){const t=this.menuItems[e],n=t.getAttribute("role");if(n===Xt&&(t.checked=!1),"separator"===n)break}const e=this.menuItems.length-1;for(let t=n+1;t<=e;++t){const e=this.menuItems[t],n=e.getAttribute("role");if(n===Xt&&(e.checked=!1),"separator"===n)break}}},this.isMenuItemElement=e=>Object(fe.c)(e)&&nn.focusableElementRoles.hasOwnProperty(e.getAttribute("role")),this.isFocusableElement=e=>this.isMenuItemElement(e)}itemsChanged(e,t){this.$fastController.isConnected&&void 0!==this.menuItems&&this.setItems()}connectedCallback(){super.connectedCallback(),Ie.a.queueUpdate(()=>{this.setItems()}),this.addEventListener("change",this.changeHandler)}disconnectedCallback(){super.disconnectedCallback(),this.removeItemListeners(),this.menuItems=void 0,this.removeEventListener("change",this.changeHandler)}focus(){this.setFocus(0,1)}collapseExpandedItem(){null!==this.expandedItem&&(this.expandedItem.expanded=!1,this.expandedItem=null)}handleMenuKeyDown(e){if(!e.defaultPrevented&&void 0!==this.menuItems)switch(e.key){case de.b:return void this.setFocus(this.focusIndex+1,1);case de.e:return void this.setFocus(this.focusIndex-1,-1);case de.f:return void this.setFocus(this.menuItems.length-1,-1);case de.j:return void this.setFocus(0,1);default:return!0}}domChildren(){return Array.from(this.children).filter(e=>!e.hasAttribute("hidden"))}setFocus(e,t){if(void 0!==this.menuItems)for(;e>=0&&e<this.menuItems.length;){const n=this.menuItems[e];if(this.isFocusableElement(n)){this.focusIndex>-1&&this.menuItems.length>=this.focusIndex-1&&this.menuItems[this.focusIndex].setAttribute("tabindex","-1"),this.focusIndex=e,n.setAttribute("tabindex","0"),n.focus();break}e+=t}}}nn.focusableElementRoles=Zt,Object(oe.a)([se.d],nn.prototype,"items",void 0);const an="horizontal";class rn extends ue.a{constructor(){super(...arguments),this.role="separator",this.orientation=an}}Object(oe.a)([ce.c],rn.prototype,"role",void 0),Object(oe.a)([ce.c],rn.prototype,"orientation",void 0);const on=class extends nn{connectedCallback(){super.connectedCallback(),Ee.v.setValueFor(this,Ee.gb)}}.compose({baseName:"menu",baseClass:nn,template:(e,t)=>Se.a`
    <template
        slot="${e=>e.slot?e.slot:e.isNestedMenu()?"submenu":void 0}"
        role="menu"
        @keydown="${(e,t)=>e.handleMenuKeyDown(t.event)}"
        @focusout="${(e,t)=>e.handleFocusOut(t.event)}"
    >
        <slot ${Object(De.a)("items")}></slot>
    </template>
`,styles:(e,t)=>Oe.a`
    ${Object(we.a)("block")} :host {
      background: ${Ee.gb};
      border: calc(${Ee.vb} * 1px) solid transparent;
      border-radius: calc(${Ee.D} * 1px);
      box-shadow: ${kt.c};
      padding: calc((${Ee.s} - ${Ee.vb}) * 1px) 0;
      max-width: 368px;
      min-width: 64px;
    }

    :host([slot='submenu']) {
      width: max-content;
      margin: 0 calc(${Ee.s} * 2px);
    }

    ::slotted(${e.tagFor(tn)}) {
      margin: 0 calc(${Ee.s} * 1px);
    }

    ::slotted(${e.tagFor(rn)}) {
      margin: calc(${Ee.s} * 1px) 0;
    }

    ::slotted(hr) {
      box-sizing: content-box;
      height: 0;
      margin: calc(${Ee.s} * 1px) 0;
      border: none;
      border-top: calc(${Ee.vb} * 1px) solid ${Ee.mb};
    }
  `.withBehaviors(Object(At.a)(Oe.a`
        :host([slot='submenu']) {
          background: ${Lt.a.Canvas};
          border-color: ${Lt.a.CanvasText};
        }
      `))}),sn="focus",cn="focusin",dn="focusout",ln="keydown",un="resize",fn="scroll";var pn=n("oZuh");class mn extends ue.a{constructor(){super(...arguments),this.anchor="",this.viewport="",this.horizontalPositioningMode="uncontrolled",this.horizontalDefaultPosition="unset",this.horizontalViewportLock=!1,this.horizontalInset=!1,this.horizontalScaling="content",this.verticalPositioningMode="uncontrolled",this.verticalDefaultPosition="unset",this.verticalViewportLock=!1,this.verticalInset=!1,this.verticalScaling="content",this.fixedPlacement=!1,this.autoUpdateMode="anchor",this.anchorElement=null,this.viewportElement=null,this.initialLayoutComplete=!1,this.resizeDetector=null,this.baseHorizontalOffset=0,this.baseVerticalOffset=0,this.pendingPositioningUpdate=!1,this.pendingReset=!1,this.currentDirection=$t.a.ltr,this.regionVisible=!1,this.forceUpdate=!1,this.updateThreshold=.5,this.update=()=>{this.pendingPositioningUpdate||this.requestPositionUpdates()},this.startObservers=()=>{this.stopObservers(),null!==this.anchorElement&&(this.requestPositionUpdates(),null!==this.resizeDetector&&(this.resizeDetector.observe(this.anchorElement),this.resizeDetector.observe(this)))},this.requestPositionUpdates=()=>{null===this.anchorElement||this.pendingPositioningUpdate||(mn.intersectionService.requestPosition(this,this.handleIntersection),mn.intersectionService.requestPosition(this.anchorElement,this.handleIntersection),null!==this.viewportElement&&mn.intersectionService.requestPosition(this.viewportElement,this.handleIntersection),this.pendingPositioningUpdate=!0)},this.stopObservers=()=>{this.pendingPositioningUpdate&&(this.pendingPositioningUpdate=!1,mn.intersectionService.cancelRequestPosition(this,this.handleIntersection),null!==this.anchorElement&&mn.intersectionService.cancelRequestPosition(this.anchorElement,this.handleIntersection),null!==this.viewportElement&&mn.intersectionService.cancelRequestPosition(this.viewportElement,this.handleIntersection)),null!==this.resizeDetector&&this.resizeDetector.disconnect()},this.getViewport=()=>"string"!=typeof this.viewport||""===this.viewport?document.documentElement:document.getElementById(this.viewport),this.getAnchor=()=>document.getElementById(this.anchor),this.handleIntersection=e=>{this.pendingPositioningUpdate&&(this.pendingPositioningUpdate=!1,this.applyIntersectionEntries(e)&&this.updateLayout())},this.applyIntersectionEntries=e=>{const t=e.find(e=>e.target===this),n=e.find(e=>e.target===this.anchorElement),a=e.find(e=>e.target===this.viewportElement);return void 0!==t&&void 0!==a&&void 0!==n&&!!(!this.regionVisible||this.forceUpdate||void 0===this.regionRect||void 0===this.anchorRect||void 0===this.viewportRect||this.isRectDifferent(this.anchorRect,n.boundingClientRect)||this.isRectDifferent(this.viewportRect,a.boundingClientRect)||this.isRectDifferent(this.regionRect,t.boundingClientRect))&&(this.regionRect=t.boundingClientRect,this.anchorRect=n.boundingClientRect,this.viewportElement===document.documentElement?this.viewportRect=new DOMRectReadOnly(a.boundingClientRect.x+document.documentElement.scrollLeft,a.boundingClientRect.y+document.documentElement.scrollTop,a.boundingClientRect.width,a.boundingClientRect.height):this.viewportRect=a.boundingClientRect,this.updateRegionOffset(),this.forceUpdate=!1,!0)},this.updateRegionOffset=()=>{this.anchorRect&&this.regionRect&&(this.baseHorizontalOffset=this.baseHorizontalOffset+(this.anchorRect.left-this.regionRect.left)+(this.translateX-this.baseHorizontalOffset),this.baseVerticalOffset=this.baseVerticalOffset+(this.anchorRect.top-this.regionRect.top)+(this.translateY-this.baseVerticalOffset))},this.isRectDifferent=(e,t)=>Math.abs(e.top-t.top)>this.updateThreshold||Math.abs(e.right-t.right)>this.updateThreshold||Math.abs(e.bottom-t.bottom)>this.updateThreshold||Math.abs(e.left-t.left)>this.updateThreshold,this.handleResize=e=>{this.update()},this.reset=()=>{this.pendingReset&&(this.pendingReset=!1,null===this.anchorElement&&(this.anchorElement=this.getAnchor()),null===this.viewportElement&&(this.viewportElement=this.getViewport()),this.currentDirection=en(this),this.startObservers())},this.updateLayout=()=>{let e,t;if("uncontrolled"!==this.horizontalPositioningMode){const e=this.getPositioningOptions(this.horizontalInset);if("center"===this.horizontalDefaultPosition)t="center";else if("unset"!==this.horizontalDefaultPosition){let e=this.horizontalDefaultPosition;if("start"===e||"end"===e){const t=en(this);if(t!==this.currentDirection)return this.currentDirection=t,void this.initialize();e=this.currentDirection===$t.a.ltr?"start"===e?"left":"right":"start"===e?"right":"left"}switch(e){case"left":t=this.horizontalInset?"insetStart":"start";break;case"right":t=this.horizontalInset?"insetEnd":"end"}}const n=void 0!==this.horizontalThreshold?this.horizontalThreshold:void 0!==this.regionRect?this.regionRect.width:0,a=void 0!==this.anchorRect?this.anchorRect.left:0,i=void 0!==this.anchorRect?this.anchorRect.right:0,r=void 0!==this.anchorRect?this.anchorRect.width:0,o=void 0!==this.viewportRect?this.viewportRect.left:0,s=void 0!==this.viewportRect?this.viewportRect.right:0;(void 0===t||"locktodefault"!==this.horizontalPositioningMode&&this.getAvailableSpace(t,a,i,r,o,s)<n)&&(t=this.getAvailableSpace(e[0],a,i,r,o,s)>this.getAvailableSpace(e[1],a,i,r,o,s)?e[0]:e[1])}if("uncontrolled"!==this.verticalPositioningMode){const t=this.getPositioningOptions(this.verticalInset);if("center"===this.verticalDefaultPosition)e="center";else if("unset"!==this.verticalDefaultPosition)switch(this.verticalDefaultPosition){case"top":e=this.verticalInset?"insetStart":"start";break;case"bottom":e=this.verticalInset?"insetEnd":"end"}const n=void 0!==this.verticalThreshold?this.verticalThreshold:void 0!==this.regionRect?this.regionRect.height:0,a=void 0!==this.anchorRect?this.anchorRect.top:0,i=void 0!==this.anchorRect?this.anchorRect.bottom:0,r=void 0!==this.anchorRect?this.anchorRect.height:0,o=void 0!==this.viewportRect?this.viewportRect.top:0,s=void 0!==this.viewportRect?this.viewportRect.bottom:0;(void 0===e||"locktodefault"!==this.verticalPositioningMode&&this.getAvailableSpace(e,a,i,r,o,s)<n)&&(e=this.getAvailableSpace(t[0],a,i,r,o,s)>this.getAvailableSpace(t[1],a,i,r,o,s)?t[0]:t[1])}const n=this.getNextRegionDimension(t,e),a=this.horizontalPosition!==t||this.verticalPosition!==e;if(this.setHorizontalPosition(t,n),this.setVerticalPosition(e,n),this.updateRegionStyle(),!this.initialLayoutComplete)return this.initialLayoutComplete=!0,void this.requestPositionUpdates();this.regionVisible||(this.regionVisible=!0,this.style.removeProperty("pointer-events"),this.style.removeProperty("opacity"),this.classList.toggle("loaded",!0),this.$emit("loaded",this,{bubbles:!1})),this.updatePositionClasses(),a&&this.$emit("positionchange",this,{bubbles:!1})},this.updateRegionStyle=()=>{this.style.width=this.regionWidth,this.style.height=this.regionHeight,this.style.transform=`translate(${this.translateX}px, ${this.translateY}px)`},this.updatePositionClasses=()=>{this.classList.toggle("top","start"===this.verticalPosition),this.classList.toggle("bottom","end"===this.verticalPosition),this.classList.toggle("inset-top","insetStart"===this.verticalPosition),this.classList.toggle("inset-bottom","insetEnd"===this.verticalPosition),this.classList.toggle("vertical-center","center"===this.verticalPosition),this.classList.toggle("left","start"===this.horizontalPosition),this.classList.toggle("right","end"===this.horizontalPosition),this.classList.toggle("inset-left","insetStart"===this.horizontalPosition),this.classList.toggle("inset-right","insetEnd"===this.horizontalPosition),this.classList.toggle("horizontal-center","center"===this.horizontalPosition)},this.setHorizontalPosition=(e,t)=>{if(void 0===e||void 0===this.regionRect||void 0===this.anchorRect||void 0===this.viewportRect)return;let n=0;switch(this.horizontalScaling){case"anchor":case"fill":n=this.horizontalViewportLock?this.viewportRect.width:t.width,this.regionWidth=`${n}px`;break;case"content":n=this.regionRect.width,this.regionWidth="unset"}let a=0;switch(e){case"start":this.translateX=this.baseHorizontalOffset-n,this.horizontalViewportLock&&this.anchorRect.left>this.viewportRect.right&&(this.translateX=this.translateX-(this.anchorRect.left-this.viewportRect.right));break;case"insetStart":this.translateX=this.baseHorizontalOffset-n+this.anchorRect.width,this.horizontalViewportLock&&this.anchorRect.right>this.viewportRect.right&&(this.translateX=this.translateX-(this.anchorRect.right-this.viewportRect.right));break;case"insetEnd":this.translateX=this.baseHorizontalOffset,this.horizontalViewportLock&&this.anchorRect.left<this.viewportRect.left&&(this.translateX=this.translateX-(this.anchorRect.left-this.viewportRect.left));break;case"end":this.translateX=this.baseHorizontalOffset+this.anchorRect.width,this.horizontalViewportLock&&this.anchorRect.right<this.viewportRect.left&&(this.translateX=this.translateX-(this.anchorRect.right-this.viewportRect.left));break;case"center":if(a=(this.anchorRect.width-n)/2,this.translateX=this.baseHorizontalOffset+a,this.horizontalViewportLock){const e=this.anchorRect.left+a,t=this.anchorRect.right-a;e<this.viewportRect.left&&!(t>this.viewportRect.right)?this.translateX=this.translateX-(e-this.viewportRect.left):t>this.viewportRect.right&&!(e<this.viewportRect.left)&&(this.translateX=this.translateX-(t-this.viewportRect.right))}}this.horizontalPosition=e},this.setVerticalPosition=(e,t)=>{if(void 0===e||void 0===this.regionRect||void 0===this.anchorRect||void 0===this.viewportRect)return;let n=0;switch(this.verticalScaling){case"anchor":case"fill":n=this.verticalViewportLock?this.viewportRect.height:t.height,this.regionHeight=`${n}px`;break;case"content":n=this.regionRect.height,this.regionHeight="unset"}let a=0;switch(e){case"start":this.translateY=this.baseVerticalOffset-n,this.verticalViewportLock&&this.anchorRect.top>this.viewportRect.bottom&&(this.translateY=this.translateY-(this.anchorRect.top-this.viewportRect.bottom));break;case"insetStart":this.translateY=this.baseVerticalOffset-n+this.anchorRect.height,this.verticalViewportLock&&this.anchorRect.bottom>this.viewportRect.bottom&&(this.translateY=this.translateY-(this.anchorRect.bottom-this.viewportRect.bottom));break;case"insetEnd":this.translateY=this.baseVerticalOffset,this.verticalViewportLock&&this.anchorRect.top<this.viewportRect.top&&(this.translateY=this.translateY-(this.anchorRect.top-this.viewportRect.top));break;case"end":this.translateY=this.baseVerticalOffset+this.anchorRect.height,this.verticalViewportLock&&this.anchorRect.bottom<this.viewportRect.top&&(this.translateY=this.translateY-(this.anchorRect.bottom-this.viewportRect.top));break;case"center":if(a=(this.anchorRect.height-n)/2,this.translateY=this.baseVerticalOffset+a,this.verticalViewportLock){const e=this.anchorRect.top+a,t=this.anchorRect.bottom-a;e<this.viewportRect.top&&!(t>this.viewportRect.bottom)?this.translateY=this.translateY-(e-this.viewportRect.top):t>this.viewportRect.bottom&&!(e<this.viewportRect.top)&&(this.translateY=this.translateY-(t-this.viewportRect.bottom))}}this.verticalPosition=e},this.getPositioningOptions=e=>e?["insetStart","insetEnd"]:["start","end"],this.getAvailableSpace=(e,t,n,a,i,r)=>{const o=t-i,s=r-(t+a);switch(e){case"start":return o;case"insetStart":return o+a;case"insetEnd":return s+a;case"end":return s;case"center":return 2*Math.min(o,s)+a}},this.getNextRegionDimension=(e,t)=>{const n={height:void 0!==this.regionRect?this.regionRect.height:0,width:void 0!==this.regionRect?this.regionRect.width:0};return void 0!==e&&"fill"===this.horizontalScaling?n.width=this.getAvailableSpace(e,void 0!==this.anchorRect?this.anchorRect.left:0,void 0!==this.anchorRect?this.anchorRect.right:0,void 0!==this.anchorRect?this.anchorRect.width:0,void 0!==this.viewportRect?this.viewportRect.left:0,void 0!==this.viewportRect?this.viewportRect.right:0):"anchor"===this.horizontalScaling&&(n.width=void 0!==this.anchorRect?this.anchorRect.width:0),void 0!==t&&"fill"===this.verticalScaling?n.height=this.getAvailableSpace(t,void 0!==this.anchorRect?this.anchorRect.top:0,void 0!==this.anchorRect?this.anchorRect.bottom:0,void 0!==this.anchorRect?this.anchorRect.height:0,void 0!==this.viewportRect?this.viewportRect.top:0,void 0!==this.viewportRect?this.viewportRect.bottom:0):"anchor"===this.verticalScaling&&(n.height=void 0!==this.anchorRect?this.anchorRect.height:0),n},this.startAutoUpdateEventListeners=()=>{window.addEventListener(un,this.update,{passive:!0}),window.addEventListener(fn,this.update,{passive:!0,capture:!0}),null!==this.resizeDetector&&null!==this.viewportElement&&this.resizeDetector.observe(this.viewportElement)},this.stopAutoUpdateEventListeners=()=>{window.removeEventListener(un,this.update),window.removeEventListener(fn,this.update),null!==this.resizeDetector&&null!==this.viewportElement&&this.resizeDetector.unobserve(this.viewportElement)}}anchorChanged(){this.initialLayoutComplete&&(this.anchorElement=this.getAnchor())}viewportChanged(){this.initialLayoutComplete&&(this.viewportElement=this.getViewport())}horizontalPositioningModeChanged(){this.requestReset()}horizontalDefaultPositionChanged(){this.updateForAttributeChange()}horizontalViewportLockChanged(){this.updateForAttributeChange()}horizontalInsetChanged(){this.updateForAttributeChange()}horizontalThresholdChanged(){this.updateForAttributeChange()}horizontalScalingChanged(){this.updateForAttributeChange()}verticalPositioningModeChanged(){this.requestReset()}verticalDefaultPositionChanged(){this.updateForAttributeChange()}verticalViewportLockChanged(){this.updateForAttributeChange()}verticalInsetChanged(){this.updateForAttributeChange()}verticalThresholdChanged(){this.updateForAttributeChange()}verticalScalingChanged(){this.updateForAttributeChange()}fixedPlacementChanged(){this.$fastController.isConnected&&this.initialLayoutComplete&&this.initialize()}autoUpdateModeChanged(e,t){this.$fastController.isConnected&&this.initialLayoutComplete&&("auto"===e&&this.stopAutoUpdateEventListeners(),"auto"===t&&this.startAutoUpdateEventListeners())}anchorElementChanged(){this.requestReset()}viewportElementChanged(){this.$fastController.isConnected&&this.initialLayoutComplete&&this.initialize()}connectedCallback(){super.connectedCallback(),"auto"===this.autoUpdateMode&&this.startAutoUpdateEventListeners(),this.initialize()}disconnectedCallback(){super.disconnectedCallback(),"auto"===this.autoUpdateMode&&this.stopAutoUpdateEventListeners(),this.stopObservers(),this.disconnectResizeDetector()}adoptedCallback(){this.initialize()}disconnectResizeDetector(){null!==this.resizeDetector&&(this.resizeDetector.disconnect(),this.resizeDetector=null)}initializeResizeDetector(){this.disconnectResizeDetector(),this.resizeDetector=new window.ResizeObserver(this.handleResize)}updateForAttributeChange(){this.$fastController.isConnected&&this.initialLayoutComplete&&(this.forceUpdate=!0,this.update())}initialize(){this.initializeResizeDetector(),null===this.anchorElement&&(this.anchorElement=this.getAnchor()),this.requestReset()}requestReset(){this.$fastController.isConnected&&!1===this.pendingReset&&(this.setInitialState(),Ie.a.queueUpdate(()=>this.reset()),this.pendingReset=!0)}setInitialState(){this.initialLayoutComplete=!1,this.regionVisible=!1,this.translateX=0,this.translateY=0,this.baseHorizontalOffset=0,this.baseVerticalOffset=0,this.viewportRect=void 0,this.regionRect=void 0,this.anchorRect=void 0,this.verticalPosition=void 0,this.horizontalPosition=void 0,this.style.opacity="0",this.style.pointerEvents="none",this.forceUpdate=!1,this.style.position=this.fixedPlacement?"fixed":"absolute",this.updatePositionClasses(),this.updateRegionStyle()}}mn.intersectionService=new class{constructor(){this.intersectionDetector=null,this.observedElements=new Map,this.requestPosition=(e,t)=>{var n;null!==this.intersectionDetector&&(this.observedElements.has(e)?null===(n=this.observedElements.get(e))||void 0===n||n.push(t):(this.observedElements.set(e,[t]),this.intersectionDetector.observe(e)))},this.cancelRequestPosition=(e,t)=>{const n=this.observedElements.get(e);if(void 0!==n){const e=n.indexOf(t);-1!==e&&n.splice(e,1)}},this.initializeIntersectionDetector=()=>{pn.a.IntersectionObserver&&(this.intersectionDetector=new IntersectionObserver(this.handleIntersection,{root:null,rootMargin:"0px",threshold:[0,1]}))},this.handleIntersection=e=>{if(null===this.intersectionDetector)return;const t=[],n=[];e.forEach(e=>{var a;null===(a=this.intersectionDetector)||void 0===a||a.unobserve(e.target);const i=this.observedElements.get(e.target);void 0!==i&&(i.forEach(a=>{let i=t.indexOf(a);-1===i&&(i=t.length,t.push(a),n.push([])),n[i].push(e)}),this.observedElements.delete(e.target))}),t.forEach((e,t)=>{e(n[t])})},this.initializeIntersectionDetector()}},Object(oe.a)([ce.c],mn.prototype,"anchor",void 0),Object(oe.a)([ce.c],mn.prototype,"viewport",void 0),Object(oe.a)([Object(ce.c)({attribute:"horizontal-positioning-mode"})],mn.prototype,"horizontalPositioningMode",void 0),Object(oe.a)([Object(ce.c)({attribute:"horizontal-default-position"})],mn.prototype,"horizontalDefaultPosition",void 0),Object(oe.a)([Object(ce.c)({attribute:"horizontal-viewport-lock",mode:"boolean"})],mn.prototype,"horizontalViewportLock",void 0),Object(oe.a)([Object(ce.c)({attribute:"horizontal-inset",mode:"boolean"})],mn.prototype,"horizontalInset",void 0),Object(oe.a)([Object(ce.c)({attribute:"horizontal-threshold"})],mn.prototype,"horizontalThreshold",void 0),Object(oe.a)([Object(ce.c)({attribute:"horizontal-scaling"})],mn.prototype,"horizontalScaling",void 0),Object(oe.a)([Object(ce.c)({attribute:"vertical-positioning-mode"})],mn.prototype,"verticalPositioningMode",void 0),Object(oe.a)([Object(ce.c)({attribute:"vertical-default-position"})],mn.prototype,"verticalDefaultPosition",void 0),Object(oe.a)([Object(ce.c)({attribute:"vertical-viewport-lock",mode:"boolean"})],mn.prototype,"verticalViewportLock",void 0),Object(oe.a)([Object(ce.c)({attribute:"vertical-inset",mode:"boolean"})],mn.prototype,"verticalInset",void 0),Object(oe.a)([Object(ce.c)({attribute:"vertical-threshold"})],mn.prototype,"verticalThreshold",void 0),Object(oe.a)([Object(ce.c)({attribute:"vertical-scaling"})],mn.prototype,"verticalScaling",void 0),Object(oe.a)([Object(ce.c)({attribute:"fixed-placement",mode:"boolean"})],mn.prototype,"fixedPlacement",void 0),Object(oe.a)([Object(ce.c)({attribute:"auto-update-mode"})],mn.prototype,"autoUpdateMode",void 0),Object(oe.a)([se.d],mn.prototype,"anchorElement",void 0),Object(oe.a)([se.d],mn.prototype,"viewportElement",void 0),Object(oe.a)([se.d],mn.prototype,"initialLayoutComplete",void 0);const _n=tn.compose({baseName:"menu-item",template:(e,t)=>Se.a`
    <template
        role="${e=>e.role}"
        aria-haspopup="${e=>e.hasSubmenu?"menu":void 0}"
        aria-checked="${e=>e.role!==Yt?e.checked:void 0}"
        aria-disabled="${e=>e.disabled}"
        aria-expanded="${e=>e.expanded}"
        @keydown="${(e,t)=>e.handleMenuItemKeyDown(t.event)}"
        @click="${(e,t)=>e.handleMenuItemClick(t.event)}"
        @mouseover="${(e,t)=>e.handleMouseOver(t.event)}"
        @mouseout="${(e,t)=>e.handleMouseOut(t.event)}"
        class="${e=>e.disabled?"disabled":""} ${e=>e.expanded?"expanded":""} ${e=>`indent-${e.startColumnCount}`}"
    >
            ${Object(Ct.a)(e=>e.role===Jt,Se.a`
                    <div part="input-container" class="input-container">
                        <span part="checkbox" class="checkbox">
                            <slot name="checkbox-indicator">
                                ${t.checkboxIndicator||""}
                            </slot>
                        </span>
                    </div>
                `)}
            ${Object(Ct.a)(e=>e.role===Xt,Se.a`
                    <div part="input-container" class="input-container">
                        <span part="radio" class="radio">
                            <slot name="radio-indicator">
                                ${t.radioIndicator||""}
                            </slot>
                        </span>
                    </div>
                `)}
        </div>
        ${Object(me.d)(e,t)}
        <span class="content" part="content">
            <slot></slot>
        </span>
        ${Object(me.b)(e,t)}
        ${Object(Ct.a)(e=>e.hasSubmenu,Se.a`
                <div
                    part="expand-collapse-glyph-container"
                    class="expand-collapse-glyph-container"
                >
                    <span part="expand-collapse" class="expand-collapse">
                        <slot name="expand-collapse-indicator">
                            ${t.expandCollapseGlyph||""}
                        </slot>
                    </span>
                </div>
            `)}
        ${Object(Ct.a)(e=>e.expanded,Se.a`
                <${e.tagFor(mn)}
                    :anchorElement="${e=>e}"
                    vertical-positioning-mode="dynamic"
                    vertical-default-position="bottom"
                    vertical-inset="true"
                    horizontal-positioning-mode="dynamic"
                    horizontal-default-position="end"
                    class="submenu-region"
                    dir="${e=>e.currentDirection}"
                    @loaded="${e=>e.submenuLoaded()}"
                    ${Object(Ot.a)("submenuRegion")}
                    part="submenu-region"
                >
                    <slot name="submenu"></slot>
                </${e.tagFor(mn)}>
            `)}
    </template>
`,styles:(e,t)=>Oe.a`
    ${Object(we.a)("grid")} :host {
      contain: layout;
      overflow: visible;
      ${Tt.a}
      box-sizing: border-box;
      height: calc(${Mt.a} * 1px);
      grid-template-columns: minmax(32px, auto) 1fr minmax(32px, auto);
      grid-template-rows: auto;
      justify-items: center;
      align-items: center;
      padding: 0;
      white-space: nowrap;
      color: ${Ee.fb};
      fill: currentcolor;
      cursor: pointer;
      border-radius: calc(${Ee.q} * 1px);
      border: calc(${Ee.vb} * 1px) solid transparent;
      position: relative;
    }

    :host(.indent-0) {
      grid-template-columns: auto 1fr minmax(32px, auto);
    }

    :host(.indent-0) .content {
      grid-column: 1;
      grid-row: 1;
      margin-inline-start: 10px;
    }

    :host(.indent-0) .expand-collapse-glyph-container {
      grid-column: 5;
      grid-row: 1;
    }

    :host(.indent-2) {
      grid-template-columns: minmax(32px, auto) minmax(32px, auto) 1fr minmax(32px, auto) minmax(32px, auto);
    }

    :host(.indent-2) .content {
      grid-column: 3;
      grid-row: 1;
      margin-inline-start: 10px;
    }

    :host(.indent-2) .expand-collapse-glyph-container {
      grid-column: 5;
      grid-row: 1;
    }

    :host(.indent-2) .start {
      grid-column: 2;
    }

    :host(.indent-2) .end {
      grid-column: 4;
    }

    :host(:${wt.a}) {
      ${Ae.a}
    }

    :host(:not([disabled]):hover) {
      background: ${Ee.Y};
    }

    :host(:not([disabled]):active),
    :host(.expanded) {
      background: ${Ee.W};
      color: ${Ee.fb};
      z-index: 2;
    }

    :host([disabled]) {
      cursor: ${Et.a};
      opacity: ${Ee.u};
    }

    .content {
      grid-column-start: 2;
      justify-self: start;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .start,
    .end {
      display: flex;
      justify-content: center;
    }

    :host(.indent-0[aria-haspopup='menu']) {
      display: grid;
      grid-template-columns: minmax(32px, auto) auto 1fr minmax(32px, auto) minmax(32px, auto);
      align-items: center;
      min-height: 32px;
    }

    :host(.indent-1[aria-haspopup='menu']),
    :host(.indent-1[role='menuitemcheckbox']),
    :host(.indent-1[role='menuitemradio']) {
      display: grid;
      grid-template-columns: minmax(32px, auto) auto 1fr minmax(32px, auto) minmax(32px, auto);
      align-items: center;
      min-height: 32px;
    }

    :host(.indent-2:not([aria-haspopup='menu'])) .end {
      grid-column: 5;
    }

    :host .input-container,
    :host .expand-collapse-glyph-container {
      display: none;
    }

    :host([aria-haspopup='menu']) .expand-collapse-glyph-container,
    :host([role='menuitemcheckbox']) .input-container,
    :host([role='menuitemradio']) .input-container {
      display: grid;
    }

    :host([aria-haspopup='menu']) .content,
    :host([role='menuitemcheckbox']) .content,
    :host([role='menuitemradio']) .content {
      grid-column-start: 3;
    }

    :host([aria-haspopup='menu'].indent-0) .content {
      grid-column-start: 1;
    }

    :host([aria-haspopup='menu']) .end,
    :host([role='menuitemcheckbox']) .end,
    :host([role='menuitemradio']) .end {
      grid-column-start: 4;
    }

    :host .expand-collapse,
    :host .checkbox,
    :host .radio {
      display: flex;
      align-items: center;
      justify-content: center;
      position: relative;
      box-sizing: border-box;
    }

    :host .checkbox-indicator,
    :host .radio-indicator,
    slot[name='checkbox-indicator'],
    slot[name='radio-indicator'] {
      display: none;
    }

    ::slotted([slot='end']:not(svg)) {
      margin-inline-end: 10px;
      color: ${Ee.cb};
    }

    :host([aria-checked='true']) .checkbox-indicator,
    :host([aria-checked='true']) slot[name='checkbox-indicator'],
    :host([aria-checked='true']) .radio-indicator,
    :host([aria-checked='true']) slot[name='radio-indicator'] {
      display: flex;
    }
  `.withBehaviors(Object(At.a)(Oe.a`
        :host,
        ::slotted([slot='end']:not(svg)) {
          forced-color-adjust: none;
          color: ${Lt.a.ButtonText};
          fill: currentcolor;
        }
        :host(:not([disabled]):hover) {
          background: ${Lt.a.Highlight};
          color: ${Lt.a.HighlightText};
          fill: currentcolor;
        }
        :host(:hover) .start,
        :host(:hover) .end,
        :host(:hover)::slotted(svg),
        :host(:active) .start,
        :host(:active) .end,
        :host(:active)::slotted(svg),
        :host(:hover) ::slotted([slot='end']:not(svg)),
        :host(:${wt.a}) ::slotted([slot='end']:not(svg)) {
          color: ${Lt.a.HighlightText};
          fill: currentcolor;
        }
        :host(.expanded) {
          background: ${Lt.a.Highlight};
          color: ${Lt.a.HighlightText};
        }
        :host(:${wt.a}) {
          background: ${Lt.a.Highlight};
          outline-color: ${Lt.a.ButtonText};
          color: ${Lt.a.HighlightText};
          fill: currentcolor;
        }
        :host([disabled]),
        :host([disabled]:hover),
        :host([disabled]:hover) .start,
        :host([disabled]:hover) .end,
        :host([disabled]:hover)::slotted(svg),
        :host([disabled]:${wt.a}) {
          background: ${Lt.a.ButtonFace};
          color: ${Lt.a.GrayText};
          fill: currentcolor;
          opacity: 1;
        }
        :host([disabled]:${wt.a}) {
          outline-color: ${Lt.a.GrayText};
        }
        :host .expanded-toggle,
        :host .checkbox,
        :host .radio {
          border-color: ${Lt.a.ButtonText};
          background: ${Lt.a.HighlightText};
        }
        :host([checked]) .checkbox,
        :host([checked]) .radio {
          background: ${Lt.a.HighlightText};
          border-color: ${Lt.a.HighlightText};
        }
        :host(:hover) .expanded-toggle,
            :host(:hover) .checkbox,
            :host(:hover) .radio,
            :host(:${wt.a}) .expanded-toggle,
            :host(:${wt.a}) .checkbox,
            :host(:${wt.a}) .radio,
            :host([checked]:hover) .checkbox,
            :host([checked]:hover) .radio,
            :host([checked]:${wt.a}) .checkbox,
            :host([checked]:${wt.a}) .radio {
          border-color: ${Lt.a.HighlightText};
        }
        :host([aria-checked='true']) {
          background: ${Lt.a.Highlight};
          color: ${Lt.a.HighlightText};
        }
        :host([aria-checked='true']) .checkbox-indicator,
        :host([aria-checked='true']) ::slotted([slot='checkbox-indicator']),
        :host([aria-checked='true']) ::slotted([slot='radio-indicator']) {
          fill: ${Lt.a.Highlight};
        }
        :host([aria-checked='true']) .radio-indicator {
          background: ${Lt.a.Highlight};
        }
      `),new zt(Oe.a`
        .expand-collapse-glyph-container {
          transform: rotate(0deg);
        }
      `,Oe.a`
        .expand-collapse-glyph-container {
          transform: rotate(180deg);
        }
      `)),checkboxIndicator:'\n    <svg width="16" height="16" xmlns="http://www.w3.org/2000/svg">\n      <path d="M13.86 3.66a.5.5 0 01-.02.7l-7.93 7.48a.6.6 0 01-.84-.02L2.4 9.1a.5.5 0 01.72-.7l2.4 2.44 7.65-7.2a.5.5 0 01.7.02z"/>\n    </svg>\n  ',expandCollapseGlyph:'\n    <svg width="16" height="16" xmlns="http://www.w3.org/2000/svg">\n      <path d="M5.65 3.15a.5.5 0 000 .7L9.79 8l-4.14 4.15a.5.5 0 00.7.7l4.5-4.5a.5.5 0 000-.7l-4.5-4.5a.5.5 0 00-.7 0z"/>\n    </svg>\n  ',radioIndicator:'\n    <svg width="16" height="16" xmlns="http://www.w3.org/2000/svg">\n      <circle cx="8" cy="8" r="2"/>\n    </svg>\n  '});var hn=n("dblZ");const bn={dotOptionsTitle:"More options"},gn=[u.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{display:flex;flex-direction:column;align-items:flex-end}:host .dot-icon{font-family:FabricMDL2Icons;min-width:40px;min-height:30px;text-align:center;line-height:30px}:host .menu{position:absolute;box-shadow:var(--neutral-fill-rest) 0 0 40px 5px;background:var(--neutral-fill-rest);z-index:1;display:none;white-space:nowrap;transform:var(--dot-options-translatey,translateY(32px))}:host .menu.open{display:block}:host fluent-button::part(control){background:inherit}
`];var vn=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},yn=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};class Sn extends hn.b{constructor(){super(...arguments),this._clickHandler=()=>this.open=!1,this.handleItemClick=(e,t)=>{e.preventDefault(),e.stopPropagation(),t(e),this.open=!1},this.handleItemKeydown=(e,t)=>{this.handleKeydownMenuOption(e),t(e),this.open=!1},this.onDotClick=e=>{e.preventDefault(),e.stopPropagation(),this.open=!this.open},this.onDotKeydown=e=>{"Enter"===e.key&&(e.preventDefault(),e.stopPropagation(),this.open=!this.open)}}static get styles(){return gn}get strings(){return bn}connectedCallback(){super.connectedCallback(),window.addEventListener("click",this._clickHandler)}disconnectedCallback(){window.removeEventListener("click",this._clickHandler),super.disconnectedCallback()}render(){const e=Object.keys(this.options);return u.c`
      <fluent-button
        appearance="stealth"
        aria-label=${this.strings.dotOptionsTitle}
        @click=${this.onDotClick}
        @keydown=${this.onDotKeydown}
        class="dot-icon">\uE712</fluent-button>
      <fluent-menu class=${Object(F.a)({menu:!0,open:this.open})}>
        ${e.map(e=>this.getMenuOption(e,this.options[e]))}
      </fluent-menu>`}getMenuOption(e,t){return u.c`
      <div
        class="dot-item"
        @click="${e=>{e.preventDefault(),e.stopPropagation(),t(e),this.open=!1}}"
        @keydown="${e=>{"Enter"!==e.code&&"Space"!==e.code||(e.preventDefault(),e.stopPropagation(),t(e),this.open=!1)}}"
      >
        <span class="dot-item-name">
          ${e}
      </fluent-menu-item>`}handleKeydownMenuOption(e){"Enter"===e.key&&(e.preventDefault(),e.stopPropagation())}}vn([Object(f.b)({type:Boolean}),yn("design:type",Boolean)],Sn.prototype,"open",void 0),vn([Object(f.b)({type:Object}),yn("design:type",Object)],Sn.prototype,"options",void 0);const Dn=[u.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{position:relative}:host .header::part(control){font-size:var(--arrow-options-button-font-size,large);font-weight:var(--arrow-options-button-font-weight,600);color:var(--arrow-options-button-font-color,var(--accent-base-color));background:var(--arrow-options-button-background-color,transparent)}:host .header::part(control):hover{background:var(--neutral-fill-stealth-hover)}:host .header::part(control):active,:host .header::part(control):focus{background:var(--neutral-fill-stealth-active)}:host .menu{position:absolute;left:var(--arrow-options-left,0);z-index:1;display:none}:host .menu.open{display:block;width:max-content}
`];var In=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},xn=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};class Cn extends hn.b{static get styles(){return Dn}constructor(){super(),this.onHeaderClick=e=>{Object.keys(this.options).length>1&&(e.preventDefault(),e.stopPropagation(),this.open=!this.open)},this.onHeaderKeyDown=e=>{if("Enter"===e.key){e.preventDefault(),e.stopPropagation(),this.open=!this.open;const t=this.renderRoot.querySelector("fluent-menu");t&&(t.classList.remove("closed"),t.classList.add("open"));const n=e.target;if(n){const e=this.renderRoot.querySelector("fluent-menu-item[tabindex='0']");e&&(n.blur(),e.focus())}}},this.value="",this.options={},this._clickHandler=()=>this.open=!1,window.addEventListener("onblur",()=>this.open=!1)}connectedCallback(){super.connectedCallback(),window.addEventListener("click",this._clickHandler)}disconnectedCallback(){window.removeEventListener("click",this._clickHandler),super.disconnectedCallback()}render(){return u.c`
      <fluent-button
        class="header"
        @click=${this.onHeaderClick}
        @keydown=${this.onHeaderKeyDown}
        appearance="lightweight">
          ${this.value}
      </fluent-button>
      <fluent-menu
        class=${Object(F.a)({menu:!0,open:this.open,closed:!this.open})}>
          ${this.getMenuOptions()}
      </fluent-menu>`}getMenuOptions(){return Object.keys(this.options).map(e=>u.c`
          <fluent-menu-item
            @click=${t=>{this.open=!1,this.options[e](t)}}
            @keydown=${t=>{const n=this.renderRoot.querySelector(".header");"Enter"===t.key?(this.open=!1,this.options[e](t),n.focus()):"Tab"===t.key?this.open=!1:"Escape"===t.key&&(this.open=!1,n&&n.focus())}}>
              ${e}
          </fluent-menu-item>`)}}In([Object(f.b)({type:Boolean}),xn("design:type",Boolean)],Cn.prototype,"open",void 0),In([Object(f.b)({type:String}),xn("design:type",String)],Cn.prototype,"value",void 0),In([Object(f.b)({type:Object}),xn("design:type",Object)],Cn.prototype,"options",void 0);var On,wn=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},En=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)},An=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};!function(e){e[e.planner=0]="planner",e[e.todo=1]="todo"}(On||(On={}));const Ln={BASE_SELF_ASSIGNED:"All Tasks",BUCKETS_SELF_ASSIGNED:"All Tasks",BUCKET_NOT_FOUND:"Folder not found",PLANS_SELF_ASSIGNED:"All groups",PLAN_NOT_FOUND:"Group not found"},kn={BASE_SELF_ASSIGNED:"Assigned to Me",BUCKETS_SELF_ASSIGNED:"All Tasks",BUCKET_NOT_FOUND:"Bucket not found",PLANS_SELF_ASSIGNED:"All Plans",PLAN_NOT_FOUND:"Plan not found"},Mn={"@odata.type":"#microsoft.graph.plannerAssignment",orderHint:" !"},Pn=()=>{Object(T.a)(Vt,Kt,qe.a,Me.a,Wt.a,Qt),Object(T.a)(on,_n,Me.a),Object(A.b)("arrow-options",Cn),Object(T.a)(on,_n,Me.a),Object(A.b)("dot-options",Sn),Object(te.a)(),M(),et(),Object(A.b)("tasks",Tn)};class Tn extends p.a{get res(){switch(this.dataSource){case On.todo:return Ln;case On.planner:default:return kn}}static get styles(){return nt}get strings(){return at}get isNewTaskVisible(){return this._isNewTaskVisible}set isNewTaskVisible(e){this._isNewTaskVisible=e,e||(this._newTaskDueDate=null,this._newTaskName="",this._newTaskGroupId="",this._newTaskContainerId="")}static get requiredScopes(){return[...new Set(["group.read.all","group.readwrite.all","tasks.read","tasks.readwrite",...P.requiredScopes,...tt.requiredScopes])]}constructor(){super(),this.dataSource=On.planner,this._isDarkMode=!1,this._me=null,this.onResize=()=>{this.mediaQuery!==this.previousMediaQuery&&(this.previousMediaQuery=this.mediaQuery,this.requestUpdate())},this.onThemeChanged=()=>{this._isDarkMode=ht(this)},this.onAddTaskClick=()=>{var e;const t=this.getPeoplePicker(null),n={};if(t)for(const a of null!==(e=null==t?void 0:t.selectedPeople)&&void 0!==e?e:[])n[a.id]=Mn;!this._newTaskBeingAdded&&this._newTaskName&&(this._currentGroup||this._newTaskGroupId)&&this.addTask(this._newTaskName,this._newTaskDueDate,this._currentGroup?this._currentGroup:this._newTaskGroupId,this._currentFolder?this._currentFolder:this._newTaskFolderId,n)},this.onAddTaskKeyDown=e=>{"Enter"!==e.key&&" "!==e.key||this.onAddTaskClick()},this.newTaskButtonKeydown=e=>{"Enter"===e.key&&(this.isNewTaskVisible=!this.isNewTaskVisible)},this.addNewTaskButtonClick=()=>{this.isNewTaskVisible=!this.isNewTaskVisible},this.handleNewTaskDateChange=e=>{const t=e.target.value;this._newTaskDueDate=t?new Date(t+"T17:00"):null},this.handleSelectedPlan=e=>{var t,n;if(this._newTaskGroupId=e.target.value,this.dataSource===On.planner){const e=this._groups.filter(e=>e.id===this._newTaskGroupId);this._newTaskContainerId=null!==(n=null===(t=e.pop())||void 0===t?void 0:t.containerId)&&void 0!==n?n:this._newTaskContainerId}},this.newTaskVisible=e=>{"Enter"===e.key&&(this.isNewTaskVisible=!1)},this.handleDateChange=e=>{const t=e.target.value;this._newTaskDueDate=t?new Date(t+"T17:00"):null},this.renderPlannerIcon=e=>Object(S.b)(S.a.Planner,e),this.renderBucketIcon=e=>Object(S.b)(S.a.Milestone,e),this.handlePeopleClick=(e,t)=>{this.togglePeoplePicker(t),e.stopPropagation()},this.handlePeopleKeydown=(e,t)=>{"Enter"!==e.key&&" "!==e.key||(this.togglePeoplePicker(t),e.stopPropagation(),e.preventDefault())},this.handlePeoplePickerKeydown=e=>{"Enter"===e.key&&e.stopPropagation()},this.clearState(),this.previousMediaQuery=this.mediaQuery}connectedCallback(){super.connectedCallback(),window.addEventListener("resize",this.onResize),window.addEventListener("darkmodechanged",this.onThemeChanged),this.onThemeChanged()}disconnectedCallback(){window.removeEventListener("resize",this.onResize),window.removeEventListener("darkmodechanged",this.onThemeChanged),super.disconnectedCallback()}attributeChangedCallback(e,t,n){super.attributeChangedCallback(e,t,n),"data-source"===e&&(this.dataSource===On.planner?(this._currentGroup=this.initialId,this._currentFolder=this.initialBucketId):this.dataSource===On.todo&&(this._currentGroup=null,this._currentFolder=this.initialId),this.clearState(),this.requestStateUpdate())}clearState(){this._newTaskFolderId="",this._newTaskGroupId="",this._newTaskDueDate=null,this._newTaskName="",this._newTaskBeingAdded=!1,this._tasks=[],this._folders=[],this._groups=[],this._hiddenTasks=[],this._loadingTasks=[],this._hasDoneInitialLoad=!1,this._inTaskLoad=!1,this._todoDefaultSet=!1}firstUpdated(e){super.firstUpdated(e),this.initialId&&!this._currentGroup&&(this.dataSource===On.planner?this._currentGroup=this.initialId:this.dataSource===On.todo&&(this._currentFolder=this.initialId)),this.dataSource===On.planner&&this.initialBucketId&&!this._currentFolder&&(this._currentFolder=this.initialBucketId)}render(){const e=this._inTaskLoad&&!this._hasDoneInitialLoad?this.renderLoadingTask():null;let t;return this.hideHeader||(t=u.c`
        <div class="header">
          ${this.renderPlanOptions()}
        </div>
      `),u.c`
      ${t}
      <div class="tasks">
        ${this._isNewTaskVisible?this.renderNewTask():null} ${e}
        ${Object(D.a)(this._tasks,e=>e.id,e=>this.renderTask(e))}
      </div>
    `}loadState(){return An(this,void 0,void 0,function*(){const e=this.getTaskSource();if(!e)return;const t=_.a.globalProvider;if(t&&t.state===i.c.SignedIn){if(this._inTaskLoad=!0,!this._me){const e=t.graph.forComponent(this);this._me=yield Object(C.e)(e)}this.groupId&&this.dataSource===On.planner?yield this._loadTasksForGroup(e):this.targetId?this.dataSource===On.todo?yield this._loadTargetTodoTasks(e):yield this._loadTargetPlannerTasks(e):yield this._loadAllTasks(e),this._tasks=this._tasks.filter(e=>this.isTaskInSelectedGroupFilter(e)).filter(e=>this.isTaskInSelectedFolderFilter(e)).filter(e=>!this._hiddenTasks.includes(e.id)),this.taskFilter&&(this._tasks=this._tasks.filter(e=>this.taskFilter(e._raw))),this._inTaskLoad=!1,this._hasDoneInitialLoad=!0}})}_loadTargetTodoTasks(e){return An(this,void 0,void 0,function*(){const t=yield e.getTaskGroups(),n=(yield Promise.all(t.map(t=>e.getTaskFoldersForTaskGroup(t.id)))).reduce((e,t)=>[...e,...t],[]),a=(yield Promise.all(n.map(t=>e.getTasksForTaskFolder(t.id,t.parentId)))).reduce((e,t)=>[...e,...t],[]);this._tasks=a,this._folders=n,this._groups=t,this._currentGroup=null})}_loadTargetPlannerTasks(e){return An(this,void 0,void 0,function*(){const t=yield e.getTaskGroup(this.targetId);let n=yield e.getTaskFoldersForTaskGroup(t.id);this.targetBucketId&&(n=n.filter(e=>e.id===this.targetBucketId));const a=(yield Promise.all(n.map(t=>e.getTasksForTaskFolder(t.id,t.parentId)))).reduce((e,t)=>[...e,...t],[]);this._tasks=a,this._folders=n,this._groups=[t]})}_loadAllTasks(e){return An(this,void 0,void 0,function*(){const t=yield e.getTaskGroups(),n=(yield Promise.all(t.map(t=>e.getTaskFoldersForTaskGroup(t.id)))).reduce((e,t)=>[...e,...t],[]);if(!this.initialId&&this.dataSource===On.todo&&!this._todoDefaultSet){this._todoDefaultSet=!0;const e=n.find(e=>e._raw.isDefaultFolder);e&&(this._currentFolder=e.id)}const a=(yield Promise.all(n.map(t=>e.getTasksForTaskFolder(t.id,t.parentId)))).reduce((e,t)=>[...e,...t],[]);this._tasks=a,this._folders=n,this._groups=t})}_loadTasksForGroup(e){return An(this,void 0,void 0,function*(){const t=yield e.getTaskGroupsForGroup(this.groupId),n=(yield Promise.all(t.map(t=>e.getTaskFoldersForTaskGroup(t.id)))).reduce((e,t)=>[...e,...t],[]),a=(yield Promise.all(n.map(t=>e.getTasksForTaskFolder(t.id,t.parentId)))).reduce((e,t)=>[...e,...t],[]);this._tasks=a,this._folders=n,this._groups=t})}addTask(e,t,n,a,i={}){return An(this,void 0,void 0,function*(){const r=this.getTaskSource();if(!r)return;const o={assignments:i,dueDate:t,immediateParentId:a,name:e,topParentId:n};this._newTaskBeingAdded=!0,o._raw=yield r.addTask(o),this.fireCustomEvent("taskAdded",o),yield this.requestStateUpdate(),this._newTaskBeingAdded=!1,this.isNewTaskVisible=!1})}completeTask(e){return An(this,void 0,void 0,function*(){const t=this.getTaskSource();t&&(this._loadingTasks=[...this._loadingTasks,e.id],yield t.setTaskComplete(e),this.fireCustomEvent("taskChanged",e),yield this.requestStateUpdate(),this._loadingTasks=this._loadingTasks.filter(t=>t!==e.id))})}uncompleteTask(e){return An(this,void 0,void 0,function*(){const t=this.getTaskSource();t&&(this._loadingTasks=[...this._loadingTasks,e.id],yield t.setTaskIncomplete(e),this.fireCustomEvent("taskChanged",e),yield this.requestStateUpdate(),this._loadingTasks=this._loadingTasks.filter(t=>t!==e.id))})}removeTask(e){return An(this,void 0,void 0,function*(){const t=this.getTaskSource();t&&(this._hiddenTasks=[...this._hiddenTasks,e.id],yield t.removeTask(e),this.fireCustomEvent("taskRemoved",e),yield this.requestStateUpdate(),this._hiddenTasks=this._hiddenTasks.filter(t=>t!==e.id))})}assignPeople(e,t=[]){return An(this,void 0,void 0,function*(){const n=this.getTaskSource();if(!n)return;let a=[];e&&e.assignments&&(a=Object.keys(e.assignments).sort());const i=t.map(e=>e.id);if(i.length===a.length&&i.sort().every((e,t)=>e===a[t]))return;const r={};for(const e of a)i.includes(e)?r[e]=Mn:r[e]=null;i.forEach(e=>{a.includes(e)||(r[e]=Mn)}),e&&(yield n.assignPeopleToTask(e,r),yield this.requestStateUpdate(),this._loadingTasks=this._loadingTasks.filter(t=>t!==e.id))})}renderPlanOptions(){var e;const t=_.a.globalProvider;if(!t||t.state!==i.c.SignedIn)return null;if(this._inTaskLoad&&!this._hasDoneInitialLoad)return u.c`<span class="loading-header"></span>`;const n=this.readOnly||this._isNewTaskVisible?null:u.c`
          <fluent-button
            appearance="accent"
            class="new-task-button"
            @keydown=${this.newTaskButtonKeydown}
            @click=${()=>this.isNewTaskVisible=!this.isNewTaskVisible}>
              <span slot="start">${Object(S.b)(S.a.Add,"currentColor")}</span>
              ${this.strings.addTaskButtonSubtitle}
          </fluent-button>
        `;if(this.dataSource===On.planner){const t=this._groups.find(e=>e.id===this._currentGroup)||{title:this.res.BASE_SELF_ASSIGNED},a={[this.res.BASE_SELF_ASSIGNED]:()=>{this._currentGroup=null,this._currentFolder=null}};for(const e of this._groups)a[e.title]=()=>{this._currentGroup=e.id,this._currentFolder=null};const i=m.a`
        <mgt-arrow-options
          class="arrow-options"
          .options="${a}"
          .value="${t.title}"></mgt-arrow-options>`,r=this._currentGroup?Object(S.b)(S.a.ChevronRight):null,o=this._folders.find(e=>e.id===this._currentFolder)||{name:this.res.BUCKETS_SELF_ASSIGNED},s={[this.res.BUCKETS_SELF_ASSIGNED]:()=>{this._currentFolder=null}};for(const e of this._folders.filter(e=>e.parentId===this._currentGroup))s[e.name]=()=>{this._currentFolder=e.id};const c=this.targetBucketId?u.c`
            <span class="plan-title">
              ${(null===(e=this._folders[0])||void 0===e?void 0:e.name)||""}
            </span>`:m.a`
            <mgt-arrow-options class="arrow-options" .options="${s}" .value="${o.name}"></mgt-arrow-options>
          `;return u.c`
        <div class="Title">
          ${i} ${r} ${this._currentGroup?c:null}
        </div>
        ${n}
      `}{const e=this._folders.find(e=>e.id===this.targetId)||{name:this.res.BUCKETS_SELF_ASSIGNED},t=this._folders.find(e=>e.id===this._currentFolder)||{name:this.res.BUCKETS_SELF_ASSIGNED},a={};for(const e of this._folders)a[e.name]=()=>{this._currentFolder=e.id};a[this.res.BUCKETS_SELF_ASSIGNED]=()=>{this._currentFolder=null};const i=this.targetId?u.c`
            <span class="plan-title">
              ${e.name}
            </span>
          `:m.a`
            <mgt-arrow-options class="arrow-options" .value="${t.name}" .options="${a}"></mgt-arrow-options>
          `;return u.c`
        <span class="title">
          ${i}
        </span>
        ${n}
      `}}renderNewTask(){const e="var(--neutral-foreground-hint)",t=u.c`
      <fluent-text-field
        autocomplete="off"
        placeholder=${this.strings.newTaskPlaceholder}
        .value="${this._newTaskName}"
        class="new-task"
        aria-label=${this.strings.newTaskPlaceholder}
        @input=${e=>this._newTaskName=e.target.value}>
      </fluent-text-field>`;this._groups.length>0&&!this._newTaskGroupId&&(this._newTaskGroupId=this._groups[0].id);const n=u.c`
      ${Object(D.a)(this._groups,e=>e.id,e=>u.c`<fluent-option value="${e.id}">${e.title}</fluent-option>`)}`,a=this.dataSource===On.todo?null:this._currentGroup?u.c`
          <span class="new-task-group">
            ${this.renderPlannerIcon(e)}
            <span>${this.getPlanTitle(this._currentGroup)}</span>
          </span>`:u.c`
            <fluent-select>
              <span slot="start">${this.renderPlannerIcon(e)}</span>
              ${this._groups.length>0?n:u.c`<fluent-option selected>No groups found</fluent-option>`}
            </fluent-select>`,i=this._folders.filter(e=>this._currentGroup&&e.parentId===this._currentGroup||!this._currentGroup&&e.parentId===this._newTaskGroupId);i.length>0&&!this._newTaskFolderId&&(this._newTaskFolderId=i[0].id);const r=u.c`
      ${Object(D.a)(i,e=>e.id,e=>u.c`<fluent-option value="${e.id}">${e.name}</fluent-option>`)}`,o=this._currentFolder?u.c`
          <span class="new-task-bucket">
            ${this.renderBucketIcon(e)}
            <span>${this.getFolderName(this._currentFolder)}</span>
          </span>
        `:u.c`
         <fluent-select>
          <span slot="start">${this.renderBucketIcon(e)}</span>
          ${i.length>0?r:u.c`<fluent-option selected>No folders found</fluent-option>`}
        </fluent-select>`,s={dark:this._isDarkMode,"new-task":!0},c=u.c`
      <fluent-text-field
        autocomplete="off"
        type="date"
        class=${Object(F.a)(s)}
        aria-label="${this.strings.addTaskDate}"
        .value="${this.dateToInputValue(this._newTaskDueDate)}"
        @change=${this.handleDateChange}>
      </fluent-text-field>`,d=this.dataSource===On.todo?null:this.renderAssignedPeople(null),l=this._newTaskBeingAdded?u.c`<div class="task-add-button-container"></div>`:u.c`
          <fluent-button
            class="add-task"
            @click=${this.onAddTaskClick}
            @keydown=${this.onAddTaskKeyDown}
            appearance="neutral">
              ${this.strings.addTaskButtonSubtitle}
          </fluent-button>
          <fluent-button
            class="cancel-task"
            @click=${()=>this.isNewTaskVisible=!1}
            @keydown=${this.newTaskVisible}
            appearance="neutral">
              ${this.strings.cancelNewTaskSubtitle}
          </fluent-button>`;return u.c`
    <div
      class=${Object(F.a)({task:!0,"new-task":!0})}>
      <div class="task-details-container">
        <div class="top add-new-task">
          <div class="check-and-title">
            ${t}
            <div class="task-content">
              <div class="task-group">${a}</div>
              <div class="task-bucket">${o}</div>
              ${d}
              <div class="task-due">${c}</div>
            </div>
          </div>
          <div class="task-options new-task-action-buttons">${l}</div>
        </div>
      </div>
    </div>
  `}togglePeoplePicker(e){const t=this.getPeoplePicker(e),n=this.getMgtPeople(e),a=this.getFlyout(e);t&&n&&a&&(a.isOpen?a.close():(t.selectedPeople=n.people,a.open(),setTimeout(()=>t.focus(),100)))}updateAssignedPeople(e){const t=this.getPeoplePicker(e),n=this.getMgtPeople(e);t&&t.selectedPeople!==n.people&&(n.people=t.selectedPeople,this.assignPeople(e,t.selectedPeople))}getPeoplePicker(e){const t=e?e.id:"new-task";return this.renderRoot.querySelector(`.picker-${t}`)}getMgtPeople(e){const t=e?e.id:"new-task";return this.renderRoot.querySelector(`.people-${t}`)}getFlyout(e){const t=e?e.id:"new-task";return this.renderRoot.querySelector(`.flyout-${t}`)}renderTask(e){const{name:t="Task",completed:n=!1,dueDate:a}=e,i=this._currentGroup?null:this.getPlanTitle(e.topParentId),r=this._currentFolder?null:this.getFolderName(e.immediateParentId),o={task:Object.assign(Object.assign({},e._raw),{groupTitle:i,folderTitle:r})},s=this.renderTemplate("task",o,e.id);if(s)return s;let c=this.renderTemplate("task-details",o,`task-details-${e.id}`);if(!c){const t="var(--neutral-foreground-hint)",n=this.dataSource===On.todo||this._currentGroup?null:u.c`
              <div class="task-group">
                <span class="task-icon">${this.renderPlannerIcon(t)}</span>
                <span class="task-icon-text">${this.getPlanTitle(e.topParentId)}</span>
              </div>
            `,i=this._currentFolder?null:u.c`
            <div class="task-bucket">
              <span class="task-icon">${this.renderBucketIcon(t)}</span>
              <span class="task-icon-text">${this.getFolderName(e.immediateParentId)}</span>
            </div>
          `,r=a?u.c`
            <div class="task-due">
              <span class="task-icon-text">${this.strings.due}${Object(Ke.n)(a)}</span>
            </div>
          `:null,o=this.dataSource===On.todo?null:this.renderAssignedPeople(e);c=u.c`${n} ${i} ${o} ${r}`}const d=this.readOnly||this.hideOptions?null:m.a`
            <mgt-dot-options
              class="dot-options"
              .options="${{[this.strings.removeTaskSubtitle]:()=>this.removeTask(e)}}"
            ></mgt-dot-options>`,l=Object(F.a)({task:!0,complete:n,incomplete:!n,"read-only":this.readOnly});return u.c`
      <div
        data-id="task-${e.id}"
        class=${l}
        @click=${()=>this.handleTaskClick(e)}>
        <div class="task-details-container">
          <div class="top">
            <div class="check-and-title">
              <fluent-checkbox
                @click=${t=>this.checkTask(t,e)}
                @keydown=${t=>this.handleTaskCheckKeyDown(t,e)}
                ?checked=${n}>
                  ${t}
              </fluent-checkbox>
            </div>
            <div class="task-options">${d}</div>
          </div>
          <div class="bottom">${c}</div>
        </div>
      </div>
    `}handleTaskCheckKeyDown(e,t){return An(this,void 0,void 0,function*(){"Enter"===e.key&&(this.readOnly||(t.completed?yield this.uncompleteTask(t):yield this.completeTask(t),e.stopPropagation(),e.preventDefault()))})}checkTask(e,t){return An(this,void 0,void 0,function*(){if(!this.readOnly){const n=this.shadowRoot.querySelector(`[data-id='task-${t.id}'`);n&&n.classList.add("updating"),t.completed?yield this.uncompleteTask(t):yield this.completeTask(t),n&&n.classList.remove("updating"),e.stopPropagation(),e.preventDefault()}})}renderAssignedPeople(e){var t;let n;const a={"new-task-assignee":null===e,"task-assignee":null!==e,"task-detail":null!==e},i=e?e.id:"new-task";a[`flyout-${i}`]=!0;const r=e?Object.keys(e.assignments).map(e=>e):[];if(!this.newTaskVisible){const a=null==e?void 0:e._raw,i=null==a?void 0:a.planId;i&&(n=null===(t=this._groups.filter(e=>e.id===i).pop())||void 0===t?void 0:t.containerId)}const o=this.isNewTaskVisible?this._newTaskContainerId:n,s=m.a`
      <mgt-people
        class="people people-${i}"
        .userIds=${r}
        .personCardInteraction=${w.a.none}
        @click=${t=>this.handlePeopleClick(t,e)}
        @keydown=${t=>this.handlePeopleKeydown(t,e)}
      >
        <template data-type="no-data">
          <fluent-button>
            <span style="display:flex;place-content:start;gap:4px;padding-inline-start:4px">
              <svg width="20" height="20" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg" class="svg" fill="currentColor">
                <path d="M9 2a4 4 0 100 8 4 4 0 000-8zM6 6a3 3 0 116 0 3 3 0 01-6 0z"></path>
                <path d="M4 11a2 2 0 00-2 2c0 1.7.83 2.97 2.13 3.8A9.14 9.14 0 009 18c.41 0 .82-.02 1.21-.06A5.5 5.5 0 019.6 17 12 12 0 019 17a8.16 8.16 0 01-4.33-1.05A3.36 3.36 0 013 13a1 1 0 011-1h5.6c.18-.36.4-.7.66-1H4z"></path>
                <path d="M14.5 19a4.5 4.5 0 100-9 4.5 4.5 0 000 9zm0-7c.28 0 .5.22.5.5V14h1.5a.5.5 0 010 1H15v1.5a.5.5 0 01-1 0V15h-1.5a.5.5 0 010-1H14v-1.5c0-.28.22-.5.5-.5z"></path>
              </svg> Assign</span>
            </fluent-button>
        </template>
      </mgt-people>`,c=m.a`
      <mgt-people-picker
        class="people-picker picker-${i}"
        .groupId=${Object($.a)(o)}
        @keydown=${this.handlePeoplePickerKeydown}>
      ></mgt-people-picker>`;return m.a`
      <mgt-flyout
        light-dismiss
        class=${Object(F.a)(a)}
        @closed=${()=>this.updateAssignedPeople(e)}
      >
        <div slot="anchor">${s}</div>
        <div slot="flyout" part="picker" class="picker">${c}</div>
      </mgt-flyout>
    `}handleTaskClick(e){e&&this.fireCustomEvent("taskClick",e)}renderLoadingTask(){return u.c`
      <div class="header">
        <div class="title">
          <fluent-skeleton shimmer class="shimmer" shape="rect"></fluent-skeleton>
        </div>
        <div class="new-task-button">
          <fluent-skeleton shimmer class="shimmer" shape="rect"></fluent-skeleton>
        </div>
      </div>
      <div class="tasks">
        <div class="task incomplete">
          <div class="task-details-container">
            <div class="top">
              <div class="check-and-title shimmer">
                <fluent-skeleton shimmer class="checkbox" shape="circle"></fluent-skeleton>
                <fluent-skeleton shimmer class="title" shape="rect"></fluent-skeleton>
              </div>
              <div class="task-options">
                <fluent-skeleton shimmer class="options" shape="rect"></fluent-skeleton>
              </div>
            </div>
            <div class="bottom">
              <div class="task-group">
                <div class="task-icon">
                  <fluent-skeleton shimmer class="shimmer icon" shape="rect"></fluent-skeleton>
                </div>
                <div class="task-icon-text">
                  <fluent-skeleton shimmer class="shimmer text" shape="rect"></fluent-skeleton>
                </div>
              </div>
              <div class="task-bucket">
                <div class="task-icon">
                  <fluent-skeleton shimmer class="shimmer icon" shape="rect"></fluent-skeleton>
                </div>
                <div class="task-icon-text">
                  <fluent-skeleton shimmer class="shimmer text" shape="rect"></fluent-skeleton>
                </div>
              </div>
              <div class="task-details shimmer">
                <fluent-skeleton shimmer class="shimmer icon" shape="circle"></fluent-skeleton>
                <fluent-skeleton shimmer class="shimmer icon" shape="circle"></fluent-skeleton>
                <fluent-skeleton shimmer class="shimmer icon" shape="circle"></fluent-skeleton>
              </div>
              <div class="task-due">
                <div class="task-icon-text">
                  <fluent-skeleton shimmer class="shimmer text" shape="rect"></fluent-skeleton>
                </div>
              </div>
              </div>
          </div>
        </div>
      </div>
    `}getTaskSource(){const e=_.a.globalProvider;if(!e||e.state!==i.c.SignedIn)return null;const t=e.graph.forComponent(this);return this.dataSource===On.planner?new ft(t):this.dataSource===On.todo?new pt(t):null}getPlanTitle(e){return e?e===this.res.PLANS_SELF_ASSIGNED?this.res.PLANS_SELF_ASSIGNED:(this._groups.find(t=>t.id===e)||{title:this.res.PLAN_NOT_FOUND}).title:this.res.BASE_SELF_ASSIGNED}getFolderName(e){return e?(this._folders.find(t=>t.id===e)||{name:this.res.BUCKET_NOT_FOUND}).name:this.res.BUCKETS_SELF_ASSIGNED}isTaskInSelectedGroupFilter(e){var t;return!this._currentGroup||e.topParentId===this._currentGroup||!this._currentGroup&&this.getTaskSource().isAssignedToMe(e,null===(t=this._me)||void 0===t?void 0:t.id)}isTaskInSelectedFolderFilter(e){return e.immediateParentId===this._currentFolder||!this._currentFolder}dateToInputValue(e){return e?new Date(e.getTime()-6e4*e.getTimezoneOffset()).toISOString().split("T")[0]:null}}wn([Object(f.b)({attribute:"read-only",type:Boolean}),En("design:type",Boolean)],Tn.prototype,"readOnly",void 0),wn([Object(f.b)({attribute:"data-source",converter:(e,t)=>(e=e.toLowerCase(),On[e]||On.planner)}),En("design:type",Number)],Tn.prototype,"dataSource",void 0),wn([Object(f.b)({attribute:"target-id",type:String}),En("design:type",String)],Tn.prototype,"targetId",void 0),wn([Object(f.b)({attribute:"target-bucket-id",type:String}),En("design:type",String)],Tn.prototype,"targetBucketId",void 0),wn([Object(f.b)({attribute:"initial-id",type:String}),En("design:type",String)],Tn.prototype,"initialId",void 0),wn([Object(f.b)({attribute:"initial-bucket-id",type:String}),En("design:type",String)],Tn.prototype,"initialBucketId",void 0),wn([Object(f.b)({attribute:"hide-header",type:Boolean}),En("design:type",Boolean)],Tn.prototype,"hideHeader",void 0),wn([Object(f.b)({attribute:"hide-options",type:Boolean}),En("design:type",Boolean)],Tn.prototype,"hideOptions",void 0),wn([Object(f.b)({attribute:"group-id",type:String}),En("design:type",String)],Tn.prototype,"groupId",void 0),wn([Object(f.b)(),En("design:type",Boolean)],Tn.prototype,"_isNewTaskVisible",void 0),wn([Object(f.b)(),En("design:type",Boolean)],Tn.prototype,"_newTaskBeingAdded",void 0),wn([Object(f.b)(),En("design:type",String)],Tn.prototype,"_newTaskName",void 0),wn([Object(f.b)(),En("design:type",Date)],Tn.prototype,"_newTaskDueDate",void 0),wn([Object(f.b)(),En("design:type",String)],Tn.prototype,"_newTaskGroupId",void 0),wn([Object(f.b)(),En("design:type",String)],Tn.prototype,"_newTaskFolderId",void 0),wn([Object(f.b)(),En("design:type",String)],Tn.prototype,"_newTaskContainerId",void 0),wn([Object(f.b)(),En("design:type",Array)],Tn.prototype,"_groups",void 0),wn([Object(f.b)(),En("design:type",Array)],Tn.prototype,"_folders",void 0),wn([Object(f.b)(),En("design:type",Array)],Tn.prototype,"_tasks",void 0),wn([Object(f.b)(),En("design:type",Array)],Tn.prototype,"_hiddenTasks",void 0),wn([Object(f.b)(),En("design:type",Array)],Tn.prototype,"_loadingTasks",void 0),wn([Object(f.b)(),En("design:type",Boolean)],Tn.prototype,"_inTaskLoad",void 0),wn([Object(f.b)(),En("design:type",Boolean)],Tn.prototype,"_hasDoneInitialLoad",void 0),wn([Object(f.b)(),En("design:type",Boolean)],Tn.prototype,"_todoDefaultSet",void 0),wn([Object(f.b)(),En("design:type",String)],Tn.prototype,"_currentGroup",void 0),wn([Object(f.b)(),En("design:type",String)],Tn.prototype,"_currentFolder",void 0),wn([Object(f.c)(),En("design:type",Object)],Tn.prototype,"_isDarkMode",void 0),wn([Object(f.c)(),En("design:type",Object)],Tn.prototype,"_me",void 0);const Un=[u.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host .container{display:flex;flex-direction:column;position:relative}:host .dropdown{display:none;position:absolute;z-index:1000;top:34px}:host .dropdown.visible{display:flex}:host .dropdown .team-photo{width:24px;position:inherit;border-radius:50%;margin:0 6px}:host .dropdown .team-start-slot{width:max-content}:host .dropdown .team-parent-name{width:auto}:host .loading-text,:host .search-error-text{font-style:normal;font-weight:400;font-size:14px;line-height:20px}:host .message-parent{display:flex;flex-direction:row;gap:5px;padding:5px}:host .message-parent .loading-text{margin:auto}:host fluent-card{background:var(--channel-picker-dropdown-background-color,var(--fill-color));padding:2px;--card-height:auto;--width:var(--card-width)}:host fluent-text-field{width:100%}:host fluent-text-field::part(root){background:padding-box linear-gradient(var(--channel-picker-input-background-color,var(--neutral-fill-input-rest)),var(--channel-picker-input-background-color,var(--neutral-fill-input-rest))),border-box var(--channel-picker-input-border-color,var(--neutral-stroke-input-rest))}:host fluent-text-field::part(root):hover{background:padding-box linear-gradient(var(--channel-picker-input-background-color-hover,var(--neutral-fill-input-hover)),var(--channel-picker-input-background-color-hover,var(--neutral-fill-input-hover))),border-box var(--channel-picker-input-hover-border-color,var(--neutral-stroke-input-hover))}:host fluent-text-field::part(root):focus,:host fluent-text-field::part(root):focus-within{background:padding-box linear-gradient(var(--channel-picker-input-background-color-focus,var(--neutral-fill-input-focus)),var(--channel-picker-input-background-color-focus,var(--neutral-fill-input-focus))),border-box var(--channel-picker-input-focus-border-color,var(--neutral-stroke-input-focus))}:host fluent-text-field::part(control){word-spacing:inherit;text-indent:inherit;letter-spacing:inherit}:host fluent-text-field::part(control)::placeholder{color:var(--channel-picker-input-placeholder-text-color,var(--input-placeholder-rest))}:host fluent-text-field::part(control):hover::placeholder{color:var(--channel-picker-input-placeholder-text-color-hover,var(--input-placeholder-hover))}:host fluent-text-field::part(control):focus-within::placeholder,:host fluent-text-field::part(control):focus::placeholder{color:var(--channel-picker-input-placeholder-text-color-focus,var(--input-placeholder-filled))}:host fluent-text-field .search-icon svg path{fill:var(--channel-picker-search-icon-color,currentColor)}:host fluent-text-field .down-chevron{height:auto;min-width:auto}:host fluent-text-field .down-chevron svg path{fill:var(--channel-picker-down-chevron-color,currentColor)}:host fluent-text-field .up-chevron{height:auto;min-width:auto}:host fluent-text-field .up-chevron svg path{fill:var(--channel-picker-up-chevron-color,currentColor)}:host fluent-text-field .close-icon{height:auto;min-width:auto}:host fluent-text-field .close-icon svg path{fill:var(--channel-picker-close-icon-color,currentColor)}:host fluent-tree-view{min-width:100%;--tree-item-nested-width:2em}:host fluent-tree-item{width:100%;--tree-item-nested-width:2em}:host fluent-tree-item:focus-visible{outline:0}:host fluent-tree-item::part(expand-collapse-button){background:0 0}:host fluent-tree-item::part(content-region),:host fluent-tree-item::part(positioning-region){color:var(--channel-picker-dropdown-item-text-color,currentColor);background:var(--channel-picker-dropdown-background-color,transparent);border:calc(var(--stroke-width) * 2px) solid transparent;height:auto}:host fluent-tree-item::part(content-region):hover,:host fluent-tree-item::part(positioning-region):hover{background:var(--channel-picker-dropdown-item-background-color-hover,var(--neutral-fill-stealth-hover))}:host fluent-tree-item::part(content-region):hover::part(expand-collapse-button),:host fluent-tree-item::part(positioning-region):hover::part(expand-collapse-button){background:var(--channel-picker-dropdown-item-background-color-hover,var(--neutral-fill-stealth-hover))}:host fluent-tree-item fluent-tree-item::part(content-region){height:auto}:host fluent-breadcrumb-item{color:var(--channel-picker-dropdown-item-text-color-selected,var(--neutral-foreground-rest))}:host fluent-breadcrumb-item .team-parent-name{font-weight:700}:host fluent-breadcrumb-item .team-photo{width:19px;position:inherit;border-radius:50%}:host fluent-breadcrumb-item .arrow{margin-left:8px;margin-right:8px}:host fluent-breadcrumb-item .arrow svg{stroke:var(--channel-picker-arrow-fill,var(--neutral-foreground-rest))}[dir=rtl] :host{--direction:rtl}[dir=rtl] .dropdown{text-align:right}[dir=rtl] .arrow{transform:scaleX(-1);filter:fliph;filter:FlipH;margin-right:0;margin-left:5px}[dir=rtl] .selected-team{padding-left:10px}[dir=rtl] .message-parent .loading-text{right:auto;left:10px;padding-right:8px;text-align:right}@media (forced-colors:active) and (prefers-color-scheme:dark){:host fluent-text-field svg{stroke:#fff!important}}@media (forced-colors:active) and (prefers-color-scheme:light){:host fluent-text-field svg{stroke:#000!important}}
`];var Fn=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const Hn={inputPlaceholderText:"Select a channel",noResultsFound:"We didn't find any matches.",loadingMessage:"Loading...",photoFor:"Teams photo for",teamsChannels:"Teams and channels results",closeButtonAriaLabel:"remove the selected channel"};class Rn extends ue.a{constructor(){super(...arguments),this.handleUnsupportedDelegatesFocus=()=>{var e;window.ShadowRoot&&!window.ShadowRoot.prototype.hasOwnProperty("delegatesFocus")&&(null===(e=this.$fastController.definition.shadowOptions)||void 0===e?void 0:e.delegatesFocus)&&(this.focus=()=>{var e;null===(e=this.control)||void 0===e||e.focus()})}}connectedCallback(){super.connectedCallback(),this.handleUnsupportedDelegatesFocus()}}Object(oe.a)([ce.c],Rn.prototype,"download",void 0),Object(oe.a)([ce.c],Rn.prototype,"href",void 0),Object(oe.a)([ce.c],Rn.prototype,"hreflang",void 0),Object(oe.a)([ce.c],Rn.prototype,"ping",void 0),Object(oe.a)([ce.c],Rn.prototype,"referrerpolicy",void 0),Object(oe.a)([ce.c],Rn.prototype,"rel",void 0),Object(oe.a)([ce.c],Rn.prototype,"target",void 0),Object(oe.a)([ce.c],Rn.prototype,"type",void 0),Object(oe.a)([se.d],Rn.prototype,"defaultSlottedContent",void 0);class Nn{}Object(oe.a)([Object(ce.c)({attribute:"aria-expanded"})],Nn.prototype,"ariaExpanded",void 0),Object(_e.a)(Nn,pe.a),Object(_e.a)(Rn,me.a,Nn);class Bn extends Rn{constructor(){super(...arguments),this.separator=!0}}Object(oe.a)([se.d],Bn.prototype,"separator",void 0),Object(_e.a)(Bn,me.a,Nn);class jn extends ue.a{slottedBreadcrumbItemsChanged(){if(this.$fastController.isConnected){if(void 0===this.slottedBreadcrumbItems||0===this.slottedBreadcrumbItems.length)return;const e=this.slottedBreadcrumbItems[this.slottedBreadcrumbItems.length-1];this.slottedBreadcrumbItems.forEach(t=>{const n=t===e;this.setItemSeparator(t,n),this.setAriaCurrent(t,n)})}}setItemSeparator(e,t){e instanceof Bn&&(e.separator=!t)}findChildWithHref(e){var t,n;return e.childElementCount>0?e.querySelector("a[href]"):(null===(t=e.shadowRoot)||void 0===t?void 0:t.childElementCount)?null===(n=e.shadowRoot)||void 0===n?void 0:n.querySelector("a[href]"):null}setAriaCurrent(e,t){const n=this.findChildWithHref(e);null===n&&e.hasAttribute("href")&&e instanceof Bn?t?e.setAttribute("aria-current","page"):e.removeAttribute("aria-current"):null!==n&&(t?n.setAttribute("aria-current","page"):n.removeAttribute("aria-current"))}}Object(oe.a)([se.d],jn.prototype,"slottedBreadcrumbItems",void 0);var Vn=n("Q5AN");const zn=jn.compose({baseName:"breadcrumb",template:(e,t)=>Se.a`
    <template role="navigation">
        <div role="list" class="list" part="list">
            <slot
                ${Object(De.a)({property:"slottedBreadcrumbItems",filter:Object(Vn.b)()})}
            ></slot>
        </div>
    </template>
`,styles:(e,t)=>Oe.a`
  ${Object(we.a)("inline-block")} :host {
    box-sizing: border-box;
    ${Tt.a};
  }

  .list {
    display: flex;
  }
`}),Gn=Bn.compose({baseName:"breadcrumb-item",template:(e,t)=>Se.a`
    <div role="listitem" class="listitem" part="listitem">
        ${Object(Ct.a)(e=>e.href&&e.href.length>0,Se.a`
                ${((e,t)=>Se.a`
    <a
        class="control"
        part="control"
        download="${e=>e.download}"
        href="${e=>e.href}"
        hreflang="${e=>e.hreflang}"
        ping="${e=>e.ping}"
        referrerpolicy="${e=>e.referrerpolicy}"
        rel="${e=>e.rel}"
        target="${e=>e.target}"
        type="${e=>e.type}"
        aria-atomic="${e=>e.ariaAtomic}"
        aria-busy="${e=>e.ariaBusy}"
        aria-controls="${e=>e.ariaControls}"
        aria-current="${e=>e.ariaCurrent}"
        aria-describedby="${e=>e.ariaDescribedby}"
        aria-details="${e=>e.ariaDetails}"
        aria-disabled="${e=>e.ariaDisabled}"
        aria-errormessage="${e=>e.ariaErrormessage}"
        aria-expanded="${e=>e.ariaExpanded}"
        aria-flowto="${e=>e.ariaFlowto}"
        aria-haspopup="${e=>e.ariaHaspopup}"
        aria-hidden="${e=>e.ariaHidden}"
        aria-invalid="${e=>e.ariaInvalid}"
        aria-keyshortcuts="${e=>e.ariaKeyshortcuts}"
        aria-label="${e=>e.ariaLabel}"
        aria-labelledby="${e=>e.ariaLabelledby}"
        aria-live="${e=>e.ariaLive}"
        aria-owns="${e=>e.ariaOwns}"
        aria-relevant="${e=>e.ariaRelevant}"
        aria-roledescription="${e=>e.ariaRoledescription}"
        ${Object(Ot.a)("control")}
    >
        ${Object(me.d)(e,t)}
        <span class="content" part="content">
            <slot ${Object(De.a)("defaultSlottedContent")}></slot>
        </span>
        ${Object(me.b)(e,t)}
    </a>
`)(e,t)}
            `)}
        ${Object(Ct.a)(e=>!e.href,Se.a`
                ${Object(me.d)(e,t)}
                <slot></slot>
                ${Object(me.b)(e,t)}
            `)}
        ${Object(Ct.a)(e=>e.separator,Se.a`
                <span class="separator" part="separator" aria-hidden="true">
                    <slot name="separator">${t.separator||""}</slot>
                </span>
            `)}
    </div>
`,styles:(e,t)=>Oe.a`
    ${Object(we.a)("inline-flex")} :host {
      background: transparent;
      color: ${Ee.fb};
      fill: currentcolor;
      box-sizing: border-box;
      ${Tt.a};
      min-width: calc(${Mt.a} * 1px);
      border-radius: calc(${Ee.q} * 1px);
    }

    .listitem {
      display: flex;
      align-items: center;
      border-radius: inherit;
    }

    .control {
      position: relative;
      align-items: center;
      box-sizing: border-box;
      color: inherit;
      fill: inherit;
      cursor: pointer;
      display: flex;
      outline: none;
      text-decoration: none;
      white-space: nowrap;
      border-radius: inherit;
    }

    .control:hover {
      color: ${Ee.eb};
    }

    .control:active {
      color: ${Ee.bb};
    }

    .control:${wt.a} {
      ${Ae.b}
    }

    :host(:not([href])),
    :host([aria-current]) .control {
      color: ${Ee.fb};
      fill: currentcolor;
      cursor: default;
    }

    .start {
      display: flex;
      margin-inline-end: 6px;
    }

    .end {
      display: flex;
      margin-inline-start: 6px;
    }

    .separator {
      display: flex;
    }
  `.withBehaviors(Object(At.a)(Oe.a`
        :host(:not([href])),
        .start,
        .end,
        .separator {
          background: ${Lt.a.ButtonFace};
          color: ${Lt.a.ButtonText};
          fill: currentcolor;
        }
        .separator {
          fill: ${Lt.a.ButtonText};
        }
        :host([href]) {
          forced-color-adjust: none;
          background: ${Lt.a.ButtonFace};
          color: ${Lt.a.LinkText};
        }
        :host([href]) .control:hover {
          background: ${Lt.a.LinkText};
          color: ${Lt.a.HighlightText};
          fill: currentcolor;
        }
        .control:${wt.a} {
          outline-color: ${Lt.a.LinkText};
        }
      `)),shadowOptions:{delegatesFocus:!0},separator:'\n    <svg width="12" height="12" xmlns="http://www.w3.org/2000/svg">\n      <path d="M4.65 2.15a.5.5 0 000 .7L7.79 6 4.65 9.15a.5.5 0 10.7.7l3.5-3.5a.5.5 0 000-.7l-3.5-3.5a.5.5 0 00-.7 0z"/>\n    </svg>\n  '});function Kn(e){return Object(fe.c)(e)&&"treeitem"===e.getAttribute("role")}class Wn extends ue.a{constructor(){super(...arguments),this.expanded=!1,this.focusable=!1,this.isNestedItem=()=>Kn(this.parentElement),this.handleExpandCollapseButtonClick=e=>{this.disabled||e.defaultPrevented||(this.expanded=!this.expanded)},this.handleFocus=e=>{this.setAttribute("tabindex","0")},this.handleBlur=e=>{this.setAttribute("tabindex","-1")}}expandedChanged(){this.$fastController.isConnected&&this.$emit("expanded-change",this)}selectedChanged(){this.$fastController.isConnected&&this.$emit("selected-change",this)}itemsChanged(e,t){this.$fastController.isConnected&&this.items.forEach(e=>{Kn(e)&&(e.nested=!0)})}static focusItem(e){e.focusable=!0,e.focus()}childItemLength(){const e=this.childItems.filter(e=>Kn(e));return e?e.length:0}}Object(oe.a)([Object(ce.c)({mode:"boolean"})],Wn.prototype,"expanded",void 0),Object(oe.a)([Object(ce.c)({mode:"boolean"})],Wn.prototype,"selected",void 0),Object(oe.a)([Object(ce.c)({mode:"boolean"})],Wn.prototype,"disabled",void 0),Object(oe.a)([se.d],Wn.prototype,"focusable",void 0),Object(oe.a)([se.d],Wn.prototype,"childItems",void 0),Object(oe.a)([se.d],Wn.prototype,"items",void 0),Object(oe.a)([se.d],Wn.prototype,"nested",void 0),Object(oe.a)([se.d],Wn.prototype,"renderCollapsedChildren",void 0),Object(_e.a)(Wn,me.a);class qn extends ue.a{constructor(){super(...arguments),this.currentFocused=null,this.handleFocus=e=>{if(!(this.slottedTreeItems.length<1))return e.target===this?(null===this.currentFocused&&(this.currentFocused=this.getValidFocusableItem()),void(null!==this.currentFocused&&Wn.focusItem(this.currentFocused))):void(this.contains(e.target)&&(this.setAttribute("tabindex","-1"),this.currentFocused=e.target))},this.handleBlur=e=>{e.target instanceof HTMLElement&&(null===e.relatedTarget||!this.contains(e.relatedTarget))&&this.setAttribute("tabindex","0")},this.handleKeyDown=e=>{if(e.defaultPrevented)return;if(this.slottedTreeItems.length<1)return!0;const t=this.getVisibleNodes();switch(e.key){case de.j:return void(t.length&&Wn.focusItem(t[0]));case de.f:return void(t.length&&Wn.focusItem(t[t.length-1]));case de.c:if(e.target&&this.isFocusableElement(e.target)){const t=e.target;t instanceof Wn&&t.childItemLength()>0&&t.expanded?t.expanded=!1:t instanceof Wn&&t.parentElement instanceof Wn&&Wn.focusItem(t.parentElement)}return!1;case de.d:if(e.target&&this.isFocusableElement(e.target)){const t=e.target;t instanceof Wn&&t.childItemLength()>0&&!t.expanded?t.expanded=!0:t instanceof Wn&&t.childItemLength()>0&&this.focusNextNode(1,e.target)}return;case de.b:return void(e.target&&this.isFocusableElement(e.target)&&this.focusNextNode(1,e.target));case de.e:return void(e.target&&this.isFocusableElement(e.target)&&this.focusNextNode(-1,e.target));case de.g:return void this.handleClick(e)}return!0},this.handleSelectedChange=e=>{if(e.defaultPrevented)return;if(!(e.target instanceof Element&&Kn(e.target)))return!0;const t=e.target;t.selected?(this.currentSelected&&this.currentSelected!==t&&(this.currentSelected.selected=!1),this.currentSelected=t):t.selected||this.currentSelected!==t||(this.currentSelected=null)},this.setItems=()=>{const e=this.treeView.querySelector("[aria-selected='true']");this.currentSelected=e,null!==this.currentFocused&&this.contains(this.currentFocused)||(this.currentFocused=this.getValidFocusableItem()),this.nested=this.checkForNestedItems(),this.getVisibleNodes().forEach(e=>{Kn(e)&&(e.nested=this.nested)})},this.isFocusableElement=e=>Kn(e),this.isSelectedElement=e=>e.selected}slottedTreeItemsChanged(){this.$fastController.isConnected&&this.setItems()}connectedCallback(){super.connectedCallback(),this.setAttribute("tabindex","0"),Ie.a.queueUpdate(()=>{this.setItems()})}handleClick(e){if(e.defaultPrevented)return;if(!(e.target instanceof Element&&Kn(e.target)))return!0;const t=e.target;t.disabled||(t.selected=!t.selected)}focusNextNode(e,t){const n=this.getVisibleNodes();if(!n)return;const a=n[n.indexOf(t)+e];Object(fe.c)(a)&&Wn.focusItem(a)}getValidFocusableItem(){const e=this.getVisibleNodes();let t=e.findIndex(this.isSelectedElement);return-1===t&&(t=e.findIndex(this.isFocusableElement)),-1!==t?e[t]:null}checkForNestedItems(){return this.slottedTreeItems.some(e=>Kn(e)&&e.querySelector("[role='treeitem']"))}getVisibleNodes(){return Object(fe.b)(this,"[role='treeitem']")||[]}}Object(oe.a)([Object(ce.c)({attribute:"render-collapsed-nodes"})],qn.prototype,"renderCollapsedNodes",void 0),Object(oe.a)([se.d],qn.prototype,"currentSelected",void 0),Object(oe.a)([se.d],qn.prototype,"slottedTreeItems",void 0);const Qn=qn.compose({baseName:"tree-view",template:(e,t)=>Se.a`
    <template
        role="tree"
        ${Object(Ot.a)("treeView")}
        @keydown="${(e,t)=>e.handleKeyDown(t.event)}"
        @focusin="${(e,t)=>e.handleFocus(t.event)}"
        @focusout="${(e,t)=>e.handleBlur(t.event)}"
        @click="${(e,t)=>e.handleClick(t.event)}"
        @selected-change="${(e,t)=>e.handleSelectedChange(t.event)}"
    >
        <slot ${Object(De.a)("slottedTreeItems")}></slot>
    </template>
`,styles:(e,t)=>Oe.a`
  :host([hidden]) {
    display: none;
  }

  ${Object(we.a)("flex")} :host {
    flex-direction: column;
    align-items: stretch;
    min-width: fit-content;
    font-size: 0;
  }
`});var Yn=n("/w5G");class Jn extends Vn.a{constructor(e,t){super(e,t),this.observer=null,t.childList=!0}observe(){null===this.observer&&(this.observer=new MutationObserver(this.handleEvent.bind(this))),this.observer.observe(this.target,this.options)}disconnect(){this.observer.disconnect()}getNodes(){return"subtree"in this.options?Array.from(this.target.querySelectorAll(this.options.selector)):Array.from(this.target.childNodes)}}var Xn=n("8GQ4");const Zn=Oe.a`
  .expand-collapse-button svg {
    transform: rotate(0deg);
  }
  :host(.nested) .expand-collapse-button {
    left: var(--expand-collapse-button-nested-width, calc(${Mt.a} * -1px));
  }
  :host([selected])::after {
    left: calc(${Ee.y} * 1px);
  }
  :host([expanded]) > .positioning-region .expand-collapse-button svg {
    transform: rotate(90deg);
  }
`,$n=Oe.a`
  .expand-collapse-button svg {
    transform: rotate(180deg);
  }
  :host(.nested) .expand-collapse-button {
    right: var(--expand-collapse-button-nested-width, calc(${Mt.a} * -1px));
  }
  :host([selected])::after {
    right: calc(${Ee.y} * 1px);
  }
  :host([expanded]) > .positioning-region .expand-collapse-button svg {
    transform: rotate(90deg);
  }
`,ea=Oe.b`((${Ee.n} / 2) * ${Ee.s}) + ((${Ee.s} * ${Ee.r}) / 2)`,ta=Xn.a.create("tree-item-expand-collapse-hover").withDefault(e=>{const t=Ee.Z.getValueFor(e);return t.evaluate(e,t.evaluate(e).hover).hover}),na=Xn.a.create("tree-item-expand-collapse-selected-hover").withDefault(e=>{const t=Ee.U.getValueFor(e);return Ee.Z.getValueFor(e).evaluate(e,t.evaluate(e).rest).hover}),aa=Wn.compose({baseName:"tree-item",template:(e,t)=>{return Se.a`
    <template
        role="treeitem"
        slot="${e=>e.isNestedItem()?"item":void 0}"
        tabindex="-1"
        class="${e=>e.expanded?"expanded":""} ${e=>e.selected?"selected":""} ${e=>e.nested?"nested":""}
            ${e=>e.disabled?"disabled":""}"
        aria-expanded="${e=>e.childItems&&e.childItemLength()>0?e.expanded:void 0}"
        aria-selected="${e=>e.selected}"
        aria-disabled="${e=>e.disabled}"
        @focusin="${(e,t)=>e.handleFocus(t.event)}"
        @focusout="${(e,t)=>e.handleBlur(t.event)}"
        ${n={property:"childItems",filter:Object(Vn.b)()},"string"==typeof n&&(n={property:n}),new Yn.a("fast-children",Jn,n)}
    >
        <div class="positioning-region" part="positioning-region">
            <div class="content-region" part="content-region">
                ${Object(Ct.a)(e=>e.childItems&&e.childItemLength()>0,Se.a`
                        <div
                            aria-hidden="true"
                            class="expand-collapse-button"
                            part="expand-collapse-button"
                            @click="${(e,t)=>e.handleExpandCollapseButtonClick(t.event)}"
                            ${Object(Ot.a)("expandCollapseButton")}
                        >
                            <slot name="expand-collapse-glyph">
                                ${t.expandCollapseGlyph||""}
                            </slot>
                        </div>
                    `)}
                ${Object(me.d)(e,t)}
                <slot></slot>
                ${Object(me.b)(e,t)}
            </div>
        </div>
        ${Object(Ct.a)(e=>e.childItems&&e.childItemLength()>0&&(e.expanded||e.renderCollapsedChildren),Se.a`
                <div role="group" class="items" part="items">
                    <slot name="item" ${Object(De.a)("items")}></slot>
                </div>
            `)}
    </template>
`;var n},styles:(e,t)=>Oe.a`
    ${Object(we.a)("block")} :host {
      contain: content;
      position: relative;
      outline: none;
      color: ${Ee.fb};
      fill: currentcolor;
      cursor: pointer;
      font-family: ${Ee.p};
      --expand-collapse-button-size: calc(${Mt.a} * 1px);
      --tree-item-nested-width: 0;
    }

    .positioning-region {
      display: flex;
      position: relative;
      box-sizing: border-box;
      background: ${Ee.ab};
      border: calc(${Ee.vb} * 1px) solid transparent;
      border-radius: calc(${Ee.q} * 1px);
      height: calc((${Mt.a} + 1) * 1px);
    }

    :host(:${wt.a}) .positioning-region {
      ${Ae.a}
    }

    .positioning-region::before {
      content: '';
      display: block;
      width: var(--tree-item-nested-width);
      flex-shrink: 0;
    }

    :host(:not([disabled])) .positioning-region:hover {
      background: ${Ee.Y};
    }

    :host(:not([disabled])) .positioning-region:active {
      background: ${Ee.W};
    }

    .content-region {
      display: inline-flex;
      align-items: center;
      white-space: nowrap;
      width: 100%;
      height: calc(${Mt.a} * 1px);
      margin-inline-start: calc(${Ee.s} * 2px + 8px);
      ${Tt.a}
    }

    .items {
      display: none;
      ${""} font-size: calc(1em + (${Ee.s} + 16) * 1px);
    }

    .expand-collapse-button {
      background: none;
      border: none;
      border-radius: calc(${Ee.q} * 1px);
      ${""} width: calc((${ea} + (${Ee.s} * 2)) * 1px);
      height: calc((${ea} + (${Ee.s} * 2)) * 1px);
      padding: 0;
      display: flex;
      justify-content: center;
      align-items: center;
      cursor: pointer;
      margin: 0 6px;
    }

    .expand-collapse-button svg {
      transition: transform 0.1s linear;
      pointer-events: none;
    }

    .start,
    .end {
      display: flex;
    }

    .start {
      ${""} margin-inline-end: calc(${Ee.s} * 2px + 2px);
    }

    .end {
      ${""} margin-inline-start: calc(${Ee.s} * 2px + 2px);
    }

    :host(.expanded) > .items {
      display: block;
    }

    :host([disabled]) {
      opacity: ${Ee.u};
      cursor: ${Et.a};
    }

    :host(.nested) .content-region {
      position: relative;
      margin-inline-start: var(--expand-collapse-button-size);
    }

    :host(.nested) .expand-collapse-button {
      position: absolute;
    }

    :host(.nested) .expand-collapse-button:hover {
      background: ${ta};
    }

    :host(:not([disabled])[selected]) .positioning-region {
      background: ${Ee.V};
    }

    :host(:not([disabled])[selected]) .expand-collapse-button:hover {
      background: ${na};
    }

    :host([selected])::after {
      content: '';
      display: block;
      position: absolute;
      top: calc((${Mt.a} / 4) * 1px);
      width: 3px;
      height: calc((${Mt.a} / 2) * 1px);
      ${""} background: ${Ee.e};
      border-radius: calc(${Ee.q} * 1px);
    }

    ::slotted(fluent-tree-item) {
      --tree-item-nested-width: 1em;
      --expand-collapse-button-nested-width: calc(${Mt.a} * -1px);
    }
  `.withBehaviors(new zt(Zn,$n),Object(At.a)(Oe.a`
        :host {
          color: ${Lt.a.ButtonText};
        }
        .positioning-region {
          border-color: ${Lt.a.ButtonFace};
          background: ${Lt.a.ButtonFace};
        }
        :host(:not([disabled])) .positioning-region:hover,
        :host(:not([disabled])) .positioning-region:active,
        :host(:not([disabled])[selected]) .positioning-region {
          background: ${Lt.a.Highlight};
        }
        :host .positioning-region:hover .content-region,
        :host([selected]) .positioning-region .content-region {
          forced-color-adjust: none;
          color: ${Lt.a.HighlightText};
        }
        :host([disabled][selected]) .positioning-region .content-region {
          color: ${Lt.a.GrayText};
        }
        :host([selected])::after {
          background: ${Lt.a.HighlightText};
        }
        :host(:${wt.a}) .positioning-region {
          forced-color-adjust: none;
          outline-color: ${Lt.a.ButtonFace};
        }
        :host([disabled]),
        :host([disabled]) .content-region,
        :host([disabled]) .positioning-region:hover .content-region {
          opacity: 1;
          color: ${Lt.a.GrayText};
        }
        :host(.nested) .expand-collapse-button:hover,
        :host(:not([disabled])[selected]) .expand-collapse-button:hover {
          background: ${Lt.a.ButtonFace};
          fill: ${Lt.a.ButtonText};
        }
      `)),expandCollapseGlyph:'\n    <svg width="12" height="12" xmlns="http://www.w3.org/2000/svg">\n      <path d="M4.65 2.15a.5.5 0 000 .7L7.79 6 4.65 9.15a.5.5 0 10.7.7l3.5-3.5a.5.5 0 000-.7l-3.5-3.5a.5.5 0 00-.7 0z"/>\n    </svg>\n  '});var ia=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},ra=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)},oa=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const sa=()=>{Object(T.a)(zn,Gn,U.a,Qn,aa,qe.a),Object(Ye.b)(),Object(A.b)("teams-channel-picker",ca)};class ca extends p.a{static get styles(){return Un}get strings(){return Hn}get selectedItem(){return this._selectedItemState?{channel:this._selectedItemState.item,team:this._selectedItemState.parent.item}:null}static get requiredScopes(){return["team.readbasic.all","channel.readbasic.all"]}set items(e){this._items!==e&&(this._items=e,this._treeViewState=e?this.generateTreeViewState(e):[],this.resetFocusState())}get items(){return this._items}get _inputWrapper(){return this.renderRoot.querySelector("fluent-text-field")}get _input(){return this._inputWrapper.shadowRoot.querySelector("input")}constructor(){super(),this.teamsPhotos={},this._inputValue="",this._treeViewState=[],this._focusList=[],this.handleInputClick=e=>{e.stopPropagation(),this.gainedFocus()},this.handleInputKeydown=e=>{const t=e.key;["ArrowDown","Enter"].includes(t)?this._isDropdownVisible?this.renderRoot.querySelector("fluent-tree-item").focus():this.gainedFocus():"Escape"===t&&this.lostFocus()},this.onClickCloseButton=()=>{this.removeSelectedChannel(null)},this.onKeydownCloseButton=e=>{"Enter"===e.key&&this.removeSelectedChannel(null)},this.onKeydownTreeView=e=>{"Escape"===e.key&&this.lostFocus()},this.handleTeamTreeItemClick=e=>{e.preventDefault(),e.stopImmediatePropagation();const t=e.target;t&&(t.getAttribute("expanded")?t.removeAttribute("expanded"):t.setAttribute("expanded","true"),t.removeAttribute("selected"),t.getAttribute("id")&&t.setAttribute("selected","true"))},this.handleInputChanged=e=>{const t=e.target;this._inputValue!==(null==t?void 0:t.value)&&(this._inputValue=null==t?void 0:t.value,this.gainedFocus(),this.debouncedSearch||(this.debouncedSearch=Object(Ke.b)(()=>{this.filterList()},400)),this.debouncedSearch())},this.handleWindowClick=e=>{e.target!==this&&this.lostFocus()},this.gainedFocus=()=>{const e=this._input;e&&e.focus(),this._isDropdownVisible=!0,this.toggleChevron(),this.resetFocusState(),this.requestUpdate()},this.lostFocus=()=>{this._input.value=this._inputValue="",this._input.textContent="",this._inputWrapper.value="",this._isDropdownVisible=!1,this.filterList(),this.toggleChevron(),this.requestUpdate(),void 0!==this._selectedItemState&&this.showCloseIcon()},this.handleFocus=()=>{this.lostFocus(),this.gainedFocus()},this.handleUpChevronClick=e=>{e.stopPropagation(),this.lostFocus()},this.addEventListener("focus",()=>this.loadTeamsIfNotLoaded()),this.addEventListener("mouseover",()=>this.loadTeamsIfNotLoaded()),this.addEventListener("blur",()=>this.lostFocus()),this._inputValue="",this._treeViewState=[],this._focusList=[],this._isDropdownVisible=!1}connectedCallback(){super.connectedCallback(),window.addEventListener("click",this.handleWindowClick);const e=this.renderRoot.ownerDocument;e&&e.documentElement.setAttribute("dir",this.direction)}disconnectedCallback(){window.removeEventListener("click",this.handleWindowClick),super.disconnectedCallback()}selectChannelById(e){return oa(this,void 0,void 0,function*(){const t=_.a.globalProvider;if(t&&t.state===i.c.SignedIn){this.items||(yield this.requestStateUpdate());for(const t of this._treeViewState)for(const n of t.channels)if(n.item.id===e)return t.isExpanded=!0,this.selectChannel(n),this.markSelectedChannelInDropdown(e),!0}return!1})}markSelectedChannelInDropdown(e){const t=this.renderRoot.querySelector(`[id='${e}']`);t&&(t.setAttribute("selected","true"),t.parentElement&&t.parentElement.setAttribute("expanded","true"))}render(){var e;const t={dropdown:!0,visible:this._isDropdownVisible};return this.renderTemplate("default",{teams:null!==(e=this.items)&&void 0!==e?e:[]})||u.c`
        <div class="container" @blur=${this.lostFocus}>
          <fluent-text-field
            autocomplete="off"
            appearance="outline"
            id="teams-channel-picker-input"
            role="combobox"
            placeholder="${this._selectedItemState?"":this.strings.inputPlaceholderText} "
            aria-label=${this.strings.inputPlaceholderText}
            aria-expanded="${this._isDropdownVisible}"
            label="teams-channel-picker-input"
            @focus=${this.handleFocus}
            @keyup=${this.handleInputChanged}
            @click=${this.handleInputClick}
            @keydown=${this.handleInputKeydown}
          >
            <div tabindex="0" slot="start" style="width: max-content;">${this.renderSelected()}</div>
            <div tabindex="0" slot="end">${this.renderChevrons()}${this.renderCloseButton()}</div>
          </fluent-text-field>
          <fluent-card
            class=${Object(F.a)(t)}
          >
            ${this.renderDropdown()}
          </fluent-card>
        </div>`}renderSelected(){var e,t,n,a,i,r;if(!this._selectedItemState)return this.renderSearchIcon();let o;if(this._selectedItemState.parent.channels){const t=null===(e=this.teamsPhotos[this._selectedItemState.parent.item.id])||void 0===e?void 0:e.photo;o=u.c`<img
        class="team-photo"
        alt="${this._selectedItemState.parent.item.displayName}"
        role="img"
        src=${t} />`}const s=null===(a=null===(n=null===(t=this._selectedItemState)||void 0===t?void 0:t.parent)||void 0===n?void 0:n.item)||void 0===a?void 0:a.displayName.trim(),c=null===(r=null===(i=this._selectedItemState)||void 0===i?void 0:i.item)||void 0===r?void 0:r.displayName.trim();return u.c`
      <fluent-breadcrumb title=${this._selectedItemState.item.displayName}>
        <fluent-breadcrumb-item>
          <span slot="start">${o}</span>
          <span class="team-parent-name">${s}</span>
          <span slot="separator" class="arrow">${Object(S.b)(S.a.TeamSeparator,"#000000")}</span>
        </fluent-breadcrumb-item>
        <fluent-breadcrumb-item>${c}</fluent-breadcrumb-item>
      </fluent-breadcrumb>`}clearState(){this._inputValue="",this._treeViewState=[],this._focusList=[],this._isDropdownVisible=!1}renderSearchIcon(){return u.c`
      <div class="search-icon">
        ${Object(S.b)(S.a.Search,"#252424")}
      </div>
    `}renderCloseButton(){return u.c`
      <fluent-button
        appearance="stealth"
        class="close-icon"
        style="display:none"
        aria-label=${this.strings.closeButtonAriaLabel}
        @click=${this.onClickCloseButton}
        @keydown=${this.onKeydownCloseButton}>
        <svg width="8" height="8" viewBox="0 0 8 8" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M0.0885911 0.215694L0.146447 0.146447C0.320013 -0.0271197 0.589437 -0.046405 0.784306 0.0885911L0.853553 0.146447L4 3.293L7.14645 0.146447C7.34171 -0.0488154 7.65829 -0.0488154 7.85355 0.146447C8.04882 0.341709 8.04882 0.658291 7.85355 0.853553L4.707 4L7.85355 7.14645C8.02712 7.32001 8.0464 7.58944 7.91141 7.78431L7.85355 7.85355C7.67999 8.02712 7.41056 8.0464 7.21569 7.91141L7.14645 7.85355L4 4.707L0.853553 7.85355C0.658291 8.04882 0.341709 8.04882 0.146447 7.85355C-0.0488154 7.65829 -0.0488154 7.34171 0.146447 7.14645L3.293 4L0.146447 0.853553C-0.0271197 0.679987 -0.046405 0.410563 0.0885911 0.215694L0.146447 0.146447L0.0885911 0.215694Z" fill="#212121"/>
        </svg>
      </fluent-button>
    `}showCloseIcon(){const e=this.renderRoot.querySelector(".down-chevron"),t=this.renderRoot.querySelector(".up-chevron"),n=this.renderRoot.querySelector(".close-icon");e&&(e.style.display="none"),t&&(t.style.display="none"),n&&(n.style.display=null)}renderDownChevron(){return u.c`
      <fluent-button appearance="stealth" class="down-chevron" @click=${this.gainedFocus}>
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M2.21967 4.46967C2.51256 4.17678 2.98744 4.17678 3.28033 4.46967L6 7.18934L8.71967 4.46967C9.01256 4.17678 9.48744 4.17678 9.78033 4.46967C10.0732 4.76256 10.0732 5.23744 9.78033 5.53033L6.53033 8.78033C6.23744 9.07322 5.76256 9.07322 5.46967 8.78033L2.21967 5.53033C1.92678 5.23744 1.92678 4.76256 2.21967 4.46967Z" fill="#212121" />
        </svg>
      </fluent-button>`}renderUpChevron(){return u.c`
      <fluent-button appearance="stealth" style="display:none" class="up-chevron" @click=${this.handleUpChevronClick}>
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M2.21967 7.53033C2.51256 7.82322 2.98744 7.82322 3.28033 7.53033L6 4.81066L8.71967 7.53033C9.01256 7.82322 9.48744 7.82322 9.78033 7.53033C10.0732 7.23744 10.0732 6.76256 9.78033 6.46967L6.53033 3.21967C6.23744 2.92678 5.76256 2.92678 5.46967 3.21967L2.21967 6.46967C1.92678 6.76256 1.92678 7.23744 2.21967 7.53033Z" fill="#212121" />
        </svg>
      </fluent-button>`}renderChevrons(){return u.c`${this.renderUpChevron()}${this.renderDownChevron()}`}renderDropdown(){return this.isLoadingState||!this._treeViewState?this.renderLoading():this._treeViewState?!this.isLoadingState&&0===this._treeViewState.length&&this._inputValue.length>0?this.renderError():this.renderDropdownList(this._treeViewState):u.c``}renderDropdownList(e){if(e&&e.length>0){let t=null;return u.c`
        <fluent-tree-view
          class="tree-view"
          dir=${this.direction}
          aria-live="polite"
          aria-relevant="all"
          aria-atomic="true"
          aria-label=${this.strings.teamsChannels}
          aria-orientation="horizontal"
          @keydown=${this.onKeydownTreeView}>
          ${Object(D.a)(e,e=>null==e?void 0:e.item,e=>{var n;return e.channels&&(t=u.c`<img
                  class="team-photo"
                  alt="${this.strings.photoFor} ${e.item.displayName}"
                  src=${null===(n=this.teamsPhotos[e.item.id])||void 0===n?void 0:n.photo}
                />`),u.c`
                <fluent-tree-item
                  ?expanded=${null==e?void 0:e.isExpanded}
                  @click=${this.handleTeamTreeItemClick}>
                    <span slot="start">${t}</span>${e.item.displayName}
                    ${Object(D.a)(null==e?void 0:e.channels,e=>e.item,e=>this.renderItem(e))}
                </fluent-tree-item>`})}
        </fluent-tree-view>`}return null}renderItem(e){var t;return u.c`
      <fluent-tree-item
        id=${null===(t=null==e?void 0:e.item)||void 0===t?void 0:t.id}
        @keydown=${t=>this.onUserKeyDown(t,e)}
        @click=${()=>this.handleItemClick(e)}>
          ${null==e?void 0:e.item.displayName}
      </fluent-tree-item>`}renderError(){return this.renderTemplate("error",null,"error")||u.c`
        <div class="message-parent">
          <div
            label="search-error-text"
            aria-live="polite"
            aria-relevant="all"
            aria-atomic="true"
            class="search-error-text"
          >
            ${this.strings.noResultsFound}
          </div>
        </div>
      `}renderLoading(){return this.renderTemplate("loading",null,"loading")||m.a`
        <div class="message-parent">
          <mgt-spinner></mgt-spinner>
          <div label="loading-text" aria-label="loading" class="loading-text">
            ${this.strings.loadingMessage}
          </div>
        </div>
      `}loadState(){var e,t;return oa(this,void 0,void 0,function*(){const n=_.a.globalProvider;let a;if(n&&n.state===i.c.SignedIn){const i=n.graph.forComponent(this);a=yield((e,t)=>Fn(void 0,void 0,void 0,function*(){const n=yield e.api("/me/joinedTeams").select(["displayName","id","isArchived"]).middlewareOptions(Object(b.a)(...t)).get();return null==n?void 0:n.value}))(i,ca.requiredScopes),a=a.filter(e=>!e.isArchived);const r=it.a.fromGraph(i),o=a.map(e=>e.id);this.teamsPhotos=yield((e,t)=>Fn(void 0,void 0,void 0,function*(){let n,a={};if(Object(z.c)()){n=V.a.getCache(K.a.photos,K.a.photos.stores.teams);for(const e of t)try{const t=yield n.getValue(e);t&&Object(z.g)()>Date.now()-t.timeCached&&(a[e]=t)}catch(e){}if(Object.keys(a).length)return a}const i=["team.readbasic.all"];a={};for(const r of t)try{const t=yield Object(z.e)(e,`/teams/${r}`,i);Object(z.c)()&&t&&(yield n.putValue(r,t)),a[r]=t}catch(e){}return a}))(r,o);const s=i.createBatch();for(const e of a)s.get(e.id,`teams/${e.id}/channels`,ca.requiredScopes);const c=yield s.executeAll();this._items=[];for(const n of a){const a=c.get(n.id);(null===(t=null===(e=a.content)||void 0===e?void 0:e.value)||void 0===t?void 0:t.length)&&this.items.push({item:n,channels:a.content.value.map(e=>({item:e}))})}}this.filterList(),this.resetFocusState()})}removeSelectedChannel(e){this.selectChannel(e);const t=this.renderRoot.querySelectorAll("fluent-tree-item");t&&t.forEach(e=>{e.removeAttribute("expanded"),e.removeAttribute("selected")})}handleItemClick(e){e.channels?e.isExpanded=!e.isExpanded:(this.selectChannel(e),this.lostFocus())}onUserKeyDown(e,t){switch(e.code){case"Enter":this.selectChannel(t),this.resetFocusState(),this.lostFocus(),e.preventDefault();break;case"Backspace":0===this._inputValue.length&&this._selectedItemState&&(this.selectChannel(null),this.resetFocusState(),e.preventDefault())}}filterList(){this.items&&(this._treeViewState=this.generateTreeViewState(this.items,this._inputValue),this.resetFocusState())}generateTreeViewState(e,t="",n=null){const a=[];if(t=t.toLowerCase(),e)for(const i of e){let e;if(0===t.length||i.item.displayName.toLowerCase().includes(t))e={item:i.item,parent:n},i.channels&&(e.channels=this.generateTreeViewState(i.channels,"",e),e.isExpanded=t.length>0);else if(i.channels){const a={item:i.item,parent:n},r=this.generateTreeViewState(i.channels,t,a);r.length>0&&(e=a,e.channels=r,e.isExpanded=!0)}e&&a.push(e)}return a}generateFocusList(e){if(!e||0===e.length)return[];let t=[];for(const n of e)t.push(n),n.channels&&n.isExpanded&&(t=[...t,...this.generateFocusList(n.channels)]);return t}resetFocusState(){this._focusList=this.generateFocusList(this._treeViewState),this.requestUpdate()}loadTeamsIfNotLoaded(){this.items||this.isLoadingState||this.requestStateUpdate()}selectChannel(e){e&&this._selectedItemState!==e?this._input.setAttribute("disabled","true"):this._input.removeAttribute("disabled"),this._selectedItemState=e,this.lostFocus(),this.fireCustomEvent("selectionChanged",this.selectedItem)}hideCloseIcon(){const e=this.renderRoot.querySelector(".close-icon");e&&(e.style.display="none")}toggleChevron(){const e=this.renderRoot.querySelector(".down-chevron"),t=this.renderRoot.querySelector(".up-chevron");this._isDropdownVisible?(e&&(e.style.display="none"),t&&(t.style.display=null)):(e&&(e.style.display=null,this.hideCloseIcon()),t&&(t.style.display="none")),this.hideCloseIcon()}}ia([Object(f.c)(),ra("design:type",Object)],ca.prototype,"_selectedItemState",void 0),ia([Object(f.c)(),ra("design:type",Boolean)],ca.prototype,"_isDropdownVisible",void 0);const da={cancelNewTaskSubtitle:"Cancel",newTaskPlaceholder:"Add a task",addTaskButtonSubtitle:"Add",removeTaskSubtitle:"Delete Task"};class la{constructor(e){if(this.dayFormat="numeric",this.weekdayFormat="long",this.monthFormat="long",this.yearFormat="numeric",this.date=new Date,e)for(const t in e){const n=e[t];"date"===t?this.date=this.getDateObject(n):this[t]=n}}getDateObject(e){if("string"==typeof e){const t=e.split(/[/-]/);return t.length<3?new Date:new Date(parseInt(t[2],10),parseInt(t[0],10)-1,parseInt(t[1],10))}if("day"in e&&"month"in e&&"year"in e){const{day:t,month:n,year:a}=e;return new Date(a,n-1,t)}return e}getDate(e=this.date,t={weekday:this.weekdayFormat,month:this.monthFormat,day:this.dayFormat,year:this.yearFormat},n=this.locale){const a=this.getDateObject(e);if(!a.getTime())return"";const i=Object.assign({timeZone:Intl.DateTimeFormat().resolvedOptions().timeZone},t);return new Intl.DateTimeFormat(n,i).format(a)}getDay(e=this.date.getDate(),t=this.dayFormat,n=this.locale){return this.getDate({month:1,day:e,year:2020},{day:t},n)}getMonth(e=this.date.getMonth()+1,t=this.monthFormat,n=this.locale){return this.getDate({month:e,day:2,year:2020},{month:t},n)}getYear(e=this.date.getFullYear(),t=this.yearFormat,n=this.locale){return this.getDate({month:2,day:2,year:e},{year:t},n)}getWeekday(e=0,t=this.weekdayFormat,n=this.locale){const a=`1-${e+1}-2017`;return this.getDate(a,{weekday:t},n)}getWeekdays(e=this.weekdayFormat,t=this.locale){return Array(7).fill(null).map((n,a)=>this.getWeekday(a,e,t))}}class ua extends ue.a{constructor(){super(...arguments),this.dateFormatter=new la,this.readonly=!1,this.locale="en-US",this.month=(new Date).getMonth()+1,this.year=(new Date).getFullYear(),this.dayFormat="numeric",this.weekdayFormat="short",this.monthFormat="long",this.yearFormat="numeric",this.minWeeks=0,this.disabledDates="",this.selectedDates="",this.oneDayInMs=864e5}localeChanged(){this.dateFormatter.locale=this.locale}dayFormatChanged(){this.dateFormatter.dayFormat=this.dayFormat}weekdayFormatChanged(){this.dateFormatter.weekdayFormat=this.weekdayFormat}monthFormatChanged(){this.dateFormatter.monthFormat=this.monthFormat}yearFormatChanged(){this.dateFormatter.yearFormat=this.yearFormat}getMonthInfo(e=this.month,t=this.year){const n=e=>new Date(e.getFullYear(),e.getMonth(),1).getDay(),a=e=>{const t=new Date(e.getFullYear(),e.getMonth()+1,1);return new Date(t.getTime()-this.oneDayInMs).getDate()},i=new Date(t,e-1),r=new Date(t,e),o=new Date(t,e-2);return{length:a(i),month:e,start:n(i),year:t,previous:{length:a(o),month:o.getMonth()+1,start:n(o),year:o.getFullYear()},next:{length:a(r),month:r.getMonth()+1,start:n(r),year:r.getFullYear()}}}getDays(e=this.getMonthInfo(),t=this.minWeeks){t=t>10?10:t;const{start:n,length:a,previous:i,next:r}=e,o=[];let s=1-n;for(;s<a+1||o.length<t||o[o.length-1].length%7!=0;){const{month:t,year:n}=s<1?i:s>a?r:e,c=s<1?i.length+s:s>a?s-a:s,d=`${t}-${c}-${n}`,l={day:c,month:t,year:n,disabled:this.dateInString(d,this.disabledDates),selected:this.dateInString(d,this.selectedDates)},u=o[o.length-1];0===o.length||u.length%7==0?o.push([l]):u.push(l),s++}return o}dateInString(e,t){const n=t.split(",").map(e=>e.trim());return e="string"==typeof e?e:`${e.getMonth()+1}-${e.getDate()}-${e.getFullYear()}`,n.some(t=>t===e)}getDayClassNames(e,t){const{day:n,month:a,year:i,disabled:r,selected:o}=e;return["day",t===`${a}-${n}-${i}`&&"today",this.month!==a&&"inactive",r&&"disabled",o&&"selected"].filter(Boolean).join(" ")}getWeekdayText(){const e=this.dateFormatter.getWeekdays().map(e=>({text:e}));if("long"!==this.weekdayFormat){const t=this.dateFormatter.getWeekdays("long");e.forEach((e,n)=>{e.abbr=t[n]})}return e}handleDateSelect(e,t){e.preventDefault,this.$emit("dateselected",t)}handleKeydown(e,t){return e.key===de.g&&this.handleDateSelect(e,t),!0}}function fa(e,t,n){return{index:e,removed:t,addedCount:n}}Object(oe.a)([Object(ce.c)({mode:"boolean"})],ua.prototype,"readonly",void 0),Object(oe.a)([ce.c],ua.prototype,"locale",void 0),Object(oe.a)([Object(ce.c)({converter:ce.e})],ua.prototype,"month",void 0),Object(oe.a)([Object(ce.c)({converter:ce.e})],ua.prototype,"year",void 0),Object(oe.a)([Object(ce.c)({attribute:"day-format",mode:"fromView"})],ua.prototype,"dayFormat",void 0),Object(oe.a)([Object(ce.c)({attribute:"weekday-format",mode:"fromView"})],ua.prototype,"weekdayFormat",void 0),Object(oe.a)([Object(ce.c)({attribute:"month-format",mode:"fromView"})],ua.prototype,"monthFormat",void 0),Object(oe.a)([Object(ce.c)({attribute:"year-format",mode:"fromView"})],ua.prototype,"yearFormat",void 0),Object(oe.a)([Object(ce.c)({attribute:"min-weeks",converter:ce.e})],ua.prototype,"minWeeks",void 0),Object(oe.a)([Object(ce.c)({attribute:"disabled-dates"})],ua.prototype,"disabledDates",void 0),Object(oe.a)([Object(ce.c)({attribute:"selected-dates"})],ua.prototype,"selectedDates",void 0);function pa(e,t,n,a,i,r){let o=0,s=0;const c=Math.min(n-t,r-i);if(0===t&&0===i&&(o=function(e,t,n){for(let a=0;a<n;++a)if(e[a]!==t[a])return a;return n}(e,a,c)),n===e.length&&r===a.length&&(s=function(e,t,n){let a=e.length,i=t.length,r=0;for(;r<n&&e[--a]===t[--i];)r++;return r}(e,a,c-o)),i+=o,r-=s,(n-=s)-(t+=o)==0&&r-i==0)return pn.d;if(t===n){const e=fa(t,[],0);for(;i<r;)e.removed.push(a[i++]);return[e]}if(i===r)return[fa(t,[],n-t)];const d=function(e){let t=e.length-1,n=e[0].length-1,a=e[t][n];const i=[];for(;t>0||n>0;){if(0===t){i.push(2),n--;continue}if(0===n){i.push(3),t--;continue}const r=e[t-1][n-1],o=e[t-1][n],s=e[t][n-1];let c;c=o<s?o<r?o:r:s<r?s:r,c===r?(r===a?i.push(0):(i.push(1),a=r),t--,n--):c===o?(i.push(3),t--,a=o):(i.push(2),n--,a=s)}return i.reverse(),i}(function(e,t,n,a,i,r){const o=r-i+1,s=n-t+1,c=new Array(o);let d,l;for(let e=0;e<o;++e)c[e]=new Array(s),c[e][0]=e;for(let e=0;e<s;++e)c[0][e]=e;for(let n=1;n<o;++n)for(let r=1;r<s;++r)e[t+r-1]===a[i+n-1]?c[n][r]=c[n-1][r-1]:(d=c[n-1][r]+1,l=c[n][r-1]+1,c[n][r]=d<l?d:l);return c}(e,t,n,a,i,r)),l=[];let u,f=t,p=i;for(let e=0;e<d.length;++e)switch(d[e]){case 0:void 0!==u&&(l.push(u),u=void 0),f++,p++;break;case 1:void 0===u&&(u=fa(f,[],0)),u.addedCount++,f++,u.removed.push(a[p]),p++;break;case 2:void 0===u&&(u=fa(f,[],0)),u.addedCount++,f++;break;case 3:void 0===u&&(u=fa(f,[],0)),u.removed.push(a[p]),p++}return void 0!==u&&l.push(u),l}const ma=Array.prototype.push;function _a(e,t,n,a){const i=fa(t,n,a);let r=!1,o=0;for(let t=0;t<e.length;t++){const n=e[t];if(n.index+=o,r)continue;const a=(s=i.index,c=i.index+i.removed.length,d=n.index,l=n.index+n.addedCount,c<d||l<s?-1:c===d||l===s?0:s<d?c<l?c-d:l-d:l<c?l-s:c-s);if(a>=0){e.splice(t,1),t--,o-=n.addedCount-n.removed.length,i.addedCount+=n.addedCount-a;const s=i.removed.length+n.removed.length-a;if(i.addedCount||s){let e=n.removed;if(i.index<n.index){const t=i.removed.slice(0,n.index-i.index);ma.apply(t,e),e=t}if(i.index+i.removed.length>n.index+n.addedCount){const t=i.removed.slice(n.index+n.addedCount-i.index);ma.apply(e,t)}i.removed=e,n.index<i.index&&(i.index=n.index)}else r=!0}else if(i.index<n.index){r=!0,e.splice(t,0,i),t++;const a=i.addedCount-i.removed.length;n.index+=a,o+=a}}var s,c,d,l;r||e.push(i)}var ha=n("O/9U");let ba=!1;function ga(e,t){let n=e.index;const a=t.length;return n>a?n=a-e.addedCount:n<0&&(n=a+e.removed.length+n-e.addedCount),n<0&&(n=0),e.index=n,e}class va extends ha.b{constructor(e){super(e),this.oldCollection=void 0,this.splices=void 0,this.needsQueue=!0,this.call=this.flush,Reflect.defineProperty(e,"$fastController",{value:this,enumerable:!1})}subscribe(e){this.flush(),super.subscribe(e)}addSplice(e){void 0===this.splices?this.splices=[e]:this.splices.push(e),this.needsQueue&&(this.needsQueue=!1,Ie.a.queueUpdate(this))}reset(e){this.oldCollection=e,this.needsQueue&&(this.needsQueue=!1,Ie.a.queueUpdate(this))}flush(){const e=this.splices,t=this.oldCollection;if(void 0===e&&void 0===t)return;this.needsQueue=!0,this.splices=void 0,this.oldCollection=void 0;const n=void 0===t?function(e,t){let n=[];const a=function(e){const t=[];for(let n=0,a=e.length;n<a;n++){const a=e[n];_a(t,a.index,a.removed,a.addedCount)}return t}(t);for(let t=0,i=a.length;t<i;++t){const i=a[t];1!==i.addedCount||1!==i.removed.length?n=n.concat(pa(e,i.index,i.index+i.addedCount,i.removed,0,i.removed.length)):i.removed[0]!==e[i.index]&&n.push(i)}return n}(this.source,e):pa(this.source,0,this.source.length,t,0,t.length);this.notify(n)}}var ya=n("Ne8q");const Sa=Object.freeze({positioning:!1,recycle:!0});function Da(e,t,n,a){e.bind(t[n],a)}function Ia(e,t,n,a){const i=Object.create(a);i.index=n,i.length=t.length,e.bind(t[n],i)}class xa{constructor(e,t,n,a,i,r){this.location=e,this.itemsBinding=t,this.templateBinding=a,this.options=r,this.source=null,this.views=[],this.items=null,this.itemsObserver=null,this.originalContext=void 0,this.childContext=void 0,this.bindView=Da,this.itemsBindingObserver=se.b.binding(t,this,n),this.templateBindingObserver=se.b.binding(a,this,i),r.positioning&&(this.bindView=Ia)}bind(e,t){this.source=e,this.originalContext=t,this.childContext=Object.create(t),this.childContext.parent=e,this.childContext.parentContext=this.originalContext,this.items=this.itemsBindingObserver.observe(e,this.originalContext),this.template=this.templateBindingObserver.observe(e,this.originalContext),this.observeItems(!0),this.refreshAllViews()}unbind(){this.source=null,this.items=null,null!==this.itemsObserver&&this.itemsObserver.unsubscribe(this),this.unbindAllViews(),this.itemsBindingObserver.disconnect(),this.templateBindingObserver.disconnect()}handleChange(e,t){e===this.itemsBinding?(this.items=this.itemsBindingObserver.observe(this.source,this.originalContext),this.observeItems(),this.refreshAllViews()):e===this.templateBinding?(this.template=this.templateBindingObserver.observe(this.source,this.originalContext),this.refreshAllViews(!0)):this.updateViews(t)}observeItems(e=!1){if(!this.items)return void(this.items=pn.d);const t=this.itemsObserver,n=this.itemsObserver=se.b.getNotifier(this.items),a=t!==n;a&&null!==t&&t.unsubscribe(this),(a||e)&&n.subscribe(this)}updateViews(e){const t=this.childContext,n=this.views,a=this.bindView,i=this.items,r=this.template,o=this.options.recycle,s=[];let c=0,d=0;for(let l=0,u=e.length;l<u;++l){const u=e[l],f=u.removed;let p=0,m=u.index;const _=m+u.addedCount,h=n.splice(u.index,f.length),b=d=s.length+h.length;for(;m<_;++m){const e=n[m],l=e?e.firstChild:this.location;let u;o&&d>0?(p<=b&&h.length>0?(u=h[p],p++):(u=s[c],c++),d--):u=r.create(),n.splice(m,0,u),a(u,i,m,t),u.insertBefore(l)}h[p]&&s.push(...h.slice(p))}for(let e=c,t=s.length;e<t;++e)s[e].dispose();if(this.options.positioning)for(let e=0,t=n.length;e<t;++e){const a=n[e].context;a.length=t,a.index=e}}refreshAllViews(e=!1){const t=this.items,n=this.childContext,a=this.template,i=this.location,r=this.bindView;let o=t.length,s=this.views,c=s.length;if(0!==o&&!e&&this.options.recycle||(ya.a.disposeContiguousBatch(s),c=0),0===c){this.views=s=new Array(o);for(let e=0;e<o;++e){const o=a.create();r(o,t,e,n),s[e]=o,o.insertBefore(i)}}else{let e=0;for(;e<o;++e)if(e<c)r(s[e],t,e,n);else{const o=a.create();r(o,t,e,n),s.push(o),o.insertBefore(i)}const d=s.splice(e,c-e);for(e=0,o=d.length;e<o;++e)d[e].dispose()}}unbindAllViews(){const e=this.views;for(let t=0,n=e.length;t<n;++t)e[t].unbind()}}class Ca extends Yn.b{constructor(e,t,n){super(),this.itemsBinding=e,this.templateBinding=t,this.options=n,this.createPlaceholder=Ie.a.createBlockPlaceholder,function(){if(ba)return;ba=!0,se.b.setArrayObserverFactory(e=>new va(e));const e=Array.prototype;if(e.$fastPatch)return;Reflect.defineProperty(e,"$fastPatch",{value:1,enumerable:!1});const t=e.pop,n=e.push,a=e.reverse,i=e.shift,r=e.sort,o=e.splice,s=e.unshift;e.pop=function(){const e=this.length>0,n=t.apply(this,arguments),a=this.$fastController;return void 0!==a&&e&&a.addSplice(fa(this.length,[n],0)),n},e.push=function(){const e=n.apply(this,arguments),t=this.$fastController;return void 0!==t&&t.addSplice(ga(fa(this.length-arguments.length,[],arguments.length),this)),e},e.reverse=function(){let e;const t=this.$fastController;void 0!==t&&(t.flush(),e=this.slice());const n=a.apply(this,arguments);return void 0!==t&&t.reset(e),n},e.shift=function(){const e=this.length>0,t=i.apply(this,arguments),n=this.$fastController;return void 0!==n&&e&&n.addSplice(fa(0,[t],0)),t},e.sort=function(){let e;const t=this.$fastController;void 0!==t&&(t.flush(),e=this.slice());const n=r.apply(this,arguments);return void 0!==t&&t.reset(e),n},e.splice=function(){const e=o.apply(this,arguments),t=this.$fastController;return void 0!==t&&t.addSplice(ga(fa(+arguments[0],e,arguments.length>2?arguments.length-2:0),this)),e},e.unshift=function(){const e=s.apply(this,arguments),t=this.$fastController;return void 0!==t&&t.addSplice(ga(fa(0,[],arguments.length),this)),e}}(),this.isItemsBindingVolatile=se.b.isVolatileBinding(e),this.isTemplateBindingVolatile=se.b.isVolatileBinding(t)}createBehavior(e){return new xa(e,this.itemsBinding,this.isItemsBindingVolatile,this.templateBinding,this.isTemplateBindingVolatile,this.options)}}function Oa(e,t,n=Sa){return new Ca(e,"function"==typeof t?t:()=>t,Object.assign(Object.assign({},Sa),n))}const wa="sticky",Ea="default",Aa="columnheader",La="default",ka=Se.a`
    <template>
        ${e=>null===e.rowData||null===e.columnDefinition||null===e.columnDefinition.columnDataKey?null:e.rowData[e.columnDefinition.columnDataKey]}
    </template>
`,Ma=Se.a`
    <template>
        ${e=>null===e.columnDefinition?null:void 0===e.columnDefinition.title?e.columnDefinition.columnDataKey:e.columnDefinition.title}
    </template>
`;class Pa extends ue.a{constructor(){super(...arguments),this.cellType=Ea,this.rowData=null,this.columnDefinition=null,this.isActiveCell=!1,this.customCellView=null,this.updateCellStyle=()=>{this.style.gridColumn=this.gridColumn}}cellTypeChanged(){this.$fastController.isConnected&&this.updateCellView()}gridColumnChanged(){this.$fastController.isConnected&&this.updateCellStyle()}columnDefinitionChanged(e,t){this.$fastController.isConnected&&this.updateCellView()}connectedCallback(){var e;super.connectedCallback(),this.addEventListener(cn,this.handleFocusin),this.addEventListener(dn,this.handleFocusout),this.addEventListener(ln,this.handleKeydown),this.style.gridColumn=`${void 0===(null===(e=this.columnDefinition)||void 0===e?void 0:e.gridColumn)?0:this.columnDefinition.gridColumn}`,this.updateCellView(),this.updateCellStyle()}disconnectedCallback(){super.disconnectedCallback(),this.removeEventListener(cn,this.handleFocusin),this.removeEventListener(dn,this.handleFocusout),this.removeEventListener(ln,this.handleKeydown),this.disconnectCellView()}handleFocusin(e){if(!this.isActiveCell){if(this.isActiveCell=!0,this.cellType===Aa){if(null!==this.columnDefinition&&!0!==this.columnDefinition.headerCellInternalFocusQueue&&"function"==typeof this.columnDefinition.headerCellFocusTargetCallback){const e=this.columnDefinition.headerCellFocusTargetCallback(this);null!==e&&e.focus()}}else if(null!==this.columnDefinition&&!0!==this.columnDefinition.cellInternalFocusQueue&&"function"==typeof this.columnDefinition.cellFocusTargetCallback){const e=this.columnDefinition.cellFocusTargetCallback(this);null!==e&&e.focus()}this.$emit("cell-focused",this)}}handleFocusout(e){this===document.activeElement||this.contains(document.activeElement)||(this.isActiveCell=!1)}handleKeydown(e){if(!(e.defaultPrevented||null===this.columnDefinition||this.cellType===Ea&&!0!==this.columnDefinition.cellInternalFocusQueue||this.cellType===Aa&&!0!==this.columnDefinition.headerCellInternalFocusQueue))switch(e.key){case de.g:case de.i:if(this.contains(document.activeElement)&&document.activeElement!==this)return;if(this.cellType===Aa){if(void 0!==this.columnDefinition.headerCellFocusTargetCallback){const t=this.columnDefinition.headerCellFocusTargetCallback(this);null!==t&&t.focus(),e.preventDefault()}}else if(void 0!==this.columnDefinition.cellFocusTargetCallback){const t=this.columnDefinition.cellFocusTargetCallback(this);null!==t&&t.focus(),e.preventDefault()}break;case de.h:this.contains(document.activeElement)&&document.activeElement!==this&&(this.focus(),e.preventDefault())}}updateCellView(){if(this.disconnectCellView(),null!==this.columnDefinition)switch(this.cellType){case Aa:void 0!==this.columnDefinition.headerCellTemplate?this.customCellView=this.columnDefinition.headerCellTemplate.render(this,this):this.customCellView=Ma.render(this,this);break;case void 0:case"rowheader":case Ea:void 0!==this.columnDefinition.cellTemplate?this.customCellView=this.columnDefinition.cellTemplate.render(this,this):this.customCellView=ka.render(this,this)}}disconnectCellView(){null!==this.customCellView&&(this.customCellView.dispose(),this.customCellView=null)}}Object(oe.a)([Object(ce.c)({attribute:"cell-type"})],Pa.prototype,"cellType",void 0),Object(oe.a)([Object(ce.c)({attribute:"grid-column"})],Pa.prototype,"gridColumn",void 0),Object(oe.a)([se.d],Pa.prototype,"rowData",void 0),Object(oe.a)([se.d],Pa.prototype,"columnDefinition",void 0);class Ta extends ue.a{constructor(){super(...arguments),this.rowType=La,this.rowData=null,this.columnDefinitions=null,this.isActiveRow=!1,this.cellsRepeatBehavior=null,this.cellsPlaceholder=null,this.focusColumnIndex=0,this.refocusOnLoad=!1,this.updateRowStyle=()=>{this.style.gridTemplateColumns=this.gridTemplateColumns}}gridTemplateColumnsChanged(){this.$fastController.isConnected&&this.updateRowStyle()}rowTypeChanged(){this.$fastController.isConnected&&this.updateItemTemplate()}rowDataChanged(){null!==this.rowData&&this.isActiveRow&&(this.refocusOnLoad=!0)}cellItemTemplateChanged(){this.updateItemTemplate()}headerCellItemTemplateChanged(){this.updateItemTemplate()}connectedCallback(){super.connectedCallback(),null===this.cellsRepeatBehavior&&(this.cellsPlaceholder=document.createComment(""),this.appendChild(this.cellsPlaceholder),this.updateItemTemplate(),this.cellsRepeatBehavior=new Ca(e=>e.columnDefinitions,e=>e.activeCellItemTemplate,{positioning:!0}).createBehavior(this.cellsPlaceholder),this.$fastController.addBehaviors([this.cellsRepeatBehavior])),this.addEventListener("cell-focused",this.handleCellFocus),this.addEventListener(dn,this.handleFocusout),this.addEventListener(ln,this.handleKeydown),this.updateRowStyle(),this.refocusOnLoad&&(this.refocusOnLoad=!1,this.cellElements.length>this.focusColumnIndex&&this.cellElements[this.focusColumnIndex].focus())}disconnectedCallback(){super.disconnectedCallback(),this.removeEventListener("cell-focused",this.handleCellFocus),this.removeEventListener(dn,this.handleFocusout),this.removeEventListener(ln,this.handleKeydown)}handleFocusout(e){this.contains(e.target)||(this.isActiveRow=!1,this.focusColumnIndex=0)}handleCellFocus(e){this.isActiveRow=!0,this.focusColumnIndex=this.cellElements.indexOf(e.target),this.$emit("row-focused",this)}handleKeydown(e){if(e.defaultPrevented)return;let t=0;switch(e.key){case de.c:t=Math.max(0,this.focusColumnIndex-1),this.cellElements[t].focus(),e.preventDefault();break;case de.d:t=Math.min(this.cellElements.length-1,this.focusColumnIndex+1),this.cellElements[t].focus(),e.preventDefault();break;case de.j:e.ctrlKey||(this.cellElements[0].focus(),e.preventDefault());break;case de.f:e.ctrlKey||(this.cellElements[this.cellElements.length-1].focus(),e.preventDefault())}}updateItemTemplate(){this.activeCellItemTemplate=this.rowType===La&&void 0!==this.cellItemTemplate?this.cellItemTemplate:this.rowType===La&&void 0===this.cellItemTemplate?this.defaultCellItemTemplate:void 0!==this.headerCellItemTemplate?this.headerCellItemTemplate:this.defaultHeaderCellItemTemplate}}Object(oe.a)([Object(ce.c)({attribute:"grid-template-columns"})],Ta.prototype,"gridTemplateColumns",void 0),Object(oe.a)([Object(ce.c)({attribute:"row-type"})],Ta.prototype,"rowType",void 0),Object(oe.a)([se.d],Ta.prototype,"rowData",void 0),Object(oe.a)([se.d],Ta.prototype,"columnDefinitions",void 0),Object(oe.a)([se.d],Ta.prototype,"cellItemTemplate",void 0),Object(oe.a)([se.d],Ta.prototype,"headerCellItemTemplate",void 0),Object(oe.a)([se.d],Ta.prototype,"rowIndex",void 0),Object(oe.a)([se.d],Ta.prototype,"isActiveRow",void 0),Object(oe.a)([se.d],Ta.prototype,"activeCellItemTemplate",void 0),Object(oe.a)([se.d],Ta.prototype,"defaultCellItemTemplate",void 0),Object(oe.a)([se.d],Ta.prototype,"defaultHeaderCellItemTemplate",void 0),Object(oe.a)([se.d],Ta.prototype,"cellElements",void 0);class Ua extends ue.a{constructor(){super(),this.noTabbing=!1,this.generateHeader="default",this.rowsData=[],this.columnDefinitions=null,this.focusRowIndex=0,this.focusColumnIndex=0,this.rowsPlaceholder=null,this.generatedHeader=null,this.isUpdatingFocus=!1,this.pendingFocusUpdate=!1,this.rowindexUpdateQueued=!1,this.columnDefinitionsStale=!0,this.generatedGridTemplateColumns="",this.focusOnCell=(e,t,n)=>{if(0===this.rowElements.length)return this.focusRowIndex=0,void(this.focusColumnIndex=0);const a=Math.max(0,Math.min(this.rowElements.length-1,e)),i=this.rowElements[a].querySelectorAll('[role="cell"], [role="gridcell"], [role="columnheader"], [role="rowheader"]'),r=i[Math.max(0,Math.min(i.length-1,t))];n&&this.scrollHeight!==this.clientHeight&&(a<this.focusRowIndex&&this.scrollTop>0||a>this.focusRowIndex&&this.scrollTop<this.scrollHeight-this.clientHeight)&&r.scrollIntoView({block:"center",inline:"center"}),r.focus()},this.onChildListChange=(e,t)=>{e&&e.length&&(e.forEach(e=>{e.addedNodes.forEach(e=>{1===e.nodeType&&"row"===e.getAttribute("role")&&(e.columnDefinitions=this.columnDefinitions)})}),this.queueRowIndexUpdate())},this.queueRowIndexUpdate=()=>{this.rowindexUpdateQueued||(this.rowindexUpdateQueued=!0,Ie.a.queueUpdate(this.updateRowIndexes))},this.updateRowIndexes=()=>{let e=this.gridTemplateColumns;if(void 0===e){if(""===this.generatedGridTemplateColumns&&this.rowElements.length>0){const e=this.rowElements[0];this.generatedGridTemplateColumns=new Array(e.cellElements.length).fill("1fr").join(" ")}e=this.generatedGridTemplateColumns}this.rowElements.forEach((t,n)=>{const a=t;a.rowIndex=n,a.gridTemplateColumns=e,this.columnDefinitionsStale&&(a.columnDefinitions=this.columnDefinitions)}),this.rowindexUpdateQueued=!1,this.columnDefinitionsStale=!1}}static generateTemplateColumns(e){let t="";return e.forEach(e=>{t=`${t}${""===t?"":" "}1fr`}),t}noTabbingChanged(){this.$fastController.isConnected&&(this.noTabbing?this.setAttribute("tabIndex","-1"):this.setAttribute("tabIndex",this.contains(document.activeElement)||this===document.activeElement?"-1":"0"))}generateHeaderChanged(){this.$fastController.isConnected&&this.toggleGeneratedHeader()}gridTemplateColumnsChanged(){this.$fastController.isConnected&&this.updateRowIndexes()}rowsDataChanged(){null===this.columnDefinitions&&this.rowsData.length>0&&(this.columnDefinitions=Ua.generateColumns(this.rowsData[0])),this.$fastController.isConnected&&this.toggleGeneratedHeader()}columnDefinitionsChanged(){null!==this.columnDefinitions?(this.generatedGridTemplateColumns=Ua.generateTemplateColumns(this.columnDefinitions),this.$fastController.isConnected&&(this.columnDefinitionsStale=!0,this.queueRowIndexUpdate())):this.generatedGridTemplateColumns=""}headerCellItemTemplateChanged(){this.$fastController.isConnected&&null!==this.generatedHeader&&(this.generatedHeader.headerCellItemTemplate=this.headerCellItemTemplate)}focusRowIndexChanged(){this.$fastController.isConnected&&this.queueFocusUpdate()}focusColumnIndexChanged(){this.$fastController.isConnected&&this.queueFocusUpdate()}connectedCallback(){super.connectedCallback(),void 0===this.rowItemTemplate&&(this.rowItemTemplate=this.defaultRowItemTemplate),this.rowsPlaceholder=document.createComment(""),this.appendChild(this.rowsPlaceholder),this.toggleGeneratedHeader(),this.rowsRepeatBehavior=new Ca(e=>e.rowsData,e=>e.rowItemTemplate,{positioning:!0}).createBehavior(this.rowsPlaceholder),this.$fastController.addBehaviors([this.rowsRepeatBehavior]),this.addEventListener("row-focused",this.handleRowFocus),this.addEventListener(sn,this.handleFocus),this.addEventListener(ln,this.handleKeydown),this.addEventListener(dn,this.handleFocusOut),this.observer=new MutationObserver(this.onChildListChange),this.observer.observe(this,{childList:!0}),this.noTabbing&&this.setAttribute("tabindex","-1"),Ie.a.queueUpdate(this.queueRowIndexUpdate)}disconnectedCallback(){super.disconnectedCallback(),this.removeEventListener("row-focused",this.handleRowFocus),this.removeEventListener(sn,this.handleFocus),this.removeEventListener(ln,this.handleKeydown),this.removeEventListener(dn,this.handleFocusOut),this.observer.disconnect(),this.rowsPlaceholder=null,this.generatedHeader=null}handleRowFocus(e){this.isUpdatingFocus=!0;const t=e.target;this.focusRowIndex=this.rowElements.indexOf(t),this.focusColumnIndex=t.focusColumnIndex,this.setAttribute("tabIndex","-1"),this.isUpdatingFocus=!1}handleFocus(e){this.focusOnCell(this.focusRowIndex,this.focusColumnIndex,!0)}handleFocusOut(e){null!==e.relatedTarget&&this.contains(e.relatedTarget)||this.setAttribute("tabIndex",this.noTabbing?"-1":"0")}handleKeydown(e){if(e.defaultPrevented)return;let t;const n=this.rowElements.length-1,a=this.offsetHeight+this.scrollTop,i=this.rowElements[n];switch(e.key){case de.e:e.preventDefault(),this.focusOnCell(this.focusRowIndex-1,this.focusColumnIndex,!0);break;case de.b:e.preventDefault(),this.focusOnCell(this.focusRowIndex+1,this.focusColumnIndex,!0);break;case de.l:if(e.preventDefault(),0===this.rowElements.length){this.focusOnCell(0,0,!1);break}if(0===this.focusRowIndex)return void this.focusOnCell(0,this.focusColumnIndex,!1);for(t=this.focusRowIndex-1;t>=0;t--){const e=this.rowElements[t];if(e.offsetTop<this.scrollTop){this.scrollTop=e.offsetTop+e.clientHeight-this.clientHeight;break}}this.focusOnCell(t,this.focusColumnIndex,!1);break;case de.k:if(e.preventDefault(),0===this.rowElements.length){this.focusOnCell(0,0,!1);break}if(this.focusRowIndex>=n||i.offsetTop+i.offsetHeight<=a)return void this.focusOnCell(n,this.focusColumnIndex,!1);for(t=this.focusRowIndex+1;t<=n;t++){const e=this.rowElements[t];if(e.offsetTop+e.offsetHeight>a){let t=0;this.generateHeader===wa&&null!==this.generatedHeader&&(t=this.generatedHeader.clientHeight),this.scrollTop=e.offsetTop-t;break}}this.focusOnCell(t,this.focusColumnIndex,!1);break;case de.j:e.ctrlKey&&(e.preventDefault(),this.focusOnCell(0,0,!0));break;case de.f:e.ctrlKey&&null!==this.columnDefinitions&&(e.preventDefault(),this.focusOnCell(this.rowElements.length-1,this.columnDefinitions.length-1,!0))}}queueFocusUpdate(){this.isUpdatingFocus&&(this.contains(document.activeElement)||this===document.activeElement)||!1===this.pendingFocusUpdate&&(this.pendingFocusUpdate=!0,Ie.a.queueUpdate(()=>this.updateFocus()))}updateFocus(){this.pendingFocusUpdate=!1,this.focusOnCell(this.focusRowIndex,this.focusColumnIndex,!0)}toggleGeneratedHeader(){if(null!==this.generatedHeader&&(this.removeChild(this.generatedHeader),this.generatedHeader=null),"none"!==this.generateHeader&&this.rowsData.length>0){const e=document.createElement(this.rowElementTag);return this.generatedHeader=e,this.generatedHeader.columnDefinitions=this.columnDefinitions,this.generatedHeader.gridTemplateColumns=this.gridTemplateColumns,this.generatedHeader.rowType=this.generateHeader===wa?"sticky-header":"header",void(null===this.firstChild&&null===this.rowsPlaceholder||this.insertBefore(e,null!==this.firstChild?this.firstChild:this.rowsPlaceholder))}}}Ua.generateColumns=e=>Object.getOwnPropertyNames(e).map((e,t)=>({columnDataKey:e,gridColumn:`${t}`})),Object(oe.a)([Object(ce.c)({attribute:"no-tabbing",mode:"boolean"})],Ua.prototype,"noTabbing",void 0),Object(oe.a)([Object(ce.c)({attribute:"generate-header"})],Ua.prototype,"generateHeader",void 0),Object(oe.a)([Object(ce.c)({attribute:"grid-template-columns"})],Ua.prototype,"gridTemplateColumns",void 0),Object(oe.a)([se.d],Ua.prototype,"rowsData",void 0),Object(oe.a)([se.d],Ua.prototype,"columnDefinitions",void 0),Object(oe.a)([se.d],Ua.prototype,"rowItemTemplate",void 0),Object(oe.a)([se.d],Ua.prototype,"cellItemTemplate",void 0),Object(oe.a)([se.d],Ua.prototype,"headerCellItemTemplate",void 0),Object(oe.a)([se.d],Ua.prototype,"focusRowIndex",void 0),Object(oe.a)([se.d],Ua.prototype,"focusColumnIndex",void 0),Object(oe.a)([se.d],Ua.prototype,"defaultRowItemTemplate",void 0),Object(oe.a)([se.d],Ua.prototype,"rowElementTag",void 0),Object(oe.a)([se.d],Ua.prototype,"rowElements",void 0);const Fa=Se.a`
    <div
        class="title"
        part="title"
        aria-label="${e=>e.dateFormatter.getDate(`${e.month}-2-${e.year}`,{month:"long",year:"numeric"})}"
    >
        <span part="month">
            ${e=>e.dateFormatter.getMonth(e.month)}
        </span>
        <span part="year">${e=>e.dateFormatter.getYear(e.year)}</span>
    </div>
`,Ha=Oe.a`
.day.disabled::before {
  transform: translate(-50%, 0) rotate(45deg);
}
`,Ra=Oe.a`
.day.disabled::before {
  transform: translate(50%, 0) rotate(-45deg);
}
`;class Na extends ua{constructor(){super(...arguments),this.readonly=!0}}Object(bt.a)([Object(ce.c)({converter:ce.d})],Na.prototype,"readonly",void 0);const Ba=Na.compose({baseName:"calendar",template:(e,t)=>{var n;const a=new Date,i=`${a.getMonth()+1}-${a.getDate()}-${a.getFullYear()}`;return Se.a`
        <template>
            ${me.e}
            ${t.title instanceof Function?t.title(e,t):null!==(n=t.title)&&void 0!==n?n:""}
            <slot></slot>
            ${Object(Ct.a)(e=>e.readonly,(e=>Se.a`
        <div class="days" part="days">
            <div class="week-days" part="week-days">
                ${Oa(e=>e.getWeekdayText(),Se.a`
                        <div class="week-day" part="week-day" abbr="${e=>e.abbr}">
                            ${e=>e.text}
                        </div>
                    `)}
            </div>
            ${Oa(e=>e.getDays(),Se.a`
                    <div class="week">
                        ${Oa(e=>e,Se.a`
                                <div
                                    class="${(t,n)=>n.parentContext.parent.getDayClassNames(t,e)}"
                                    part="day"
                                    aria-label="${(e,t)=>t.parentContext.parent.dateFormatter.getDate(`${e.month}-${e.day}-${e.year}`,{month:"long",day:"numeric"})}"
                                >
                                    <div
                                        class="date"
                                        part="${t=>e===`${t.month}-${t.day}-${t.year}`?"today":"date"}"
                                    >
                                        ${(e,t)=>t.parentContext.parent.dateFormatter.getDay(e.day)}
                                    </div>
                                    <slot
                                        name="${e=>e.month}-${e=>e.day}-${e=>e.year}"
                                    ></slot>
                                </div>
                            `)}
                    </div>
                `)}
        </div>
    `)(i),((e,t)=>{const n=e.tagFor(Ua),a=e.tagFor(Ta);return Se.a`
    <${n} class="days interact" part="days" generate-header="none">
        <${a}
            class="week-days"
            part="week-days"
            role="row"
            row-type="header"
            grid-template-columns="1fr 1fr 1fr 1fr 1fr 1fr 1fr"
        >
            ${Oa(e=>e.getWeekdayText(),(e=>{const t=e.tagFor(Pa);return Se.a`
        <${t}
            class="week-day"
            part="week-day"
            tabindex="-1"
            grid-column="${(e,t)=>t.index+1}"
            abbr="${e=>e.abbr}"
        >
            ${e=>e.text}
        </${t}>
    `})(e),{positioning:!0})}
        </${a}>
        ${Oa(e=>e.getDays(),((e,t)=>{const n=e.tagFor(Ta);return Se.a`
        <${n}
            class="week"
            part="week"
            role="row"
            role-type="default"
            grid-template-columns="1fr 1fr 1fr 1fr 1fr 1fr 1fr"
        >
        ${Oa(e=>e,((e,t)=>{const n=e.tagFor(Pa);return Se.a`
        <${n}
            class="${(e,n)=>n.parentContext.parent.getDayClassNames(e,t)}"
            part="day"
            tabindex="-1"
            role="gridcell"
            grid-column="${(e,t)=>t.index+1}"
            @click="${(e,t)=>t.parentContext.parent.handleDateSelect(t.event,e)}"
            @keydown="${(e,t)=>t.parentContext.parent.handleKeydown(t.event,e)}"
            aria-label="${(e,t)=>t.parentContext.parent.dateFormatter.getDate(`${e.month}-${e.day}-${e.year}`,{month:"long",day:"numeric"})}"
        >
            <div
                class="date"
                part="${e=>t===`${e.month}-${e.day}-${e.year}`?"today":"date"}"
            >
                ${(e,t)=>t.parentContext.parent.dateFormatter.getDay(e.day)}
            </div>
            <slot name="${e=>e.month}-${e=>e.day}-${e=>e.year}"></slot>
        </${n}>
    `})(e,t),{positioning:!0})}
        </${n}>
    `})(e,t))}
    </${n}>
`})(e,i))}
            ${me.c}
        </template>
    `},styles:(e,t)=>Oe.a`
${Object(we.a)("inline-block")} :host {
  --calendar-cell-size: calc((${Ee.n} + 2 + ${Ee.r}) * ${Ee.s} * 1px);
  --calendar-gap: 2px;
  ${Tt.a}
  color: ${Ee.fb};
}

.title {
  padding: calc(${Ee.s} * 2px);
  font-weight: 600;
}

.days {
  text-align: center;
}

.week-days,
.week {
  display: grid;
  grid-template-columns: repeat(7, 1fr);
  grid-gap: var(--calendar-gap);
  border: 0;
  padding: 0;
}

.day,
.week-day {
  border: 0;
  width: var(--calendar-cell-size);
  height: var(--calendar-cell-size);
  line-height: var(--calendar-cell-size);
  padding: 0;
  box-sizing: initial;
}

.week-day {
  font-weight: 600;
}

.day {
  border: calc(${Ee.vb} * 1px) solid transparent;
  border-radius: calc(${Ee.q} * 1px);
}

.interact .day {
  cursor: pointer;
}

.date {
  height: 100%;
}

.inactive .date,
.inactive.disabled::before {
  color: ${Ee.cb};
}

.disabled::before {
  content: '';
  display: inline-block;
  width: calc(var(--calendar-cell-size) * .8);
  height: calc(${Ee.vb} * 1px);
  background: currentColor;
  position: absolute;
  margin-top: calc(var(--calendar-cell-size) / 2);
  transform-origin: center;
  z-index: 1;
}

.selected {
  color: ${Ee.e};
  border: 1px solid ${Ee.e};
  background: ${Ee.v};
}

.selected + .selected {
  border-start-start-radius: 0;
  border-end-start-radius: 0;
  border-inline-start-width: 0;
  padding-inline-start: calc(var(--calendar-gap) + (${Ee.vb} + ${Ee.q}) * 1px);
  margin-inline-start: calc((${Ee.q} * -1px) - var(--calendar-gap));
}

.today.disabled::before {
  color: ${Ee.C};
}

.today .date {
  color: ${Ee.C};
  background: ${Ee.e};
  border-radius: 50%;
  position: relative;
}
`.withBehaviors(Object(At.a)(Oe.a`
          .day.selected {
              color: ${Lt.a.Highlight};
          }

          .today .date {
              background: ${Lt.a.Highlight};
              color: ${Lt.a.HighlightText};
          }
      `),new zt(Ha,Ra)),title:Fa});var ja=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},Va=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};class za extends p.a{get strings(){return da}constructor(){super(),this.handleTaskClick=e=>{this.fireCustomEvent("taskClick",{task:e})},this.onResize=()=>{this.mediaQuery!==this._previousMediaQuery&&(this._previousMediaQuery=this.mediaQuery,this.requestUpdate())},Object(T.a)(qe.a,Me.a,Ba),this.clearState(),this._previousMediaQuery=this.mediaQuery}attributeChangedCallback(e,t,n){switch(super.attributeChangedCallback(e,t,n),e){case"target-id":case"initial-id":this.clearState(),this.requestStateUpdate()}}connectedCallback(){super.connectedCallback(),window.addEventListener("resize",this.onResize)}disconnectedCallback(){window.removeEventListener("resize",this.onResize),super.disconnectedCallback()}render(){const e=_.a.globalProvider;if(!e||e.state!==i.c.SignedIn)return u.c``;if(this.isLoadingState)return this.renderLoadingTask();const t=this.renderPicker(),n=this.renderNewTask(),a=this.isLoadingState?this.renderLoadingTask():this.renderTasks();return u.c`
      ${t}
      ${n}
      <div class="tasks" dir=${this.direction}>
        ${a}
      </div>
    `}renderLoadingTask(){return u.c`
      <div class="task loading-task">
        <div class="task-details">
          <div class="title"></div>
          <div class="task-due"></div>
          <div class="task-delete"></div>
        </div>
      </div>
    `}clearState(){this.requestUpdate()}dateToInputValue(e){return e?new Date(e.getTime()-6e4*e.getTimezoneOffset()).toISOString().split("T")[0]:null}}ja([Object(f.b)({attribute:"read-only",type:Boolean}),Va("design:type",Boolean)],za.prototype,"readOnly",void 0),ja([Object(f.b)({attribute:"hide-header",type:Boolean}),Va("design:type",Boolean)],za.prototype,"hideHeader",void 0),ja([Object(f.b)({attribute:"hide-options",type:Boolean}),Va("design:type",Boolean)],za.prototype,"hideOptions",void 0),ja([Object(f.b)({attribute:"target-id",type:String}),Va("design:type",String)],za.prototype,"targetId",void 0),ja([Object(f.b)({attribute:"initial-id",type:String}),Va("design:type",String)],za.prototype,"initialId",void 0);var Ga=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const Ka=(e,t)=>Ga(void 0,void 0,void 0,function*(){const n=yield e.api(`/me/todo/lists/${t}/tasks`).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Tasks.Read")).get();return null==n?void 0:n.value}),Wa=(e,t)=>Ga(void 0,void 0,void 0,function*(){return yield e.api(`/me/todo/lists/${t}`).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Tasks.Read")).get()}),qa=[u.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{display:flex;flex-direction:column;color:var(--color,var(--neutral-foreground-rest))}:host input[type=date]::-webkit-calendar-picker-indicator,:host input[type=date]::-webkit-inner-spin-button{display:none;appearance:none}:host .task-icon{font-family:FabricMDL2Icons;user-select:none}:host .task-icon.divider{vertical-align:initial;margin:0 12px;font-size:16px}:host .header{margin:var(--tasks-header-margin,0 0 12px 0);padding:var(--tasks-title-padding,0);display:flex;align-items:center;justify-content:space-between}:host .header .header__loading{max-width:90px;width:100%;height:20px;background:#f2f2f2}:host .header select{font-size:var(--tasks-plan-title-font-size,1.1em);padding:var(--tasks-plan-title-padding,5px);border:none;appearance:none;cursor:pointer}:host .header select::-ms-expand{display:none}:host .header .plan-title{font-size:var(--tasks-plan-title-font-size,1.1em);padding:var(--tasks-plan-title-padding,5px)}:host .header .add-bar{display:flex}:host .header .add-bar .add-bar-item{flex:1 1 auto}:host .header .new-task-due{display:flex}:host .header .new-task-due input{flex:1 1 auto;background-color:var(--task-background-color,var(--neutral-layer-floating))}:host .header .title-cont{flex:1 1 auto;display:flex;align-items:center;height:var(--tasks-new-button-height,34px)}:host .header .new-task-button{flex:0 0 auto;display:flex;justify-content:center;align-items:center;width:var(--tasks-new-button-width,auto);height:var(--tasks-new-button-height,32px);border-radius:2px;padding:0 20px;background:var(--tasks-new-button-background,#0078d4);border:var(--tasks-new-button-border,solid 1px transparent);color:var(--tasks-new-button-color,#fff);user-select:none;cursor:pointer}:host .header .new-task-button span{font-size:14px;font-weight:600;letter-spacing:.1px;line-height:14px}:host .header .new-task-button .task-icon{margin-right:8px}:host .header .new-task-button.hidden{visibility:hidden}:host .header .new-task-button:hover{background:var(--tasks-new-button-hover-background,#106ebe)}:host .header .new-task-button:active{background:var(--tasks-new-button-active-background,#005a9e)}:host .task{position:relative;margin:var(--task-margin,0 0 0 0);padding:var(--task-padding,0 0 0 0);background-color:var(--task-background-color,var(--neutral-layer-floating));border:var(--task-border,var(--neutral-stroke-input-active));border-radius:8px}:host .task .task-content{display:flex}:host .task .task-content .divider{position:absolute;height:2px;left:0;right:0;bottom:0;background-color:var(--task-background-color,var(--neutral-layer-floating))}:host .task .task-content .task-details-container{flex:1;display:grid;grid-template-columns:auto 1fr;-ms-grid-columns:auto 1fr;grid-template-rows:auto auto auto auto;-ms-grid-rows:auto auto auto auto;justify-content:space-between;align-items:flex-start;color:var(--task-detail-color,var(--neutral-foreground-hint));font-size:12px;font-weight:600;white-space:normal;margin-bottom:12px}:host .task .task-content .task-details-container .task-detail{width:100%;height:100%;margin:4px 24px 6px 0;display:flex;justify-content:flex-start;align-items:center}:host .task .task-content .task-details-container .task-title{color:var(--task-color,var(--neutral-foreground-rest));font-size:14px;font-weight:600;grid-row:1;grid-column:1/3;grid-column:1;-ms-grid-column-span:2;margin:22px 0 4px}:host .task .task-content .task-details-container .task-group{min-height:24px;grid-row:2;grid-column:1}:host .task .task-content .task-details-container .task-bucket{min-height:24px;grid-row:2;grid-column:2}:host .task .task-content .task-details-container .task-due{justify-content:flex-end;align-items:flex-start;grid-row:4;grid-column:1/3;grid-column:1;-ms-grid-column-span:2}:host .task .task-content .task-details-container.tablet{grid-template-columns:1fr .5fr 1fr .5fr;-ms-grid-columns:1fr .5fr 1fr .5fr;grid-template-rows:auto auto;-ms-grid-rows:auto auto}:host .task .task-content .task-details-container.tablet.no-plan{grid-template-columns:0 1fr 1fr 1fr;-ms-grid-columns:0 1fr 1fr 1fr}:host .task .task-content .task-details-container.tablet .task-detail{margin:4px 24px 6px 0}:host .task .task-content .task-details-container.tablet .task-title{grid-row:1;grid-column:1/5;grid-column:1;-ms-grid-column-span:4}:host .task .task-content .task-details-container.tablet .task-group{grid-row:2;grid-column:1}:host .task .task-content .task-details-container.tablet .task-bucket{grid-row:2;grid-column:2}:host .task .task-content .task-details-container.tablet .task-assignee{grid-row:2;grid-column:3}:host .task .task-content .task-details-container.tablet .task-due{justify-content:flex-start;align-items:center;grid-row:2;grid-column:4}:host .task .task-content .task-details-container.desktop{grid-template-columns:2fr 1fr .5fr 1fr .5fr;-ms-grid-columns:2fr 1fr .5fr 1fr .5fr;grid-template-rows:auto;-ms-grid-rows:auto;margin:0}:host .task .task-content .task-details-container.desktop.no-plan{grid-template-columns:2fr 0 1fr 1fr 1fr;-ms-grid-columns:2fr 0 1fr 1fr 1fr}:host .task .task-content .task-details-container.desktop .task-detail{margin:0 24px 0 0}:host .task .task-content .task-details-container.desktop .task-title{padding:0;grid-row:1;grid-column:1}:host .task .task-content .task-details-container.desktop .task-group{min-height:61px;grid-row:1;grid-column:2}:host .task .task-content .task-details-container.desktop .task-bucket{grid-row:1;grid-column:3}:host .task .task-content .task-details-container.desktop .task-assignee{grid-row:1;grid-column:4}:host .task .task-content .task-details-container.desktop .task-due{justify-content:flex-start;align-items:center;grid-row:1;grid-column:5}:host .task .task-content .task-details-container svg{vertical-align:middle;margin-right:4px}:host .task .task-content .task-details-container svg path{fill:var(--task-detail-color,var(--neutral-foreground-hint))}:host .task .task-content .task-details-container select,:host .task .task-content .task-details-container span{vertical-align:middle;color:var(--task-detail-color,var(--neutral-foreground-hint))}:host .task .task-content .task-details-container .task-icon{color:#797775;margin-right:8px}:host .task .task-content .task-details-container .people{color:var(--task-detail-color,var(--neutral-foreground-hint));font-size:16px}:host .task .task-content .task-details-container .person{display:inline-block}:host .task .task-content .task-details-container .picker{background-color:var(--task-background-color,var(--neutral-layer-floating));background-clip:padding-box;width:var(--mgt-flyout-set-width,350px);color:var(--task-detail-color,var(--neutral-foreground-hint))}:host .task .task-content .task-details-container .picker .people-picker{--separator-margin:0px 10px 0px 10px}:host .task .task-content .task-details-container input,:host .task .task-content .task-details-container select{color:var(--task-detail-color,var(--neutral-foreground-hint));font-size:.9em;background-color:var(--task-background-color,var(--neutral-layer-floating))}:host .task .task-check-container{font-family:FabricMDL2Icons;border-radius:50%;color:#fff;cursor:pointer;display:flex;align-items:var(--task-icon-alignment,flex-start);margin:21px 10px 20px 20px;user-select:none}:host .task .task-check-container.complete .task-check{background-color:var(--task-icon-background-completed,#00ad56);border:var(--task-icon-border-completed,solid 1px #fff);color:var(--task-icon-color-completed,#fff)}:host .task .task-check-container.complete .task-content .task-details-container .task-title{text-decoration:line-through;color:var(--task-detail-color,var(--neutral-foreground-hint))}:host .task .task-check-container .task-check{font-family:FabricMDL2Icons;font-size:12px;width:18px;height:18px;border-radius:var(--task-icon-border-radius,50%);border:var(--task-icon-border,solid 1px #797775);color:var(--task-color,var(--neutral-foreground-rest));display:flex;justify-content:center;align-items:center;background-color:var(--task-icon-background,transparent);user-select:none}:host .task .task-check-container .task-check.loading .task-check-content{animation:rotate-icon 2s infinite linear}:host .task .task-options{cursor:pointer;user-select:none;margin:16px 8px 0 0}:host .task.read-only .task-check-container{cursor:default}:host .task.complete{background:var(--task-complete-background,var(--neutral-layer-1));border:var(--task-complete-border,2px dotted inherit)}:host .task.complete .task-content .task-details-container .task-title{text-decoration:line-through;color:var(--task-detail-color,var(--neutral-foreground-hint))}:host .task.new-task{margin:var(--task-new-margin,var(--task-margin,0 0 24px 0));display:flex;flex-direction:row}:host .task.new-task .self-assign{display:none}:host .task.new-task .assign-to{border:0;background:0 0}:host .task.new-task .fake-check-box{width:15px;height:15px;cursor:pointer;margin:0 5px}:host .task.new-task .fake-check-box::after{font-family:FabricMDL2Icons;content:"uE739"}:host .task.new-task .fake-check-box~:checked{font-family:FabricMDL2Icons;content:"uE73A"}:host .task.new-task .task-content{flex:1 1 auto;align-content:center;vertical-align:middle;margin:0 0 0 12px}:host .task.new-task .task-content .task-details-container{display:flex;flex-direction:column;align-items:stretch;margin:0}:host .task.new-task .task-content .task-details-container .task-title{display:flex;height:32px;padding:2px}:host .task.new-task .task-content .task-details-container .task-title input{flex:1;margin:var(--task-new-input-margin,0 24px 0 16px);padding:var(--task-new-input-padding,6px);font-size:var(--task-new-input-font-size,14px);font-weight:600;border:var(--task-new-border,none);border-bottom:1px solid #e1dfdd;outline:0;border-radius:0;background-color:var(--task-background-color,var(--neutral-layer-floating))}:host .task.new-task .task-content .task-details-container .task-title input:hover{border-bottom:1px solid #106ebe}:host .task.new-task .task-content .task-details-container .task-title input:active{border-bottom:1px solid #005a9e}:host .task.new-task .task-content .task-details-container .task-title input:focus{border-bottom:1px solid #0078d4}:host .task.new-task .task-content .task-details-container .task-details{display:flex;justify-content:stretch;align-items:center;flex-wrap:wrap;margin:14px 0 14px 4px}:host .task.new-task .task-content .task-details-container .task-details .new-task-group{margin:8px 16px}:host .task.new-task .task-content .task-details-container .task-details .new-task-bucket{margin:8px 16px}:host .task.new-task .task-content .task-details-container .task-details .new-task-due{margin:8px 16px}:host .task.new-task .task-content .task-details-container .task-details .new-task-assignee{margin:8px 16px;min-width:80px}:host .task.new-task .task-content .task-details-container .task-details .task-people label{display:flex;align-content:center;align-items:center}:host .task.new-task .task-content .task-details-container .task-details input,:host .task.new-task .task-content .task-details-container .task-details select{font-size:12px;font-weight:600;border:var(--task-new-select-border,none);border-bottom:1px solid #e1dfdd}:host .task.new-task .task-content .task-details-container .task-details input:hover,:host .task.new-task .task-content .task-details-container .task-details select:hover{border-bottom:1px solid #106ebe}:host .task.new-task .task-content .task-details-container .task-details input:active,:host .task.new-task .task-content .task-details-container .task-details select:active{border-bottom:1px solid #005a9e}:host .task.new-task .task-content .task-details-container .task-details input:focus,:host .task.new-task .task-content .task-details-container .task-details select:focus{border-bottom:1px solid #0078d4}:host .task.new-task .task-add-button-container{margin-right:28px}:host .task.new-task .task-add-button-container .task-add,:host .task.new-task .task-add-button-container .task-cancel{justify-content:center;align-items:center;cursor:pointer;flex:0 0 auto;display:flex;width:var(--tasks-new-button-width,100px);height:var(--tasks-new-button-height,32px);border-radius:4px;border:var(--tasks-new-button-border,solid 1px #e5e5e5);font-size:14px;line-height:20px}:host .task.new-task .task-add-button-container .task-add{color:#fff;background:var(--task-new-add-button-background,#0078d4);margin:22px 0 12px auto}:host .task.new-task .task-add-button-container .task-cancel{color:var(--task-new-cancel-button-color,var(--neutral-foreground-rest))}:host .task.new-task .task-add-button-container.disabled .task-add{color:var(--task-new-cancel-button-color,var(--neutral-foreground-rest));background:var(--task-new-add-button-disabled-background,#fff);cursor:default}@keyframes rotate-icon{from{transform:rotate(0)}to{transform:rotate(360deg)}}[dir=rtl] .arrow-options{--arrow-options-left:auto}[dir=rtl] .dot-options{--dot-options-translatex:translateX(60px)}[dir=rtl] .task-details{margin-right:14px!important}[dir=rtl] .task-icon{margin-left:8px}[dir=rtl] .task-detail svg{margin-left:4px}@media (forced-colors:active) and (prefers-color-scheme:dark){:host svg{fill:#fff!important;fill-rule:nonzero!important;clip-rule:nonzero!important}:host svg path{fill:#fff!important;fill-rule:nonzero!important;clip-rule:nonzero!important}}@media (forced-colors:active) and (prefers-color-scheme:light){:host svg{fill:#000!important;fill-rule:nonzero!important;clip-rule:nonzero!important}:host svg path{fill:#000!important;fill-rule:nonzero!important;clip-rule:nonzero!important}}:host{border-radius:8px;width:100%}:host .task,:host.loading-task{margin-block:1px;box-shadow:var(--task-box-shadow,var(--elevation-shadow-card-active));width:100%;background-color:var(--task-background-color,var(--neutral-layer-floating))}:host .task.new-task,:host.loading-task.new-task{margin:14px 0 1px;box-shadow:var(--task-box-shadow,var(--elevation-shadow-card-active))}:host .task.complete,:host.loading-task.complete{text-decoration:line-through;border:1px solid var(--task-border-completed,var(--neutral-stroke-input-active));background:var(--task-complete-background,var(--neutral-layer-1))}:host .task.read-only,:host.loading-task.read-only{opacity:1}:host .task:hover,:host.loading-task:hover{background-color:var(--task-background-color-hover,var(--neutral-fill-hover));border-radius:8px}:host .task .task-details,:host.loading-task .task-details{box-sizing:border-box;display:flex;flex-direction:row;align-items:center;padding:2px;line-height:24px;border-radius:4px}:host .task .task-details .task>div,:host.loading-task .task-details .task>div{display:flex;align-items:center;width:200px}:host .task .task-details .title,:host.loading-task .task-details .title{flex-grow:1}:host .task .task-details .task-delete,:host.loading-task .task-details .task-delete{display:flex}:host .task .task-details .task-due,:host.loading-task .task-details .task-due{min-width:120px;margin-inline-end:12px;height:32px;text-decoration:inherit;display:flex}:host .task .task-details .task-due .task-calendar,:host.loading-task .task-details .task-due .task-calendar{display:flex;margin-top:5px;margin-inline-end:10px}:host .task .task-details .task-due .task-calendar svg,:host.loading-task .task-details .task-due .task-calendar svg{fill:var(--task-color,var(--neutral-foreground-rest))}:host .task .task-details .task-due .task-due-date,:host.loading-task .task-details .task-due .task-due-date{display:flex;margin-top:5px}:host fluent-text-field::part(end),:host fluent-text-field::part(start){margin-inline:unset}:host fluent-text-field::part(control){padding:0;cursor:pointer}:host fluent-text-field::part(root){background:0 0}:host fluent-text-field.new-task{width:100%;height:34px}:host fluent-text-field.new-task div.start .add-icon{display:flex;margin-inline:10px}:host fluent-text-field.new-task div:nth-child(2){display:flex;align-items:center}:host fluent-text-field.new-task div:nth-child(2) .calendar{display:flex;align-items:center}:host fluent-text-field.new-task div:nth-child(2) .calendar svg{fill:var(--task-color,var(--neutral-foreground-rest))}:host fluent-text-field.new-task div:nth-child(2) .calendar .date{margin-inline-start:10px;color:var(--task-color,var(--neutral-foreground-rest));width:auto;cursor:pointer}:host fluent-text-field.new-task div:nth-child(2) .calendar .date::after{border-bottom:none}:host fluent-text-field.new-task div:nth-child(2) .calendar .date.dark::part(control){color-scheme:dark}:host fluent-text-field.new-task div:nth-child(2) .calendar input{flex:1;border:none;border-bottom:1px solid var(--task-color,var(--neutral-foreground-rest));outline:0;border-radius:0}:host fluent-text-field.new-task div:nth-child(2) .calendar input:hover{border-bottom:1px solid var(--task-date-input-hover-color,var(--neutral-layer-1))}:host fluent-text-field.new-task div:nth-child(2) .calendar input:active{border-bottom:1px solid var(--task-date-input-active-color,var(--accent-fill-rest))}:host fluent-text-field.new-task div:nth-child(2) .calendar input:focus{border-bottom:1px solid var(--task-date-input-active-color,var(--accent-fill-rest))}:host fluent-button.task-add-icon.neutral,:host fluent-button.task-cancel-icon.neutral,:host fluent-button.task-delete.neutral{fill:var(--task-color,var(--neutral-foreground-rest))}:host fluent-button.task-add-icon.neutral::part(control),:host fluent-button.task-cancel-icon.neutral::part(control),:host fluent-button.task-delete.neutral::part(control){border:none;background:inherit}:host fluent-button.task-add-icon.neutral::part(control) svg,:host fluent-button.task-cancel-icon.neutral::part(control) svg,:host fluent-button.task-delete.neutral::part(control) svg{fill:var(--task-color,var(--neutral-foreground-rest))}:host fluent-checkbox.task.complete div>svg .filled{display:block}:host fluent-checkbox.task.complete div>svg .regular{display:none}:host fluent-checkbox.task.complete div>svg path{fill:var(--task-radio-background-color,var(--accent-fill-rest))}:host fluent-checkbox.task div>svg .filled{display:none}:host fluent-checkbox.task div>svg .regular{display:block}:host fluent-checkbox.task div>svg path{fill:var(--task-background-color,var(--neutral-layer-floating))}:host fluent-checkbox::part(control){margin-inline-start:10px;background:0 0;border-radius:50%}:host fluent-checkbox::part(label){margin-inline-end:unset;width:100%}
`],Qa={cancelNewTaskSubtitle:"Cancel",newTaskPlaceholder:"Add a task",newTaskLabel:"New Task Input",addTaskButtonSubtitle:"Add",deleteTaskLabel:"Delete Task",dueDate:"Due date",newTaskDateInputLabel:"New Task Date Input",newTaskNameInputLabel:"New Task Name Input",cancelAddingTask:"Cancel adding a new task"};class Ya extends ue.a{constructor(){super(...arguments),this.orientation=an,this.radioChangeHandler=e=>{const t=e.target;t.checked&&(this.slottedRadioButtons.forEach(e=>{e!==t&&(e.checked=!1,this.isInsideFoundationToolbar||e.setAttribute("tabindex","-1"))}),this.selectedRadio=t,this.value=t.value,t.setAttribute("tabindex","0"),this.focusedRadio=t),e.stopPropagation()},this.moveToRadioByIndex=(e,t)=>{const n=e[t];this.isInsideToolbar||(n.setAttribute("tabindex","0"),n.readOnly?this.slottedRadioButtons.forEach(e=>{e!==n&&e.setAttribute("tabindex","-1")}):(n.checked=!0,this.selectedRadio=n)),this.focusedRadio=n,n.focus()},this.moveRightOffGroup=()=>{var e;null===(e=this.nextElementSibling)||void 0===e||e.focus()},this.moveLeftOffGroup=()=>{var e;null===(e=this.previousElementSibling)||void 0===e||e.focus()},this.focusOutHandler=e=>{const t=this.slottedRadioButtons,n=e.target,a=null!==n?t.indexOf(n):0,i=this.focusedRadio?t.indexOf(this.focusedRadio):-1;return(0===i&&a===i||i===t.length-1&&i===a)&&(this.selectedRadio?(this.focusedRadio=this.selectedRadio,this.isInsideFoundationToolbar||(this.selectedRadio.setAttribute("tabindex","0"),t.forEach(e=>{e!==this.selectedRadio&&e.setAttribute("tabindex","-1")}))):(this.focusedRadio=t[0],this.focusedRadio.setAttribute("tabindex","0"),t.forEach(e=>{e!==this.focusedRadio&&e.setAttribute("tabindex","-1")}))),!0},this.clickHandler=e=>{const t=e.target;if(t){const e=this.slottedRadioButtons;t.checked||0===e.indexOf(t)?(t.setAttribute("tabindex","0"),this.selectedRadio=t):(t.setAttribute("tabindex","-1"),this.selectedRadio=null),this.focusedRadio=t}e.preventDefault()},this.shouldMoveOffGroupToTheRight=(e,t,n)=>e===t.length&&this.isInsideToolbar&&n===de.d,this.shouldMoveOffGroupToTheLeft=(e,t)=>(this.focusedRadio?e.indexOf(this.focusedRadio)-1:0)<0&&this.isInsideToolbar&&t===de.c,this.checkFocusedRadio=()=>{null===this.focusedRadio||this.focusedRadio.readOnly||this.focusedRadio.checked||(this.focusedRadio.checked=!0,this.focusedRadio.setAttribute("tabindex","0"),this.focusedRadio.focus(),this.selectedRadio=this.focusedRadio)},this.moveRight=e=>{const t=this.slottedRadioButtons;let n=0;if(n=this.focusedRadio?t.indexOf(this.focusedRadio)+1:1,this.shouldMoveOffGroupToTheRight(n,t,e.key))this.moveRightOffGroup();else for(n===t.length&&(n=0);n<t.length&&t.length>1;){if(!t[n].disabled){this.moveToRadioByIndex(t,n);break}if(this.focusedRadio&&n===t.indexOf(this.focusedRadio))break;if(n+1>=t.length){if(this.isInsideToolbar)break;n=0}else n+=1}},this.moveLeft=e=>{const t=this.slottedRadioButtons;let n=0;if(n=this.focusedRadio?t.indexOf(this.focusedRadio)-1:0,n=n<0?t.length-1:n,this.shouldMoveOffGroupToTheLeft(t,e.key))this.moveLeftOffGroup();else for(;n>=0&&t.length>1;){if(!t[n].disabled){this.moveToRadioByIndex(t,n);break}if(this.focusedRadio&&n===t.indexOf(this.focusedRadio))break;n-1<0?n=t.length-1:n-=1}},this.keydownHandler=e=>{const t=e.key;if(t in de.a&&this.isInsideFoundationToolbar)return!0;switch(t){case de.g:this.checkFocusedRadio();break;case de.d:case de.b:this.direction===$t.a.ltr?this.moveRight(e):this.moveLeft(e);break;case de.c:case de.e:this.direction===$t.a.ltr?this.moveLeft(e):this.moveRight(e);break;default:return!0}}}readOnlyChanged(){void 0!==this.slottedRadioButtons&&this.slottedRadioButtons.forEach(e=>{this.readOnly?e.readOnly=!0:e.readOnly=!1})}disabledChanged(){void 0!==this.slottedRadioButtons&&this.slottedRadioButtons.forEach(e=>{this.disabled?e.disabled=!0:e.disabled=!1})}nameChanged(){this.slottedRadioButtons&&this.slottedRadioButtons.forEach(e=>{e.setAttribute("name",this.name)})}valueChanged(){this.slottedRadioButtons&&this.slottedRadioButtons.forEach(e=>{e.value===this.value&&(e.checked=!0,this.selectedRadio=e)}),this.$emit("change")}slottedRadioButtonsChanged(e,t){this.slottedRadioButtons&&this.slottedRadioButtons.length>0&&this.setupRadioButtons()}get parentToolbar(){return this.closest('[role="toolbar"]')}get isInsideToolbar(){var e;return null!==(e=this.parentToolbar)&&void 0!==e&&e}get isInsideFoundationToolbar(){var e;return!!(null===(e=this.parentToolbar)||void 0===e?void 0:e.$fastController)}connectedCallback(){super.connectedCallback(),this.direction=en(this),this.setupRadioButtons()}disconnectedCallback(){this.slottedRadioButtons.forEach(e=>{e.removeEventListener("change",this.radioChangeHandler)})}setupRadioButtons(){const e=this.slottedRadioButtons.filter(e=>e.hasAttribute("checked")),t=e?e.length:0;t>1&&(e[t-1].checked=!0);let n=!1;if(this.slottedRadioButtons.forEach(e=>{void 0!==this.name&&e.setAttribute("name",this.name),this.disabled&&(e.disabled=!0),this.readOnly&&(e.readOnly=!0),this.value&&this.value===e.value?(this.selectedRadio=e,this.focusedRadio=e,e.checked=!0,e.setAttribute("tabindex","0"),n=!0):(this.isInsideFoundationToolbar||e.setAttribute("tabindex","-1"),e.checked=!1),e.addEventListener("change",this.radioChangeHandler)}),void 0===this.value&&this.slottedRadioButtons.length>0){const e=this.slottedRadioButtons.filter(e=>e.hasAttribute("checked")),t=null!==e?e.length:0;if(t>0&&!n){const n=e[t-1];n.checked=!0,this.focusedRadio=n,n.setAttribute("tabindex","0")}else this.slottedRadioButtons[0].setAttribute("tabindex","0"),this.focusedRadio=this.slottedRadioButtons[0]}}}Object(oe.a)([Object(ce.c)({attribute:"readonly",mode:"boolean"})],Ya.prototype,"readOnly",void 0),Object(oe.a)([Object(ce.c)({attribute:"disabled",mode:"boolean"})],Ya.prototype,"disabled",void 0),Object(oe.a)([ce.c],Ya.prototype,"name",void 0),Object(oe.a)([ce.c],Ya.prototype,"value",void 0),Object(oe.a)([ce.c],Ya.prototype,"orientation",void 0),Object(oe.a)([se.d],Ya.prototype,"childItems",void 0),Object(oe.a)([se.d],Ya.prototype,"slottedRadioButtons",void 0);const Ja=Ya.compose({baseName:"radio-group",template:(e,t)=>Se.a`
    <template
        role="radiogroup"
        aria-disabled="${e=>e.disabled}"
        aria-readonly="${e=>e.readOnly}"
        @click="${(e,t)=>e.clickHandler(t.event)}"
        @keydown="${(e,t)=>e.keydownHandler(t.event)}"
        @focusout="${(e,t)=>e.focusOutHandler(t.event)}"
    >
        <slot name="label"></slot>
        <div
            class="positioning-region ${e=>e.orientation===an?"horizontal":"vertical"}"
            part="positioning-region"
        >
            <slot
                ${Object(De.a)({property:"slottedRadioButtons",filter:Object(Vn.b)("[role=radio]")})}
            ></slot>
        </div>
    </template>
`,styles:(e,t)=>Oe.a`
  ${Object(we.a)("flex")} :host {
    align-items: flex-start;
    flex-direction: column;
  }

  .positioning-region {
    display: flex;
    flex-wrap: wrap;
  }

  :host([orientation='vertical']) .positioning-region {
    flex-direction: column;
  }

  :host([orientation='horizontal']) .positioning-region {
    flex-direction: row;
  }
`});var Xa=n("R2TB");const Za={comboboxPlaceholder:"Select an item"};class $a extends ve{}class ei extends(Object(gt.b)($a)){constructor(){super(...arguments),this.proxy=document.createElement("input")}}class ti extends ei{constructor(){super(...arguments),this._value="",this.filteredOptions=[],this.filter="",this.forcedPosition=!1,this.listboxId=Object(le.a)("listbox-"),this.maxHeight=0,this.open=!1}formResetCallback(){super.formResetCallback(),this.setDefaultSelectedOption(),this.updateValue()}validate(){super.validate(this.control)}get isAutocompleteInline(){return"inline"===this.autocomplete||this.isAutocompleteBoth}get isAutocompleteList(){return"list"===this.autocomplete||this.isAutocompleteBoth}get isAutocompleteBoth(){return"both"===this.autocomplete}openChanged(){if(this.open)return this.ariaControls=this.listboxId,this.ariaExpanded="true",this.setPositioning(),this.focusAndScrollOptionIntoView(),void Ie.a.queueUpdate(()=>this.focus());this.ariaControls="",this.ariaExpanded="false"}get options(){return se.b.track(this,"options"),this.filteredOptions.length?this.filteredOptions:this._options}set options(e){this._options=e,se.b.notify(this,"options")}placeholderChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.placeholder=this.placeholder)}positionChanged(e,t){this.positionAttribute=t,this.setPositioning()}get value(){return se.b.track(this,"value"),this._value}set value(e){var t,n,a;const i=`${this._value}`;if(this.$fastController.isConnected&&this.options){const i=this.options.findIndex(t=>t.text.toLowerCase()===e.toLowerCase()),r=null===(t=this.options[this.selectedIndex])||void 0===t?void 0:t.text,o=null===(n=this.options[i])||void 0===n?void 0:n.text;this.selectedIndex=r!==o?i:this.selectedIndex,e=(null===(a=this.firstSelectedOption)||void 0===a?void 0:a.text)||e}i!==e&&(this._value=e,super.valueChanged(i,e),se.b.notify(this,"value"))}clickHandler(e){if(!this.disabled){if(this.open){const t=e.target.closest("option,[role=option]");if(!t||t.disabled)return;this.selectedOptions=[t],this.control.value=t.text,this.clearSelectionRange(),this.updateValue(!0)}return this.open=!this.open,this.open&&this.control.focus(),!0}}connectedCallback(){super.connectedCallback(),this.forcedPosition=!!this.positionAttribute,this.value&&(this.initialValue=this.value)}disabledChanged(e,t){super.disabledChanged&&super.disabledChanged(e,t),this.ariaDisabled=this.disabled?"true":"false"}filterOptions(){this.autocomplete&&"none"!==this.autocomplete||(this.filter="");const e=this.filter.toLowerCase();this.filteredOptions=this._options.filter(e=>e.text.toLowerCase().startsWith(this.filter.toLowerCase())),this.isAutocompleteList&&(this.filteredOptions.length||e||(this.filteredOptions=this._options),this._options.forEach(e=>{e.hidden=!this.filteredOptions.includes(e)}))}focusAndScrollOptionIntoView(){this.contains(document.activeElement)&&(this.control.focus(),this.firstSelectedOption&&requestAnimationFrame(()=>{var e;null===(e=this.firstSelectedOption)||void 0===e||e.scrollIntoView({block:"nearest"})}))}focusoutHandler(e){if(this.syncValue(),!this.open)return!0;const t=e.relatedTarget;this.isSameNode(t)?this.focus():this.options&&this.options.includes(t)||(this.open=!1)}inputHandler(e){if(this.filter=this.control.value,this.filterOptions(),this.isAutocompleteInline||(this.selectedIndex=this.options.map(e=>e.text).indexOf(this.control.value)),e.inputType.includes("deleteContent")||!this.filter.length)return!0;this.isAutocompleteList&&!this.open&&(this.open=!0),this.isAutocompleteInline&&(this.filteredOptions.length?(this.selectedOptions=[this.filteredOptions[0]],this.selectedIndex=this.options.indexOf(this.firstSelectedOption),this.setInlineSelection()):this.selectedIndex=-1)}keydownHandler(e){const t=e.key;if(e.ctrlKey||e.shiftKey)return!0;switch(t){case"Enter":this.syncValue(),this.isAutocompleteInline&&(this.filter=this.value),this.open=!1,this.clearSelectionRange();break;case"Escape":if(this.isAutocompleteInline||(this.selectedIndex=-1),this.open){this.open=!1;break}this.value="",this.control.value="",this.filter="",this.filterOptions();break;case"Tab":if(this.setInputToSelection(),!this.open)return!0;e.preventDefault(),this.open=!1;break;case"ArrowUp":case"ArrowDown":if(this.filterOptions(),!this.open){this.open=!0;break}this.filteredOptions.length>0&&super.keydownHandler(e),this.isAutocompleteInline&&this.setInlineSelection();break;default:return!0}}keyupHandler(e){switch(e.key){case"ArrowLeft":case"ArrowRight":case"Backspace":case"Delete":case"Home":case"End":this.filter=this.control.value,this.selectedIndex=-1,this.filterOptions()}}selectedIndexChanged(e,t){if(this.$fastController.isConnected){if((t=Object(xe.b)(-1,this.options.length-1,t))!==this.selectedIndex)return void(this.selectedIndex=t);super.selectedIndexChanged(e,t)}}selectPreviousOption(){!this.disabled&&this.selectedIndex>=0&&(this.selectedIndex=this.selectedIndex-1)}setDefaultSelectedOption(){if(this.$fastController.isConnected&&this.options){const e=this.options.findIndex(e=>null!==e.getAttribute("selected")||e.selected);this.selectedIndex=e,!this.dirtyValue&&this.firstSelectedOption&&(this.value=this.firstSelectedOption.text),this.setSelectedOptions()}}setInputToSelection(){this.firstSelectedOption&&(this.control.value=this.firstSelectedOption.text,this.control.focus())}setInlineSelection(){this.firstSelectedOption&&(this.setInputToSelection(),this.control.setSelectionRange(this.filter.length,this.control.value.length,"backward"))}syncValue(){var e;const t=this.selectedIndex>-1?null===(e=this.firstSelectedOption)||void 0===e?void 0:e.text:this.control.value;this.updateValue(this.value!==t)}setPositioning(){const e=this.getBoundingClientRect(),t=window.innerHeight-e.bottom;this.position=this.forcedPosition?this.positionAttribute:e.top>t?St:Dt,this.positionAttribute=this.forcedPosition?this.positionAttribute:this.position,this.maxHeight=this.position===St?~~e.top:~~t}selectedOptionsChanged(e,t){this.$fastController.isConnected&&this._options.forEach(e=>{e.selected=t.includes(e)})}slottedOptionsChanged(e,t){super.slottedOptionsChanged(e,t),this.updateValue()}updateValue(e){var t;this.$fastController.isConnected&&(this.value=(null===(t=this.firstSelectedOption)||void 0===t?void 0:t.text)||this.control.value,this.control.value=this.value),e&&this.$emit("change")}clearSelectionRange(){const e=this.control.value.length;this.control.setSelectionRange(e,e)}}Object(oe.a)([Object(ce.c)({attribute:"autocomplete",mode:"fromView"})],ti.prototype,"autocomplete",void 0),Object(oe.a)([se.d],ti.prototype,"maxHeight",void 0),Object(oe.a)([Object(ce.c)({attribute:"open",mode:"boolean"})],ti.prototype,"open",void 0),Object(oe.a)([ce.c],ti.prototype,"placeholder",void 0),Object(oe.a)([Object(ce.c)({attribute:"position"})],ti.prototype,"positionAttribute",void 0),Object(oe.a)([se.d],ti.prototype,"position",void 0);class ni{}Object(oe.a)([se.d],ni.prototype,"ariaAutoComplete",void 0),Object(oe.a)([se.d],ni.prototype,"ariaControls",void 0),Object(_e.a)(ni,ye),Object(_e.a)(ti,me.a,ni);const ai=".control",ii=":not([disabled]):not([open])";class ri extends ti{appearanceChanged(e,t){e!==t&&(this.classList.add(t),this.classList.remove(e))}connectedCallback(){super.connectedCallback(),this.appearance||(this.appearance="outline"),this.listbox&&Ee.v.setValueFor(this.listbox,Ee.gb)}}Object(bt.a)([Object(ce.c)({mode:"fromView"})],ri.prototype,"appearance",void 0);const oi=ri.compose({baseName:"combobox",baseClass:ti,shadowOptions:{delegatesFocus:!0},template:(e,t)=>Se.a`
    <template
        aria-disabled="${e=>e.ariaDisabled}"
        autocomplete="${e=>e.autocomplete}"
        class="${e=>e.open?"open":""} ${e=>e.disabled?"disabled":""} ${e=>e.position}"
        ?open="${e=>e.open}"
        tabindex="${e=>e.disabled?null:"0"}"
        @click="${(e,t)=>e.clickHandler(t.event)}"
        @focusout="${(e,t)=>e.focusoutHandler(t.event)}"
        @keydown="${(e,t)=>e.keydownHandler(t.event)}"
    >
        <div class="control" part="control">
            ${Object(me.d)(e,t)}
            <slot name="control">
                <input
                    aria-activedescendant="${e=>e.open?e.ariaActiveDescendant:null}"
                    aria-autocomplete="${e=>e.ariaAutoComplete}"
                    aria-controls="${e=>e.ariaControls}"
                    aria-disabled="${e=>e.ariaDisabled}"
                    aria-expanded="${e=>e.ariaExpanded}"
                    aria-haspopup="listbox"
                    class="selected-value"
                    part="selected-value"
                    placeholder="${e=>e.placeholder}"
                    role="combobox"
                    type="text"
                    ?disabled="${e=>e.disabled}"
                    :value="${e=>e.value}"
                    @input="${(e,t)=>e.inputHandler(t.event)}"
                    @keyup="${(e,t)=>e.keyupHandler(t.event)}"
                    ${Object(Ot.a)("control")}
                />
                <div class="indicator" part="indicator" aria-hidden="true">
                    <slot name="indicator">
                        ${t.indicator||""}
                    </slot>
                </div>
            </slot>
            ${Object(me.b)(e,t)}
        </div>
        <div
            class="listbox"
            id="${e=>e.listboxId}"
            part="listbox"
            role="listbox"
            ?disabled="${e=>e.disabled}"
            ?hidden="${e=>!e.open}"
            ${Object(Ot.a)("listbox")}
        >
            <slot
                ${Object(De.a)({filter:ve.slottedOptionFilter,flatten:!0,property:"slottedOptions"})}
            ></slot>
        </div>
    </template>
`,styles:(e,t)=>Oe.a`
    ${Bt()}

    ${Object(Ft.e)(e,t,ai)}

    :host(:empty) .listbox {
      display: none;
    }

    :host([disabled]) *,
    :host([disabled]) {
      cursor: ${Et.a};
      user-select: none;
    }

    :host(:active) .selected-value {
      user-select: none;
    }

    .selected-value {
      -webkit-appearance: none;
      background: transparent;
      border: none;
      color: inherit;
      ${Tt.a}
      height: calc(100% - ${Ee.vb} * 1px));
      margin: auto 0;
      width: 100%;
      outline: none;
    }
  `.withBehaviors(Object(Pt.a)("outline",Object(Ft.d)(e,t,ai,ii)),Object(Pt.a)("filled",Object(Ft.b)(e,t,ai,ii)),Object(At.a)(Object(Ft.c)(e,t,ai,ii))),indicator:'\n    <svg width="12" height="12" xmlns="http://www.w3.org/2000/svg">\n      <path d="M2.15 4.65c.2-.2.5-.2.7 0L6 7.79l3.15-3.14a.5.5 0 11.7.7l-3.5 3.5a.5.5 0 01-.7 0l-3.5-3.5a.5.5 0 010-.7z"/>\n    </svg>\n  '}),si=[u.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{--max-height:var(--picker-max-height, 380px);font-family:var(--default-font-family)}:host .picker{background-color:var(--picker-background-color,transparent)}:host fluent-combobox::part(selected-value)::placeholder{color:var(--picker-text-color,var(--input-filled-placeholder-rest))}:host fluent-combobox::part(selected-value):hover::placeholder{color:var(--picker-text-color-hover,var(--secondary-text-hover-color))}[dir=rtl] .picker{direction:rtl}
`];var ci=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},di=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};const li=()=>{Object(T.a)(oi,Kt),X(),Object(A.b)("picker",ui)};class ui extends p.a{get strings(){return Za}static get styles(){return si}constructor(){super(),this.version="v1.0",this.maxPages=3,this.scopes=[],this.cacheEnabled=!1,this.cacheInvalidationPeriod=0,this.handleComboboxKeydown=e=>{let t,n;const a=e.key,i=e.target.querySelector(".selected");i&&(t=i.getAttribute("value")),"Enter"===a&&t&&(n=this.response.filter(e=>e.id===t).pop(),this.fireCustomEvent("selectionChanged",n,!0,!1,!0))},this.placeholder=this.strings.comboboxPlaceholder,this.entityType=null,this.keyName=null,this.isRefreshing=!1}refresh(e=!1){this.isRefreshing=!0,e&&this.clearState(),this.requestStateUpdate(e)}clearState(){this.response=null,this.error=null}render(){var e;if(this.isLoadingState&&!this.response)return this.renderTemplate("loading",null);if(this.hasTemplate("error")){const e=this.error?this.error:null;return this.renderTemplate("error",{error:e},"error")}return this.hasTemplate("no-data")?this.renderTemplate("no-data",null):(null===(e=this.response)||void 0===e?void 0:e.length)>0?this.renderPicker():this.renderGet()}renderPicker(){return m.a`
      <fluent-combobox
        @keydown=${this.handleComboboxKeydown}
        current-value=${Object($.a)(this.selectedValue)}
        part="picker"
        class="picker"
        id="combobox"
        autocomplete="list"
        placeholder=${this.placeholder}>
        ${this.response.map(e=>u.c`
          <fluent-option
            value=${e.id}
            @click=${t=>this.handleClick(t,e)}
          >
            ${e[this.keyName]}
          </fluent-option>`)}
      </fluent-combobox>
     `}renderGet(){return m.a`
      <mgt-get
        class="mgt-get"
        resource=${this.resource}
        version=${this.version}
        .scopes=${this.scopes}
        max-pages=${this.maxPages}
        ?cache-enabled=${this.cacheEnabled}
        ?cache-invalidation-period=${this.cacheInvalidationPeriod}>
      </mgt-get>`}loadState(){return e=this,void 0,n=function*(){if(!this.response){const e=this.renderRoot.querySelector(".mgt-get");e?e.addEventListener("dataChange",e=>this.handleDataChange(e)):Object(Xa.a)("mgt-picker component requires a child mgt-get component. Something has gone horribly wrong.")}this.isRefreshing=!1,yield Promise.resolve()},new((t=void 0)||(t=Promise))(function(a,i){function r(e){try{s(n.next(e))}catch(e){i(e)}}function o(e){try{s(n.throw(e))}catch(e){i(e)}}function s(e){var n;e.done?a(e.value):(n=e.value,n instanceof t?n:new t(function(e){e(n)})).then(r,o)}s((n=n.apply(e,[])).next())});var e,t,n}handleDataChange(e){const t=e.detail.response.value,n=e.detail.error?e.detail.error:null;this.response=t,this.error=n}handleClick(e,t){this.fireCustomEvent("selectionChanged",t,!0,!1,!0)}}ci([Object(f.b)({attribute:"resource",type:String}),di("design:type",String)],ui.prototype,"resource",void 0),ci([Object(f.b)({attribute:"version",type:String}),di("design:type",Object)],ui.prototype,"version",void 0),ci([Object(f.b)({attribute:"max-pages",type:Number}),di("design:type",Object)],ui.prototype,"maxPages",void 0),ci([Object(f.b)({attribute:"placeholder",type:String}),di("design:type",String)],ui.prototype,"placeholder",void 0),ci([Object(f.b)({attribute:"key-name",type:String}),di("design:type",String)],ui.prototype,"keyName",void 0),ci([Object(f.b)({attribute:"entity-type",type:String}),di("design:type",String)],ui.prototype,"entityType",void 0),ci([Object(f.b)({attribute:"scopes",converter:e=>e?e.toLowerCase().split(","):null}),di("design:type",Array)],ui.prototype,"scopes",void 0),ci([Object(f.b)({attribute:"cache-enabled",type:Boolean}),di("design:type",Object)],ui.prototype,"cacheEnabled",void 0),ci([Object(f.b)({attribute:"cache-invalidation-period",type:Number}),di("design:type",Object)],ui.prototype,"cacheInvalidationPeriod",void 0),ci([Object(f.b)({attribute:"selected-value",type:String}),di("design:type",String)],ui.prototype,"selectedValue",void 0),ci([Object(f.c)(),di("design:type",Array)],ui.prototype,"response",void 0);var fi=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},pi=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)},mi=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const _i=()=>{Object(T.a)(Wt.a,Ja,Me.a),li(),Object(A.b)("todo",hi)};class hi extends za{static get styles(){return qa}get strings(){return Qa}static get requiredScopes(){return["tasks.read","tasks.readwrite"]}constructor(){super(),this._isDarkMode=!1,this.onThemeChanged=()=>{this._isDarkMode=ht(this)},this.addTask=()=>mi(this,void 0,void 0,function*(){if(!this._isNewTaskBeingAdded&&this._newTaskName){this._isNewTaskBeingAdded=!0,this.requestUpdate();try{yield this.createNewTask()}finally{this.clearNewTaskData(),this._isNewTaskBeingAdded=!1,this.requestUpdate()}}}),this.renderNewTask=()=>{const e=this._newTaskName?u.c`
        <fluent-checkbox
          class="task-add-icon"
          @click="${this.addTask}">
        </fluent-checkbox>
      `:u.c`
        <span class="add-icon">${Object(S.b)(S.a.Add)}</span>
      `,t=u.c`
      <fluent-button
        aria-label=${this.strings.cancelAddingTask}
        class="task-cancel-icon" 
        @click="${this.clearNewTaskData}"
      >
        ${Object(S.b)(S.a.Cancel)}
      </fluent-button>
    `,n={dark:this._isDarkMode,date:!0},a=u.c`
      <fluent-text-field
        autocomplete="off"
        type="date"
        id="new-taskDate-input"
        class="${Object(F.a)(n)}"
        aria-label="${this.strings.newTaskDateInputLabel}"
        .value="${this.dateToInputValue(this._newTaskDueDate)}"
        @change="${this.handleDateChange}"
      >
      </fluent-text-field>
    `,i=this.readOnly?u.d:u.c`
      <fluent-text-field
        autocomplete="off"
        appearance="outline"
        class="new-task"
        id="new-task-name-input"
        aria-label="${this.strings.newTaskLabel}"
        .value=${this._newTaskName}
        placeholder="${this.strings.newTaskPlaceholder}"
        @keydown="${this.handleKeyDown}"
        @input="${this.handleInput}"
      >
        <div slot="start" class="start">${e}</div>
        ${this._newTaskName?u.c`
              <div slot="end" class="end">
                <span class="calendar">${a}</span>
                ${t}
              </div> `:u.c``}
      </fluent-text-field>
    `;return u.c`
      ${this.currentList?u.c`
            <div dir=${this.direction} class="task new-task incomplete">
              ${i}
            </div>
        `:u.c``}  
     `},this.handleSelectionChanged=e=>{this.currentList=e.detail,this.loadTasks(this.currentList)},this.renderTaskDetails=e=>{const t={task:e,list:this.currentList};if(this.hasTemplate("task"))return this.renderTemplate("task",t,e.id);let n=null;const a=e.dueDateTime?u.c`
        <span class="task-calendar">${Object(S.b)(S.a.Calendar)}</span>
        <span class="task-due-date">${Object(Ke.f)(new Date(e.dueDateTime.dateTime))}</span>
      `:u.c``,i=this.readOnly?u.c``:u.c`
        <fluent-button class="task-delete"
          @click="${()=>this.removeTask(e.id)}"
          aria-label="${this.strings.deleteTaskLabel}"
        >
          ${Object(S.b)(S.a.Delete)}
        </fluent-button>
      `;return n=this.hasTemplate("task-details")?this.renderTemplate("task-details",t,`task-details-${e.id}`):u.c`
      <div class="task-details">
        <div class="title">${e.title}</div>
        <div class="task-due">${a}</div>
        ${i}
      </div>
      `,u.c`${n}`},this.renderTask=e=>{const t=Object(F.a)({"read-only":this.readOnly,task:!0});return u.c`
      <fluent-checkbox 
        id=${e.id}
        class=${t}
        ?disabled=${this.readOnly}
        @click="${()=>this.handleTaskCheckClick(e)}"
      >
        ${this.renderTaskDetails(e)}
      </fluent-checkbox>
    `},this.renderCompletedTask=e=>{const t=Object(F.a)({complete:!0,"read-only":this.readOnly,task:!0}),n=u.c`${Object(S.b)(S.a.CheckMark)}`;return u.c`
      <fluent-checkbox 
        id=${e.id} 
        class=${t} 
        checked 
        ?disabled=${this.readOnly} 
        @click="${()=>this.handleTaskCheckClick(e)}"
      >
        <div slot="checked-indicator">
          ${n}
        </div>
        ${this.renderTaskDetails(e)}
      </fluent-checkbox>
    `},this.loadState=()=>mi(this,void 0,void 0,function*(){const e=_.a.globalProvider;if(e&&e.state===i.c.SignedIn){if(this._isLoadingTasks=!0,!this._graph){const t=e.graph.forComponent(this);this._graph=t}if(!this.currentList&&!this.initialId){const e=yield(t=this._graph,Ga(void 0,void 0,void 0,function*(){const e=yield t.api("/me/todo/lists").header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Tasks.Read")).get();return null==e?void 0:e.value})),n=null==e?void 0:e.find(e=>"defaultList"===e.wellknownListName);n&&(yield this.loadTasks(n))}var t;this.targetId?(this.currentList=yield Wa(this._graph,this.targetId),this._tasks=yield Ka(this._graph,this.targetId)):this.initialId&&(this.currentList=yield Wa(this._graph,this.initialId),this._tasks=yield Ka(this._graph,this.initialId)),this._isLoadingTasks=!1}}),this.clearNewTaskData=()=>{this._newTaskDueDate=null,this._newTaskName=""},this.clearState=()=>{super.clearState(),this.currentList=null,this._tasks=[],this._loadingTasks=[],this._isLoadingTasks=!1},this.loadTasks=e=>mi(this,void 0,void 0,function*(){this._isLoadingTasks=!0,this.currentList=e,this._tasks=yield Ka(this._graph,e.id),this._isLoadingTasks=!1,this.requestUpdate()}),this.updateTaskStatus=(e,t)=>mi(this,void 0,void 0,function*(){this._loadingTasks=[...this._loadingTasks,e.id],this.requestUpdate(),e.status=t;const n=this.currentList.id;e=yield((e,t,n,a)=>Ga(void 0,void 0,void 0,function*(){return yield e.api(`/me/todo/lists/${t}/tasks/${n}`).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Tasks.ReadWrite")).patch(a)}))(this._graph,n,e.id,e);const a=this._tasks.findIndex(t=>t.id===e.id);this._tasks[a]=e,this._loadingTasks=this._loadingTasks.filter(t=>t!==e.id),this.requestUpdate()}),this.removeTask=e=>mi(this,void 0,void 0,function*(){this._tasks=this._tasks.filter(t=>t.id!==e),this.requestUpdate();const t=this.currentList.id;yield((e,t,n)=>Ga(void 0,void 0,void 0,function*(){yield e.api(`/me/todo/lists/${t}/tasks/${n}`).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Tasks.ReadWrite")).delete()}))(this._graph,t,e),this._tasks=this._tasks.filter(t=>t.id!==e)}),this.handleInput=e=>{"new-task-name-input"===e.target.id&&(this._newTaskName=e.target.value)},this.handleKeyDown=e=>mi(this,void 0,void 0,function*(){"Enter"===e.key&&(yield this.addTask())}),this.handleDateChange=e=>{const t=e.target.value;this._newTaskDueDate=t?new Date(t+"T17:00"):null},this._graph=null,this._newTaskDueDate=null,this._tasks=[],this._loadingTasks=[],this._isLoadingTasks=!1,this.addEventListener("selectionChanged",this.handleSelectionChanged)}connectedCallback(){super.connectedCallback(),window.addEventListener("darkmodechanged",this.onThemeChanged),this.onThemeChanged()}disconnectedCallback(){window.removeEventListener("darkmodechanged",this.onThemeChanged),super.disconnectedCallback()}renderTasks(){if(this._isLoadingTasks)return this.renderLoadingTask();let e=this._tasks;e&&this.taskFilter&&(e=e.filter(e=>this.taskFilter(e)));const t=e.filter(e=>"completed"===e.status),n=Object(D.a)(e.filter(e=>"completed"!==e.status),e=>e.id,e=>this.renderTask(e)),a=Object(D.a)(t.sort((e,t)=>new Date(e.lastModifiedDateTime).getTime()-new Date(t.lastModifiedDateTime).getTime()),e=>e.id,e=>this.renderCompletedTask(e));return u.c`
      ${n}
      ${a}
    `}renderPicker(){var e,t;return this.targetId?u.c`<p>${null===(e=this.currentList)||void 0===e?void 0:e.displayName}</p>`:m.a`
        <mgt-picker
          resource="me/todo/lists"
          scopes="tasks.read, tasks.readwrite"
          key-name="displayName"
          selected-value="${Object($.a)(null===(t=this.currentList)||void 0===t?void 0:t.displayName)}"
          placeholder="Select a task list">
        </mgt-picker>`}createNewTask(){return mi(this,void 0,void 0,function*(){const e=this.currentList.id,t={title:this._newTaskName};this._newTaskDueDate&&(t.dueDateTime={dateTime:new Date(this._newTaskDueDate).toLocaleDateString(),timeZone:"UTC"});const n=yield((e,t,n)=>Ga(void 0,void 0,void 0,function*(){return yield e.api(`/me/todo/lists/${t}/tasks`).header("Cache-Control","no-store").middlewareOptions(Object(b.a)("Tasks.ReadWrite")).post(n)}))(this._graph,e,t);this._tasks.unshift(n)})}handleTaskCheckClick(e){this.handleTaskClick(e),this.readOnly||("completed"===e.status?this.updateTaskStatus(e,"notStarted"):this.updateTaskStatus(e,"completed"))}}fi([Object(f.c)(),pi("design:type",String)],hi.prototype,"_newTaskName",void 0),fi([Object(f.c)(),pi("design:type",Object)],hi.prototype,"currentList",void 0),fi([Object(f.c)(),pi("design:type",Object)],hi.prototype,"_isDarkMode",void 0);var bi=n("Y1eq"),gi=n("dupE");const vi={termsetIdRequired:"The termsetId property or termset-id attribute is required",noTermsFound:"No terms found",comboboxPlaceholder:"Select a term",loadingMessage:"Loading..."},yi=[u.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{--max-height:var(--taxonomy-picker-list-max-height, 380px)}:host .picker{background-color:var(--taxonomy-picker-background-color,transparent)}:host fluent-combobox::part(selected-value)::placeholder{color:var(--taxonomy-picker-placeholder-color,var(--input-filled-placeholder-rest))}:host fluent-combobox::part(selected-value):hover::placeholder{color:var(--taxonomy-picker-placeholder-color-hover,var(--secondary-text-hover-color))}[dir=rtl] .picker{direction:rtl}
`];var Si=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},Di=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};const Ii=()=>{Object(T.a)(oi,Kt),Object(Ye.b)(),X(),Object(A.b)("taxonomy-picker",xi)};class xi extends p.a{get strings(){return vi}static get styles(){return yi}get defaultSelectedTermId(){return this._defaultSelectedTermId}set defaultSelectedTermId(e){e!==this._defaultSelectedTermId&&(this._defaultSelectedTermId=e,this.requestStateUpdate(!0))}get selectedTerm(){return this._selectedTerm}set selectedTerm(e){this._selectedTerm=e}constructor(){super(),this.version="beta",this.position="below",this.cacheEnabled=!1,this.cacheInvalidationPeriod=0,this.placeholder=this.strings.comboboxPlaceholder,this.isRefreshing=!1,this.noTerms=!1}refresh(e=!1){this.isRefreshing=!0,e&&this.clearState(),this.requestStateUpdate(e)}clearState(){this.terms=null,this.error=null,this.noTerms=!1}render(){var e;return this.isLoadingState&&!this.terms?this.renderLoading():this.error?this.renderError():this.noTerms?this.renderNoData():(null===(e=this.terms)||void 0===e?void 0:e.length)>0?this.renderTaxonomyPicker():this.renderGet()}renderLoading(){return this.renderTemplate("loading",null,"loading")||m.a`
        <div class="message-parent">
          <mgt-spinner></mgt-spinner>
          <div label="loading-text" aria-label="loading">
            ${this.strings.loadingMessage}
          </div>
        </div>
      `}renderError(){return this.renderTemplate("error",null,"error")||u.c`
              <span>
                ${this.error}
            </span>
          `}renderNoData(){return this.renderTemplate("no-data",null)||u.c`
            <span>
              ${this.strings.noTermsFound}
            </span>
          `}renderTaxonomyPicker(){return m.a`
      <fluent-combobox class="taxonomy-picker" autocomplete="both" placeholder=${this.placeholder} position=${this.position} ?disabled=${this.disabled}>
        ${this.terms.map(e=>this.renderTaxonomyPickerItem(e))}
      </fluent-combobox>
     `}renderTaxonomyPickerItem(e){const t=this.defaultSelectedTermId&&this.defaultSelectedTermId===e.id;return u.c`
        <fluent-option value=${e.id} ?selected=${t} @click=${t=>this.handleClick(t,e)}> ${this.renderTemplate("term",{term:e},e.id)||e.labels[0].name} </fluent-option>
        `}renderGet(){if(!this.termsetId)return u.c`
            <span>
                ${this.strings.termsetIdRequired}
            </span>
            `;let e=`/termStore/sets/${this.termsetId}/children`;return this.termId&&(e=`/termStore/sets/${this.termsetId}/terms/${this.termId}/children`),this.siteId&&(e=`/sites/${this.siteId}${e}`),e+="?$select=id,labels,descriptions,properties",m.a`
      <mgt-get
        class="mgt-get"
        resource=${e}
        version=${this.version}
        scopes=${["TermStore.Read.All"]}
        ?cache-enabled=${this.cacheEnabled}
        ?cache-invalidation-period=${this.cacheInvalidationPeriod}>
      </mgt-get>`}loadState(){return e=this,void 0,n=function*(){this.terms||this.renderRoot.querySelector(".mgt-get").addEventListener("dataChange",e=>this.handleDataChange(e)),this.isRefreshing=!1,yield Promise.resolve()},new((t=void 0)||(t=Promise))(function(a,i){function r(e){try{s(n.next(e))}catch(e){i(e)}}function o(e){try{s(n.throw(e))}catch(e){i(e)}}function s(e){var n;e.done?a(e.value):(n=e.value,n instanceof t?n:new t(function(e){e(n)})).then(r,o)}s((n=n.apply(e,[])).next())});var e,t,n}handleDataChange(e){const t=e.detail.error?e.detail.error:null;if(t)return void(this.error=t);this.locale&&(this.locale=this.locale.toLowerCase());const n=e.detail.response.value.map(e=>{const t=e.labels;if(t&&t.length>0&&this.locale){const n=t.find(e=>e.languageTag.toLowerCase()===this.locale);n&&(e.labels=[n,...t.filter(e=>e.languageTag.toLowerCase()!==this.locale)])}return e});this.terms=n,0===n.length&&(this.noTerms=!0)}handleClick(e,t){this.selectedTerm=t,this.fireCustomEvent("selectionChanged",t)}}Si([Object(f.b)({attribute:"term-set-id",type:String}),Di("design:type",String)],xi.prototype,"termsetId",void 0),Si([Object(f.b)({attribute:"term-id",type:String}),Di("design:type",String)],xi.prototype,"termId",void 0),Si([Object(f.b)({attribute:"site-id",type:String}),Di("design:type",String)],xi.prototype,"siteId",void 0),Si([Object(f.b)({attribute:"locale",type:String}),Di("design:type",String)],xi.prototype,"locale",void 0),Si([Object(f.b)({attribute:"version",type:String}),Di("design:type",Object)],xi.prototype,"version",void 0),Si([Object(f.b)({attribute:"placeholder",type:String}),Di("design:type",String)],xi.prototype,"placeholder",void 0),Si([Object(f.b)({attribute:"position",type:String,converter:e=>"above"===e?"above":"below"}),Di("design:type",String)],xi.prototype,"position",void 0),Si([Object(f.b)({attribute:"default-selected-term-id",type:String}),Di("design:type",String),Di("design:paramtypes",[String])],xi.prototype,"defaultSelectedTermId",null),Si([Object(f.b)({attribute:"selected-term",type:Object}),Di("design:type",Object),Di("design:paramtypes",[Object])],xi.prototype,"selectedTerm",null),Si([Object(f.b)({attribute:"disabled",type:Boolean}),Di("design:type",Boolean)],xi.prototype,"disabled",void 0),Si([Object(f.b)({attribute:"cache-enabled",type:Boolean}),Di("design:type",Object)],xi.prototype,"cacheEnabled",void 0),Si([Object(f.b)({attribute:"cache-invalidation-period",type:Number}),Di("design:type",Object)],xi.prototype,"cacheInvalidationPeriod",void 0),Si([Object(f.c)(),Di("design:type",Array)],xi.prototype,"terms",void 0),Si([Object(f.c)(),Di("design:type",Boolean)],xi.prototype,"noTerms",void 0);class Ci extends ue.a{}class Oi extends(Object(gt.a)(Ci)){constructor(){super(...arguments),this.proxy=document.createElement("input")}}class wi extends Oi{constructor(){super(),this.initialValue="on",this.keypressHandler=e=>{if(!this.readOnly)switch(e.key){case de.g:case de.m:this.checked=!this.checked}},this.clickHandler=e=>{this.disabled||this.readOnly||(this.checked=!this.checked)},this.proxy.setAttribute("type","checkbox")}readOnlyChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.readOnly=this.readOnly),this.readOnly?this.classList.add("readonly"):this.classList.remove("readonly")}checkedChanged(e,t){super.checkedChanged(e,t),this.checked?this.classList.add("checked"):this.classList.remove("checked")}}Object(oe.a)([Object(ce.c)({attribute:"readonly",mode:"boolean"})],wi.prototype,"readOnly",void 0),Object(oe.a)([se.d],wi.prototype,"defaultSlottedNodes",void 0);const Ei=wi.compose({baseName:"switch",template:(e,t)=>Se.a`
    <template
        role="switch"
        aria-checked="${e=>e.checked}"
        aria-disabled="${e=>e.disabled}"
        aria-readonly="${e=>e.readOnly}"
        tabindex="${e=>e.disabled?null:0}"
        @keypress="${(e,t)=>e.keypressHandler(t.event)}"
        @click="${(e,t)=>e.clickHandler(t.event)}"
        class="${e=>e.checked?"checked":""}"
    >
        <label
            part="label"
            class="${e=>e.defaultSlottedNodes&&e.defaultSlottedNodes.length?"label":"label label__hidden"}"
        >
            <slot ${Object(De.a)("defaultSlottedNodes")}></slot>
        </label>
        <div part="switch" class="switch">
            <slot name="switch">${t.switch||""}</slot>
        </div>
        <span class="status-message" part="status-message">
            <span class="checked-message" part="checked-message">
                <slot name="checked-message"></slot>
            </span>
            <span class="unchecked-message" part="unchecked-message">
                <slot name="unchecked-message"></slot>
            </span>
        </span>
    </template>
`,styles:(e,t)=>Oe.a`
    :host([hidden]) {
      display: none;
    }

    ${Object(we.a)("inline-flex")} :host {
      align-items: center;
      outline: none;
      font-family: ${Ee.p};
      ${""} user-select: none;
    }

    :host(.disabled) {
      opacity: ${Ee.u};
    }

    :host(.disabled) .label,
    :host(.readonly) .label,
    :host(.disabled) .switch,
    :host(.readonly) .switch,
    :host(.disabled) .status-message,
    :host(.readonly) .status-message {
      cursor: ${Et.a};
    }

    .switch {
      position: relative;
      box-sizing: border-box;
      width: calc(((${Mt.a} / 2) + ${Ee.s}) * 2px);
      height: calc(((${Mt.a} / 2) + ${Ee.s}) * 1px);
      background: ${Ee.K};
      border-radius: calc(${Mt.a} * 1px);
      border: calc(${Ee.vb} * 1px) solid ${Ee.ub};
      cursor: pointer;
    }

    :host(:not(.disabled):hover) .switch {
      background: ${Ee.J};
      border-color: ${Ee.tb};
    }

    :host(:not(.disabled):active) .switch {
      background: ${Ee.H};
      border-color: ${Ee.sb};
    }

    :host(:${wt.a}) .switch {
      ${Ae.b}
      background: ${Ee.I};
    }

    :host(.checked) .switch {
      background: ${Ee.e};
      border-color: transparent;
    }

    :host(.checked:not(.disabled):hover) .switch {
      background: ${Ee.d};
      border-color: transparent;
    }

    :host(.checked:not(.disabled):active) .switch {
      background: ${Ee.b};
      border-color: transparent;
    }

    slot[name='switch'] {
      position: absolute;
      display: flex;
      border: 1px solid transparent; /* Spacing included in the transform reference box */
      fill: ${Ee.fb};
      transition: all 0.2s ease-in-out;
    }

    .status-message {
      color: ${Ee.fb};
      cursor: pointer;
      ${Tt.a}
    }

    .label__hidden {
      display: none;
      visibility: hidden;
    }

    .label {
      color: ${Ee.fb};
      ${Tt.a}
      margin-inline-end: calc(${Ee.s} * 2px + 2px);
      cursor: pointer;
    }

    ::slotted([slot="checked-message"]),
    ::slotted([slot="unchecked-message"]) {
        margin-inline-start: calc(${Ee.s} * 2px + 2px);
    }

    :host(.checked) .switch {
      background: ${Ee.e};
    }

    :host(.checked) .switch slot[name='switch'] {
      fill: ${Ee.C};
      filter: drop-shadow(0px 1px 1px rgba(0, 0, 0, 0.15));
    }

    :host(.checked:not(.disabled)) .switch:hover {
      background: ${Ee.d};
    }

    :host(.checked:not(.disabled)) .switch:hover slot[name='switch'] {
      fill: ${Ee.B};
    }

    :host(.checked:not(.disabled)) .switch:active {
      background: ${Ee.b};
    }

    :host(.checked:not(.disabled)) .switch:active slot[name='switch'] {
      fill: ${Ee.z};
    }

    .unchecked-message {
      display: block;
    }

    .checked-message {
      display: none;
    }

    :host(.checked) .unchecked-message {
      display: none;
    }

    :host(.checked) .checked-message {
      display: block;
    }
  `.withBehaviors(new zt(Oe.a`
        slot[name='switch'] {
          left: 0;
        }

        :host(.checked) slot[name='switch'] {
          left: 100%;
          transform: translateX(-100%);
        }
      `,Oe.a`
        slot[name='switch'] {
          right: 0;
        }

        :host(.checked) slot[name='switch'] {
          right: 100%;
          transform: translateX(100%);
        }
      `),Object(At.a)(Oe.a`
        :host(:not(.disabled)) .switch slot[name='switch'] {
          forced-color-adjust: none;
          fill: ${Lt.a.FieldText};
        }
        .switch {
          background: ${Lt.a.Field};
          border-color: ${Lt.a.FieldText};
        }
        :host(.checked) .switch {
          background: ${Lt.a.Highlight};
          border-color: ${Lt.a.Highlight};
        }
        :host(:not(.disabled):hover) .switch ,
        :host(:not(.disabled):active) .switch,
        :host(.checked:not(.disabled):hover) .switch {
          background: ${Lt.a.HighlightText};
          border-color: ${Lt.a.Highlight};
        }
        :host(.checked:not(.disabled)) .switch slot[name="switch"] {
          fill: ${Lt.a.HighlightText};
        }
        :host(.checked:not(.disabled):hover) .switch slot[name='switch'] {
          fill: ${Lt.a.Highlight};
        }
        :host(:${wt.a}) .switch {
          forced-color-adjust: none;
          background: ${Lt.a.Field}; 
          border-color: ${Lt.a.Highlight};
          outline-color: ${Lt.a.FieldText};
        }
        :host(.disabled) {
          opacity: 1;
        }
        :host(.disabled) slot[name='switch'] {
          forced-color-adjust: none;
          fill: ${Lt.a.GrayText};
        }
        :host(.disabled) .switch {
          background: ${Lt.a.Field};
          border-color: ${Lt.a.GrayText};
        }
        .status-message,
        .label {
          color: ${Lt.a.FieldText};
        }
      `)),switch:'\n    <svg width="16" height="16" viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg">\n      <rect x="2" y="2" width="12" height="12" rx="6"/>\n    </svg>\n  '});var Ai=n("yQ0v"),Li=n("YBl9");const ki="#717171",Mi=Xn.a.create("secondary-text-color").withDefault(ki),Pi="#1a1a1a",Ti=Xn.a.create("secondary-text-hover-color").withDefault(Pi),Ui=(e,t=document.body)=>{const n=Hi(e);Fi(n,t)},Fi=(e,t=document.body)=>{var n;Ee.a.setValueFor(t,_t.a.from(Object(Li.a)(e.accentBaseColor))),Ee.E.setValueFor(t,_t.a.from(Object(Li.a)(e.neutralBaseColor))),Ee.o.setValueFor(t,e.baseLayerLuminance),null===(n=e.designTokenOverrides)||void 0===n||n.call(e,t)},Hi=e=>{switch(e){case"contrast":return{accentBaseColor:"#7f85f5",neutralBaseColor:"#adadad",baseLayerLuminance:Ai.a.DarkMode};case"default":return{accentBaseColor:"#5b5fc7",neutralBaseColor:"#616161",baseLayerLuminance:Ai.a.LightMode};case"dark":return{accentBaseColor:"#479ef5",neutralBaseColor:"#adadad",baseLayerLuminance:Ai.a.DarkMode,designTokenOverrides:e=>{Ee.e.setValueFor(e,_t.a.from(Object(Li.a)("#115ea3"))),Ee.d.setValueFor(e,_t.a.from(Object(Li.a)("#0f6cbd"))),Ee.b.setValueFor(e,_t.a.from(Object(Li.a)("#0c3b5e"))),Ee.c.setValueFor(e,_t.a.from(Object(Li.a)("#0f548c"))),Ee.i.setValueFor(e,_t.a.from(Object(Li.a)("#479EF5"))),Ee.h.setValueFor(e,_t.a.from(Object(Li.a)("#62abf5"))),Ee.f.setValueFor(e,_t.a.from(Object(Li.a)("#2886de"))),Ee.g.setValueFor(e,_t.a.from(Object(Li.a)("#479ef5"))),Ee.m.setValueFor(e,_t.a.from(Object(Li.a)("#115ea3"))),Ee.l.setValueFor(e,_t.a.from(Object(Li.a)("#0f6cbd"))),Ee.j.setValueFor(e,_t.a.from(Object(Li.a)("#0c3b5e"))),Ee.k.setValueFor(e,_t.a.from(Object(Li.a)("#0f548c"))),Ee.z.setValueFor(e,_t.a.from(Object(Li.a)("#ffffff"))),Ee.C.setValueFor(e,_t.a.from(Object(Li.a)("#ffffff"))),Ee.B.setValueFor(e,_t.a.from(Object(Li.a)("#ffffff"))),Ee.A.setValueFor(e,_t.a.from(Object(Li.a)("#ffffff"))),Mi.setValueFor(e,"#8e8e8e"),Ti.setValueFor(e,"#ffffff")}};default:return{accentBaseColor:"#0f6cbd",neutralBaseColor:"#616161",baseLayerLuminance:Ai.a.LightMode,designTokenOverrides:e=>{Mi.setValueFor(e,ki),Ti.setValueFor(e,Pi)}}}},Ri={label:"Color mode:",on:"Dark",off:"Light"},Ni=()=>{Object(T.a)(Ei),Object(A.b)("theme-toggle",Bi)};class Bi extends hn.b{constructor(){super(),this.onSwitchChanged=e=>{this.darkModeActive=e.target.checked};const e=window.matchMedia("(prefers-color-scheme:dark)").matches;this.darkModeActive=e,this.applyTheme(this.darkModeActive)}get strings(){return Ri}updated(e){e.has("darkModeActive")&&this.applyTheme(this.darkModeActive)}render(){return u.c`
      <fluent-switch checked=${this.darkModeActive} @change=${this.onSwitchChanged}>
        <span slot="checked-message">${Ri.on}</span>
        <span slot="unchecked-message">${Ri.off}</span>
        <label for="direction-switch">${Ri.label}</label>
      </fluent-switch>
`}applyTheme(e){const t=e?"dark":"light";Ui(t),document.body.classList.remove("mgt-dark-mode","mgt-light-mode"),document.body.classList.add(`mgt-${t}-mode`),this.fireCustomEvent("darkmodechanged",this.darkModeActive,!0,!1,!0)}}!function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);r>3&&o&&Object.defineProperty(t,n,o)}([Object(f.b)({attribute:"mode",reflect:!0,type:String,converter:{fromAttribute:e=>"dark"===e,toAttribute:e=>e?"dark":"light"}}),function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata("design:type",t)}(0,Boolean)],Bi.prototype,"darkModeActive",void 0);class ji extends ue.a{}class Vi extends(Object(gt.b)(ji)){constructor(){super(...arguments),this.proxy=document.createElement("input")}}class zi extends Vi{readOnlyChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.readOnly=this.readOnly,this.validate())}autofocusChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.autofocus=this.autofocus,this.validate())}placeholderChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.placeholder=this.placeholder)}listChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.setAttribute("list",this.list),this.validate())}maxlengthChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.maxLength=this.maxlength,this.validate())}minlengthChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.minLength=this.minlength,this.validate())}patternChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.pattern=this.pattern,this.validate())}sizeChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.size=this.size)}spellcheckChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.spellcheck=this.spellcheck)}connectedCallback(){super.connectedCallback(),this.validate(),this.autofocus&&Ie.a.queueUpdate(()=>{this.focus()})}validate(){super.validate(this.control)}handleTextInput(){this.value=this.control.value}handleClearInput(){this.value="",this.control.focus(),this.handleChange()}handleChange(){this.$emit("change")}}Object(oe.a)([Object(ce.c)({attribute:"readonly",mode:"boolean"})],zi.prototype,"readOnly",void 0),Object(oe.a)([Object(ce.c)({mode:"boolean"})],zi.prototype,"autofocus",void 0),Object(oe.a)([ce.c],zi.prototype,"placeholder",void 0),Object(oe.a)([ce.c],zi.prototype,"list",void 0),Object(oe.a)([Object(ce.c)({converter:ce.e})],zi.prototype,"maxlength",void 0),Object(oe.a)([Object(ce.c)({converter:ce.e})],zi.prototype,"minlength",void 0),Object(oe.a)([ce.c],zi.prototype,"pattern",void 0),Object(oe.a)([Object(ce.c)({converter:ce.e})],zi.prototype,"size",void 0),Object(oe.a)([Object(ce.c)({mode:"boolean"})],zi.prototype,"spellcheck",void 0),Object(oe.a)([se.d],zi.prototype,"defaultSlottedNodes",void 0);class Gi{}Object(_e.a)(Gi,pe.a),Object(_e.a)(zi,me.a,Gi);var Ki=n("6gCt"),Wi=n("Mjrb");const qi=".root",Qi=Xn.a.create("clear-button-hover").withDefault(e=>{const t=Ee.Z.getValueFor(e),n=Ee.N.getValueFor(e);return t.evaluate(e,n.evaluate(e).focus).hover}),Yi=Xn.a.create("clear-button-active").withDefault(e=>{const t=Ee.Z.getValueFor(e),n=Ee.N.getValueFor(e);return t.evaluate(e,n.evaluate(e).focus).active});class Ji extends zi{constructor(){super(...arguments),this.appearance="outline"}}Object(bt.a)([ce.c],Ji.prototype,"appearance",void 0);const Xi=Ji.compose({baseName:"search",baseClass:zi,template:(e,t)=>Se.a`
  <template
    class="
            ${e=>e.readOnly?"readonly":""}
        "
  >
    <label
      part="label"
      for="control"
      class="${e=>e.defaultSlottedNodes&&e.defaultSlottedNodes.length?"label":"label label__hidden"}"
    >
      <slot ${Object(De.a)({property:"defaultSlottedNodes",filter:Ki.a})}></slot>
    </label>
    <div class="root" part="root" ${Object(Ot.a)("root")}>
      ${Object(me.d)(e,t)}
      <div class="input-wrapper" part="input-wrapper">
        <input
          class="control"
          part="control"
          id="control"
          @input="${e=>e.handleTextInput()}"
          @change="${e=>e.handleChange()}"
          ?autofocus="${e=>e.autofocus}"
          ?disabled="${e=>e.disabled}"
          list="${e=>e.list}"
          maxlength="${e=>e.maxlength}"
          minlength="${e=>e.minlength}"
          pattern="${e=>e.pattern}"
          placeholder="${e=>e.placeholder}"
          ?readonly="${e=>e.readOnly}"
          ?required="${e=>e.required}"
          size="${e=>e.size}"
          ?spellcheck="${e=>e.spellcheck}"
          :value="${e=>e.value}"
          type="search"
          aria-atomic="${e=>e.ariaAtomic}"
          aria-busy="${e=>e.ariaBusy}"
          aria-controls="${e=>e.ariaControls}"
          aria-current="${e=>e.ariaCurrent}"
          aria-describedby="${e=>e.ariaDescribedby}"
          aria-details="${e=>e.ariaDetails}"
          aria-disabled="${e=>e.ariaDisabled}"
          aria-errormessage="${e=>e.ariaErrormessage}"
          aria-flowto="${e=>e.ariaFlowto}"
          aria-haspopup="${e=>e.ariaHaspopup}"
          aria-hidden="${e=>e.ariaHidden}"
          aria-invalid="${e=>e.ariaInvalid}"
          aria-keyshortcuts="${e=>e.ariaKeyshortcuts}"
          aria-label="${e=>e.ariaLabel}"
          aria-labelledby="${e=>e.ariaLabelledby}"
          aria-live="${e=>e.ariaLive}"
          aria-owns="${e=>e.ariaOwns}"
          aria-relevant="${e=>e.ariaRelevant}"
          aria-roledescription="${e=>e.ariaRoledescription}"
          ${Object(Ot.a)("control")}
        />
        <slot name="clear-button">
          <button
            class="clear-button ${e=>e.value?"":"clear-button__hidden"}"
            part="clear-button"
            tabindex="-1"
            @click=${e=>e.handleClearInput()}
          >
            <slot name="clear-glyph">
              <svg width="12" height="12" viewBox="0 0 12 12" xmlns="http://www.w3.org/2000/svg">
                <path
                  d="m2.09 2.22.06-.07a.5.5 0 0 1 .63-.06l.07.06L6 5.29l3.15-3.14a.5.5 0 1 1 .7.7L6.71 6l3.14 3.15c.18.17.2.44.06.63l-.06.07a.5.5 0 0 1-.63.06l-.07-.06L6 6.71 2.85 9.85a.5.5 0 0 1-.7-.7L5.29 6 2.15 2.85a.5.5 0 0 1-.06-.63l.06-.07-.06.07Z"
                />
              </svg>
            </slot>
          </button>
        </slot>
      </div>
      ${Object(me.b)(e,t)}
    </div>
  </template>
`,styles:(e,t)=>Oe.a`
    ${Object(we.a)("inline-block")}

    ${Object(Ft.a)(e,t,qi)}

    ${Object(Ft.e)(e,t,qi)}

    .root {
      display: flex;
      flex-direction: row;
    }
    .control {
      -webkit-appearance: none;
      color: inherit;
      background: transparent;
      border: 0;
      height: calc(100% - 4px);
      margin-top: auto;
      margin-bottom: auto;
      padding: 0 calc(${Ee.s} * 2px + 1px);
      font-family: inherit;
      font-size: inherit;
      line-height: inherit;
    }
    .clear-button {
      display: inline-flex;
      align-items: center;
      margin: 1px;
      height: calc(100% - 2px);
      opacity: 0;
      background: transparent;
      color: ${Ee.fb};
      fill: currentcolor;
      border: none;
      border-radius: calc(${Ee.q} * 1px);
      min-width: calc(${Mt.a} * 1px);
      ${Tt.a}
      outline: none;
      padding: 0 calc((10 + (${Ee.s} * 2 * ${Ee.r})) * 1px);
    }
    .clear-button:hover {
      background: ${Qi};
    }
    .clear-button:active {
      background: ${Yi};
    }
    :host(:hover:not([disabled], [readOnly])) .clear-button,
    :host(:active:not([disabled], [readOnly])) .clear-button,
    :host(:focus-within:not([disabled], [readOnly])) .clear-button {
        opacity: 1;
    }
    :host(:hover:not([disabled], [readOnly])) .clear-button__hidden,
    :host(:active:not([disabled], [readOnly])) .clear-button__hidden,
    :host(:focus-within:not([disabled], [readOnly])) .clear-button__hidden {
        opacity: 0;
    }
    .control::-webkit-search-cancel-button {
      -webkit-appearance: none;
    }
    .input-wrapper {
      display: flex;
      position: relative;
      width: 100%;
    }
    .start,
    .end {
      display: flex;
      margin: 1px;
      align-items: center;
    }
    .start {
      display: flex;
      margin-inline-start: 11px;
    }
    ::slotted([slot="end"]) {
      height: 100%
    }
    .clear-button__hidden {
      opacity: 0;
    }
    .end {
        margin-inline-end: 11px;
    }
    ::slotted(${e.tagFor(Wi.a)}) {
      margin-inline-end: 1px;
    }
  `.withBehaviors(Object(Pt.a)("outline",Object(Ft.d)(e,t,qi)),Object(Pt.a)("filled",Object(Ft.b)(e,t,qi)),Object(At.a)(Object(Ft.c)(e,t,qi))),start:'<svg width="20" height="20" xmlns="http://www.w3.org/2000/svg%22%3E"><path d="M8.5 3a5.5 5.5 0 0 1 4.23 9.02l4.12 4.13a.5.5 0 0 1-.63.76l-.07-.06-4.13-4.12A5.5 5.5 0 1 1 8.5 3Zm0 1a4.5 4.5 0 1 0 0 9 4.5 4.5 0 0 0 0-9Z"/></svg>',shadowOptions:{delegatesFocus:!0}}),Zi={placeholder:"Search",title:"Search"},$i=[u.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host fluent-search{width:100%}
`];var er=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},tr=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};const nr=()=>{Object(T.a)(Xi),Object(A.b)("search-box",ar)};class ar extends hn.b{static get styles(){return $i}get strings(){return Zi}get searchTerm(){return this._searchTerm}set searchTerm(e){this._searchTerm=e,this.fireSearchTermChanged()}constructor(){super(),this._searchTerm="",this.onInputChanged=e=>{this.searchTerm=e.target.value},Object(Xa.c)("<mgt-search-box> is a preview component and may change prior to becoming generally available. See more information https://aka.ms/mgt/preview-components"),this.debounceDelay=300}render(){var e;return u.c`
      <fluent-search
        autocomplete="off"
        class="search-term-input"
        appearance="outline"
        value=${null!==(e=this.searchTerm)&&void 0!==e?e:""}
        placeholder=${this.placeholder?this.placeholder:Zi.placeholder}
        title=${this.title?this.title:Zi.title}
        @input=${this.onInputChanged}
        @change=${this.onInputChanged}
      >
      </fluent-search>`}fireSearchTermChanged(){this.debouncedSearchTermChanged||(this.debouncedSearchTermChanged=Object(Ke.b)(()=>{this.fireCustomEvent("searchTermChanged",this.searchTerm)},this.debounceDelay)),this.debouncedSearchTermChanged()}}er([Object(f.b)({attribute:"placeholder",type:String}),tr("design:type",String)],ar.prototype,"placeholder",void 0),er([Object(f.b)({attribute:"search-term",type:String}),tr("design:type",Object),tr("design:paramtypes",[Object])],ar.prototype,"searchTerm",null),er([Object(f.b)({attribute:"debounce-delay",type:Number,reflect:!0}),tr("design:type",Number)],ar.prototype,"debounceDelay",void 0);const ir={modified:"modified on",back:"Back",next:"Next",pages:"pages",page:"Page"},rr=[u.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}.search-results{scroll-margin:100px}.search-result-grid{font-size:14px;margin:16px 4px;display:grid;grid-template-columns:32px 2fr 0fr;gap:.5rem}.search-result{font-size:14px;margin:16px 4px}.file-icon{--file-type-icon-height:32px}.search-result-info{margin:4px 0;display:inline-flex}.search-result-date{padding-top:3px}.search-result-date__shimmer{border-radius:4px;margin-top:2%;margin-left:5px;height:10px;width:200px}.search-result-summary{margin:4px 0;color:var(--body-subtle-text-color)}.search-result-thumbnail__shimmer{width:126px;height:72px}.search-result-thumbnail img{height:72px;max-width:126px;width:126px;object-fit:cover}.search-result-url{font-size:14px;font-weight:600;margin:4px 0}.search-result-url a{color:var(--color);text-decoration:none}.search-result-url a:hover{text-decoration:underline}.search-result-content__shimmer{border-radius:4px;margin-top:10px;height:10px}.search-result-name{font-size:16px;font-weight:600;margin:4px 0}.search-result-name__shimmer{border-radius:4px;margin-top:10px;height:10px;width:20%}.search-result-name a{color:var(--color);text-decoration:none}.search-result-name a:hover{text-decoration:underline}.search-result-author__shimmer{width:24px;height:24px}.search-result-icon{width:30px;height:30px}.search-result-icon__shimmer{width:32px;height:32px}.search-result-icon img{width:100%}.search-result-icon svg,.search-result-icon svg>path{width:100%;height:100%;fill:var(--neutral-foreground-rest);fill-rule:nonzero!important;clip-rule:nonzero!important}.search-results-page-active{border-bottom-style:solid;border-bottom-color:var(--accent-base-color)}.search-results-page svg,.search-results-page svg>path{height:12px;fill:var(--neutral-foreground-rest);fill-rule:nonzero!important;clip-rule:nonzero!important}.search-result-answer{box-shadow:var(--answer-box-shadow,0 3.2px 7.2px rgba(0,0,0,.132),0 .6px 1.8px rgba(0,0,0,.108));border-radius:var(--answer-border-radius,4px);border:var(--answer-border,none);padding:var(--answer-padding,10px)}.search-results-pages{margin:16px 4px}
`];class or extends ue.a{constructor(){super(...arguments),this.anchor="",this.delay=300,this.autoUpdateMode="anchor",this.anchorElement=null,this.viewportElement=null,this.verticalPositioningMode="dynamic",this.horizontalPositioningMode="dynamic",this.horizontalInset="false",this.verticalInset="false",this.horizontalScaling="content",this.verticalScaling="content",this.verticalDefaultPosition=void 0,this.horizontalDefaultPosition=void 0,this.tooltipVisible=!1,this.currentDirection=$t.a.ltr,this.showDelayTimer=null,this.hideDelayTimer=null,this.isAnchorHoveredFocused=!1,this.isRegionHovered=!1,this.handlePositionChange=e=>{this.classList.toggle("top","start"===this.region.verticalPosition),this.classList.toggle("bottom","end"===this.region.verticalPosition),this.classList.toggle("inset-top","insetStart"===this.region.verticalPosition),this.classList.toggle("inset-bottom","insetEnd"===this.region.verticalPosition),this.classList.toggle("center-vertical","center"===this.region.verticalPosition),this.classList.toggle("left","start"===this.region.horizontalPosition),this.classList.toggle("right","end"===this.region.horizontalPosition),this.classList.toggle("inset-left","insetStart"===this.region.horizontalPosition),this.classList.toggle("inset-right","insetEnd"===this.region.horizontalPosition),this.classList.toggle("center-horizontal","center"===this.region.horizontalPosition)},this.handleRegionMouseOver=e=>{this.isRegionHovered=!0},this.handleRegionMouseOut=e=>{this.isRegionHovered=!1,this.startHideDelayTimer()},this.handleAnchorMouseOver=e=>{this.tooltipVisible?this.isAnchorHoveredFocused=!0:this.startShowDelayTimer()},this.handleAnchorMouseOut=e=>{this.isAnchorHoveredFocused=!1,this.clearShowDelayTimer(),this.startHideDelayTimer()},this.handleAnchorFocusIn=e=>{this.startShowDelayTimer()},this.handleAnchorFocusOut=e=>{this.isAnchorHoveredFocused=!1,this.clearShowDelayTimer(),this.startHideDelayTimer()},this.startHideDelayTimer=()=>{this.clearHideDelayTimer(),this.tooltipVisible&&(this.hideDelayTimer=window.setTimeout(()=>{this.updateTooltipVisibility()},60))},this.clearHideDelayTimer=()=>{null!==this.hideDelayTimer&&(clearTimeout(this.hideDelayTimer),this.hideDelayTimer=null)},this.startShowDelayTimer=()=>{this.isAnchorHoveredFocused||(this.delay>1?null===this.showDelayTimer&&(this.showDelayTimer=window.setTimeout(()=>{this.startHover()},this.delay)):this.startHover())},this.startHover=()=>{this.isAnchorHoveredFocused=!0,this.updateTooltipVisibility()},this.clearShowDelayTimer=()=>{null!==this.showDelayTimer&&(clearTimeout(this.showDelayTimer),this.showDelayTimer=null)},this.getAnchor=()=>{const e=this.getRootNode();return e instanceof ShadowRoot?e.getElementById(this.anchor):document.getElementById(this.anchor)},this.handleDocumentKeydown=e=>{!e.defaultPrevented&&this.tooltipVisible&&e.key===de.h&&(this.isAnchorHoveredFocused=!1,this.updateTooltipVisibility(),this.$emit("dismiss"))},this.updateTooltipVisibility=()=>{if(!1===this.visible)this.hideTooltip();else{if(!0===this.visible)return void this.showTooltip();if(this.isAnchorHoveredFocused||this.isRegionHovered)return void this.showTooltip();this.hideTooltip()}},this.showTooltip=()=>{this.tooltipVisible||(this.currentDirection=en(this),this.tooltipVisible=!0,document.addEventListener("keydown",this.handleDocumentKeydown),Ie.a.queueUpdate(this.setRegionProps))},this.hideTooltip=()=>{this.tooltipVisible&&(this.clearHideDelayTimer(),null!==this.region&&void 0!==this.region&&(this.region.removeEventListener("positionchange",this.handlePositionChange),this.region.viewportElement=null,this.region.anchorElement=null,this.region.removeEventListener("mouseover",this.handleRegionMouseOver),this.region.removeEventListener("mouseout",this.handleRegionMouseOut)),document.removeEventListener("keydown",this.handleDocumentKeydown),this.tooltipVisible=!1)},this.setRegionProps=()=>{this.tooltipVisible&&(this.region.viewportElement=this.viewportElement,this.region.anchorElement=this.anchorElement,this.region.addEventListener("positionchange",this.handlePositionChange),this.region.addEventListener("mouseover",this.handleRegionMouseOver,{passive:!0}),this.region.addEventListener("mouseout",this.handleRegionMouseOut,{passive:!0}))}}visibleChanged(){this.$fastController.isConnected&&(this.updateTooltipVisibility(),this.updateLayout())}anchorChanged(){this.$fastController.isConnected&&(this.anchorElement=this.getAnchor())}positionChanged(){this.$fastController.isConnected&&this.updateLayout()}anchorElementChanged(e){if(this.$fastController.isConnected){if(null!=e&&(e.removeEventListener("mouseover",this.handleAnchorMouseOver),e.removeEventListener("mouseout",this.handleAnchorMouseOut),e.removeEventListener("focusin",this.handleAnchorFocusIn),e.removeEventListener("focusout",this.handleAnchorFocusOut)),null!==this.anchorElement&&void 0!==this.anchorElement){this.anchorElement.addEventListener("mouseover",this.handleAnchorMouseOver,{passive:!0}),this.anchorElement.addEventListener("mouseout",this.handleAnchorMouseOut,{passive:!0}),this.anchorElement.addEventListener("focusin",this.handleAnchorFocusIn,{passive:!0}),this.anchorElement.addEventListener("focusout",this.handleAnchorFocusOut,{passive:!0});const e=this.anchorElement.id;null!==this.anchorElement.parentElement&&this.anchorElement.parentElement.querySelectorAll(":hover").forEach(t=>{t.id===e&&this.startShowDelayTimer()})}null!==this.region&&void 0!==this.region&&this.tooltipVisible&&(this.region.anchorElement=this.anchorElement),this.updateLayout()}}viewportElementChanged(){null!==this.region&&void 0!==this.region&&(this.region.viewportElement=this.viewportElement),this.updateLayout()}connectedCallback(){super.connectedCallback(),this.anchorElement=this.getAnchor(),this.updateTooltipVisibility()}disconnectedCallback(){this.hideTooltip(),this.clearShowDelayTimer(),this.clearHideDelayTimer(),super.disconnectedCallback()}updateLayout(){switch(this.verticalPositioningMode="locktodefault",this.horizontalPositioningMode="locktodefault",this.position){case"top":case"bottom":this.verticalDefaultPosition=this.position,this.horizontalDefaultPosition="center";break;case"right":case"left":case"start":case"end":this.verticalDefaultPosition="center",this.horizontalDefaultPosition=this.position;break;case"top-left":this.verticalDefaultPosition="top",this.horizontalDefaultPosition="left";break;case"top-right":this.verticalDefaultPosition="top",this.horizontalDefaultPosition="right";break;case"bottom-left":this.verticalDefaultPosition="bottom",this.horizontalDefaultPosition="left";break;case"bottom-right":this.verticalDefaultPosition="bottom",this.horizontalDefaultPosition="right";break;case"top-start":this.verticalDefaultPosition="top",this.horizontalDefaultPosition="start";break;case"top-end":this.verticalDefaultPosition="top",this.horizontalDefaultPosition="end";break;case"bottom-start":this.verticalDefaultPosition="bottom",this.horizontalDefaultPosition="start";break;case"bottom-end":this.verticalDefaultPosition="bottom",this.horizontalDefaultPosition="end";break;default:this.verticalPositioningMode="dynamic",this.horizontalPositioningMode="dynamic",this.verticalDefaultPosition=void 0,this.horizontalDefaultPosition="center"}}}Object(oe.a)([Object(ce.c)({mode:"boolean"})],or.prototype,"visible",void 0),Object(oe.a)([ce.c],or.prototype,"anchor",void 0),Object(oe.a)([ce.c],or.prototype,"delay",void 0),Object(oe.a)([ce.c],or.prototype,"position",void 0),Object(oe.a)([Object(ce.c)({attribute:"auto-update-mode"})],or.prototype,"autoUpdateMode",void 0),Object(oe.a)([Object(ce.c)({attribute:"horizontal-viewport-lock"})],or.prototype,"horizontalViewportLock",void 0),Object(oe.a)([Object(ce.c)({attribute:"vertical-viewport-lock"})],or.prototype,"verticalViewportLock",void 0),Object(oe.a)([se.d],or.prototype,"anchorElement",void 0),Object(oe.a)([se.d],or.prototype,"viewportElement",void 0),Object(oe.a)([se.d],or.prototype,"verticalPositioningMode",void 0),Object(oe.a)([se.d],or.prototype,"horizontalPositioningMode",void 0),Object(oe.a)([se.d],or.prototype,"horizontalInset",void 0),Object(oe.a)([se.d],or.prototype,"verticalInset",void 0),Object(oe.a)([se.d],or.prototype,"horizontalScaling",void 0),Object(oe.a)([se.d],or.prototype,"verticalScaling",void 0),Object(oe.a)([se.d],or.prototype,"verticalDefaultPosition",void 0),Object(oe.a)([se.d],or.prototype,"horizontalDefaultPosition",void 0),Object(oe.a)([se.d],or.prototype,"tooltipVisible",void 0),Object(oe.a)([se.d],or.prototype,"currentDirection",void 0);const sr=class extends or{connectedCallback(){super.connectedCallback(),Ee.v.setValueFor(this,Ee.gb)}}.compose({baseName:"tooltip",baseClass:or,template:(e,t)=>Se.a`
        ${Object(Ct.a)(e=>e.tooltipVisible,Se.a`
            <${e.tagFor(mn)}
                fixed-placement="true"
                auto-update-mode="${e=>e.autoUpdateMode}"
                vertical-positioning-mode="${e=>e.verticalPositioningMode}"
                vertical-default-position="${e=>e.verticalDefaultPosition}"
                vertical-inset="${e=>e.verticalInset}"
                vertical-scaling="${e=>e.verticalScaling}"
                horizontal-positioning-mode="${e=>e.horizontalPositioningMode}"
                horizontal-default-position="${e=>e.horizontalDefaultPosition}"
                horizontal-scaling="${e=>e.horizontalScaling}"
                horizontal-inset="${e=>e.horizontalInset}"
                vertical-viewport-lock="${e=>e.horizontalViewportLock}"
                horizontal-viewport-lock="${e=>e.verticalViewportLock}"
                dir="${e=>e.currentDirection}"
                ${Object(Ot.a)("region")}
            >
                <div class="tooltip" part="tooltip" role="tooltip">
                    <slot></slot>
                </div>
            </${e.tagFor(mn)}>
        `)}
    `,styles:(e,t)=>Oe.a`
    :host {
      position: relative;
      contain: layout;
      overflow: visible;
      height: 0;
      width: 0;
      z-index: 10000;
    }

    .tooltip {
      box-sizing: border-box;
      border-radius: calc(${Ee.q} * 1px);
      border: calc(${Ee.vb} * 1px) solid ${Ee.qb};
      background: ${Ee.v};
      color: ${Ee.fb};
      padding: 4px 12px;
      height: fit-content;
      width: fit-content;
      ${Tt.a}
      white-space: nowrap;
      box-shadow: ${kt.d};
    }

    ${e.tagFor(mn)} {
      display: flex;
      justify-content: center;
      align-items: center;
      overflow: visible;
      flex-direction: row;
    }

    ${e.tagFor(mn)}.right,
    ${e.tagFor(mn)}.left {
      flex-direction: column;
    }

    ${e.tagFor(mn)}.top .tooltip::after,
    ${e.tagFor(mn)}.bottom .tooltip::after,
    ${e.tagFor(mn)}.left .tooltip::after,
    ${e.tagFor(mn)}.right .tooltip::after {
      content: '';
      width: 12px;
      height: 12px;
      background: ${Ee.v};
      border-top: calc(${Ee.vb} * 1px) solid ${Ee.qb};
      border-left: calc(${Ee.vb} * 1px) solid ${Ee.qb};
      position: absolute;
    }

    ${e.tagFor(mn)}.top .tooltip::after {
      transform: translateX(-50%) rotate(225deg);
      bottom: 5px;
      left: 50%;
    }

    ${e.tagFor(mn)}.top .tooltip {
      margin-bottom: 12px;
    }

    ${e.tagFor(mn)}.bottom .tooltip::after {
      transform: translateX(-50%) rotate(45deg);
      top: 5px;
      left: 50%;
    }

    ${e.tagFor(mn)}.bottom .tooltip {
      margin-top: 12px;
    }

    ${e.tagFor(mn)}.left .tooltip::after {
      transform: translateY(-50%) rotate(135deg);
      top: 50%;
      right: 5px;
    }

    ${e.tagFor(mn)}.left .tooltip {
      margin-right: 12px;
    }

    ${e.tagFor(mn)}.right .tooltip::after {
      transform: translateY(-50%) rotate(-45deg);
      top: 50%;
      left: 5px;
    }

    ${e.tagFor(mn)}.right .tooltip {
      margin-left: 12px;
    }
  `.withBehaviors(Object(At.a)(Oe.a`
        :host([disabled]) {
          opacity: 1;
        }
        ${e.tagFor(mn)}.top .tooltip::after,
        ${e.tagFor(mn)}.bottom .tooltip::after,
        ${e.tagFor(mn)}.left .tooltip::after,
        ${e.tagFor(mn)}.right .tooltip::after {
          content: '';
          width: unset;
          height: unset;
        }
      `))}),cr=rn.compose({baseName:"divider",template:(e,t)=>Se.a`
    <template role="${e=>e.role}" aria-orientation="${e=>e.orientation}"></template>
`,styles:(e,t)=>Oe.a`
    ${Object(we.a)("block")} :host {
      box-sizing: content-box;
      height: 0;
      border: none;
      border-top: calc(${Ee.vb} * 1px) solid ${Ee.mb};
    }

    :host([orientation="vertical"]) {
      border: none;
      height: 100%;
      margin: 0 calc(${Ee.s} * 1px);
      border-left: calc(${Ee.vb} * 1px) solid ${Ee.mb};
  }
  `});var dr=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},lr=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};const ur=()=>{Object(T.a)(Qt,Me.a,sr,cr),Object(bi.b)(),Object(d.c)(),Object(A.b)("search-results",fr)};class fr extends p.a{static get styles(){return rr}get strings(){return ir}get queryString(){return this._queryString}set queryString(e){this._queryString!==e&&(this._queryString=e,this._currentPage=1,this.setLoadingState(!0),this.requestStateUpdate(!0))}get from(){return(this.currentPage-1)*this.size}get size(){return this._size}set size(e){e>this.maxPageSize?this._size=this.maxPageSize:this._size=e}get searchEndpoint(){return"/search/query"}get maxPageSize(){return 1e3}get currentPage(){return this._currentPage}set currentPage(e){this._currentPage!==e&&(this._currentPage=e,this.requestStateUpdate(!0))}constructor(){super(),this._size=10,this.entityTypes=["driveItem","listItem","site"],this.scopes=[],this.contentSources=[],this.version="v1.0",this.pagingMax=7,this.enableTopResults=!1,this.cacheEnabled=!1,this.cacheInvalidationPeriod=3e4,this.isRefreshing=!1,this.defaultFields=["webUrl","lastModifiedBy","lastModifiedDateTime","summary","displayName","name"],this._currentPage=1,this.onFirstPageClick=()=>{this.currentPage=1,this.scrollToFirstResult()},this.onPageBackClick=()=>{this.currentPage--,this.scrollToFirstResult()},this.onPageNextClick=()=>{this.currentPage++,this.scrollToFirstResult()},Object(Xa.c)("<mgt-search-results> is a preview component and may change prior to becoming generally available. See more information https://aka.ms/mgt/preview-components")}attributeChangedCallback(e,t,n){super.attributeChangedCallback(e,t,n),this.requestStateUpdate()}refresh(e=!1){this.isRefreshing=!0,e&&this.clearState(),this.requestStateUpdate(e)}clearState(){this.response=null}render(){var e,t,n,a,i,r,o,s;let c=null,d=null,l=null;return this.hasTemplate("header")&&(d=this.renderTemplate("header",this.response)),l=this.renderFooter(null===(t=null===(e=this.response)||void 0===e?void 0:e.value[0])||void 0===t?void 0:t.hitsContainers[0]),c=this.isLoadingState?this.renderLoading():this.error?this.renderError():this.response&&this.hasTemplate("default")?this.renderTemplate("default",this.response)||u.c``:(null===(a=null===(n=this.response)||void 0===n?void 0:n.value[0])||void 0===a?void 0:a.hitsContainers[0])?u.c`${null===(s=null===(o=null===(r=null===(i=this.response)||void 0===i?void 0:i.value[0])||void 0===r?void 0:r.hitsContainers[0])||void 0===o?void 0:o.hits)||void 0===s?void 0:s.map(e=>this.renderResult(e))}`:this.hasTemplate("no-data")?this.renderTemplate("no-data",null):u.c``,u.c`
      ${d}
      <div class="search-results">
        ${c}
      </div>
      ${l}`}loadState(){var e,t,n,a,r,o,s,c,d;return s=this,void 0,d=function*(){const s=_.a.globalProvider;if(this.error=null,s&&s.state===i.c.SignedIn){if(this.queryString){try{const i=this.getRequestOptions();let c;const d=JSON.stringify({endpoint:`${this.version}${this.searchEndpoint}`,requestOptions:i});let l=null;if(this.shouldRetrieveCache()){c=V.a.getCache(K.a.search,K.a.search.stores.responses);const e=Object(Ke.i)()?yield c.getValue(d):null;e&&Object(Ke.m)(this.cacheInvalidationPeriod)>Date.now()-e.timeCached&&(l=JSON.parse(e.response))}if(!l){const u=s.graph.forComponent(this);let f=u.api(this.searchEndpoint).version(this.version);if((null===(e=this.scopes)||void 0===e?void 0:e.length)&&(f=f.middlewareOptions(Object(b.a)(...this.scopes))),l=yield f.post({requests:[i]}),this.fetchThumbnail){const e=u.createBatch(),i=it.a.fromGraph(u).createBatch(),s=(null===(t=l.value)||void 0===t?void 0:t.length)&&(null===(n=l.value[0].hitsContainers)||void 0===n?void 0:n.length)&&null!==(r=null===(a=l.value[0].hitsContainers[0])||void 0===a?void 0:a.hits)&&void 0!==r?r:[];for(const t of s){const n=t.resource;!(n.size>0||(null===(o=n.webUrl)||void 0===o?void 0:o.endsWith(".aspx")))||"#microsoft.graph.driveItem"!==n["@odata.type"]&&"#microsoft.graph.listItem"!==n["@odata.type"]||("#microsoft.graph.listItem"===n["@odata.type"]?i.get(t.hitId.toString(),`/sites/${n.parentReference.siteId}/pages/${n.id}`):e.get(t.hitId.toString(),`/drives/${n.parentReference.driveId}/items/${n.id}/thumbnails/0/medium`))}const c=e=>{if(e&&e.size>0)for(const[t,n]of e){const e=l.value[0].hitsContainers[0].hits[t],a="#microsoft.graph.listItem"===e.resource["@odata.type"]?{url:n.content.thumbnailWebUrl}:{url:n.content.url};e.resource.thumbnail=a}};try{c(yield e.executeAll()),c(yield i.executeAll())}catch(e){}}this.shouldUpdateCache()&&l&&(c=V.a.getCache(K.a.search,K.a.search.stores.responses),yield c.putValue(d,{response:JSON.stringify(l)}))}Object(O.b)(this.response,l)||(this.response=l)}catch(e){this.error=e}this.response&&(this.error=null)}else this.response=null;this.isRefreshing=!1,this.fireCustomEvent("dataChange",{response:this.response,error:this.error})}},new((c=void 0)||(c=Promise))(function(e,t){function n(e){try{i(d.next(e))}catch(e){t(e)}}function a(e){try{i(d.throw(e))}catch(e){t(e)}}function i(t){var i;t.done?e(t.value):(i=t.value,i instanceof c?i:new c(function(e){e(i)})).then(n,a)}i((d=d.apply(s,[])).next())})}renderLoading(){return this.renderTemplate("loading",null)||u.c`
        ${[...Array(this.size)].map(()=>u.c`
            <div class="search-result">
              <div class="search-result-grid">
                <div class="search-result-icon">
                  <fluent-skeleton class="search-result-icon__shimmer" shape="rect" shimmer></fluent-skeleton>
                </div>
                <div class="searc-result-content">
                  <div class="search-result-name">
                    <fluent-skeleton class="search-result-name__shimmer" shape="rect" shimmer></fluent-skeleton>
                  </div>
                  <div class="search-result-info">
                    <div class="search-result-author">
                      <fluent-skeleton class="search-result-author__shimmer" shape="circle" shimmer></fluent-skeleton>
                    </div>
                    <div class="search-result-date">
                      <fluent-skeleton class="search-result-date__shimmer" shape="rect" shimmer></fluent-skeleton>
                    </div>
                  </div>
                  <fluent-skeleton class="search-result-content__shimmer" shape="rect" shimmer></fluent-skeleton>
                  <fluent-skeleton class="search-result-content__shimmer" shape="rect" shimmer></fluent-skeleton>
                </div>
                ${this.fetchThumbnail&&u.c`
                    <div class="search-result-thumbnail">
                      <fluent-skeleton class="search-result-thumbnail__shimmer" shape="rect" shimmer></fluent-skeleton>
                    </div>
                  `}
              </div>
              <fluent-divider></fluent-divider>
            </div>
          `)}
       `}renderResult(e){const t=this.getResourceType(e.resource);if(this.hasTemplate(`result-${t}`))return this.renderTemplate(`result-${t}`,e,e.hitId);switch(e.resource["@odata.type"]){case"#microsoft.graph.driveItem":return this.renderDriveItem(e);case"#microsoft.graph.site":return this.renderSite(e);case"#microsoft.graph.person":return this.renderPerson(e);case"#microsoft.graph.drive":case"#microsoft.graph.list":return this.renderList(e);case"#microsoft.graph.listItem":return this.renderListItem(e);case"#microsoft.graph.search.bookmark":return this.renderBookmark(e);case"#microsoft.graph.search.acronym":return this.renderAcronym(e);case"#microsoft.graph.search.qna":return this.renderQnA(e);default:return this.renderDefault(e)}}renderFooter(e){if(this.pagingRequired(e)){const t=this.getActivePages(e.total);return u.c`
        <div class="search-results-pages">
          ${this.renderPreviousPage()}
          ${this.renderFirstPage(t)}
          ${this.renderAllPages(t)}
          ${this.renderNextPage()}
        </div>
      `}}pagingRequired(e){return(null==e?void 0:e.moreResultsAvailable)||this.currentPage*this.size<(null==e?void 0:e.total)}getActivePages(e){const t=[],n=(()=>{const e=this.currentPage-Math.floor(this.pagingMax/2)-1;return e>=Math.floor(this.pagingMax/2)?e:0})();if(n+1>this.pagingMax-this.currentPage||this.pagingMax===this.currentPage)for(let a=n+1;a<Math.ceil(e/this.size)&&a<this.pagingMax+(this.currentPage-1)&&t.length<this.pagingMax-2;++a)t.push(a+1);else for(let e=n;e<this.pagingMax;++e)t.push(e+1);return t}renderAllPages(e){return u.c`
      ${e.map(e=>u.c`
            <fluent-button
              title="${ir.page} ${e}"
              appearance="stealth"
              class="${e===this.currentPage?"search-results-page-active":"search-results-page"}"
              @click="${()=>this.onPageClick(e)}">
                ${e}
            </fluent-button>`)}`}renderFirstPage(e){return u.c`
      ${e.some(e=>1===e)?u.d:u.c`
              <fluent-button
                 title="${ir.page} 1"
                 appearance="stealth"
                 class="search-results-page"
                 @click="${this.onFirstPageClick}">
                 1
               </fluent-button>`?u.c`
              <fluent-button
                id="page-back-dot"
                appearance="stealth"
                class="search-results-page"
                title="${this.getDotButtonTitle()}"
                @click="${()=>this.onPageClick(this.currentPage-Math.ceil(this.pagingMax/2))}"
              >
                ...
              </fluent-button>`:u.d}`}getDotButtonTitle(){return`${ir.back} ${Math.ceil(this.pagingMax/2)} ${ir.pages}`}renderPreviousPage(){return this.currentPage>1?u.c`
          <fluent-button
            appearance="stealth"
            class="search-results-page"
            title="${ir.back}"
            @click="${this.onPageBackClick}">
              ${Object(S.b)(S.a.ChevronLeft)}
            </fluent-button>`:u.d}renderNextPage(){return this.isLastPage()?u.d:u.c`
          <fluent-button
            appearance="stealth"
            class="search-results-page"
            title="${ir.next}"
            aria-label="${ir.next}"
            @click="${this.onPageNextClick}">
              ${Object(S.b)(S.a.ChevronRight)}
            </fluent-button>`}onPageClick(e){this.currentPage=e,this.scrollToFirstResult()}isLastPage(){return this.currentPage===Math.ceil(this.response.value[0].hitsContainers[0].total/this.size)}scrollToFirstResult(){this.renderRoot.querySelector(".search-results").scrollIntoView({block:"start",behavior:"smooth"})}getResourceType(e){return e["@odata.type"].split(".").pop()}renderDriveItem(e){var t,n;const a=e.resource;return m.a`
      <div class="search-result-grid">
        <div class="search-result-icon">
          <mgt-file
            .fileDetails="${e.resource}"
            view="image"
            class="file-icon">
          </mgt-file>
        </div>
        <div class="search-result-content">
          <div class="search-result-name">
            <a href="${a.webUrl}?Web=1" target="_blank">${Object(Ke.q)(a.name)}</a>
          </div>
          <div class="search-result-info">
            <div class="search-result-author">
              <mgt-person
                person-query=${a.lastModifiedBy.user.email}
                view="oneLine"
                person-card="hover"
                show-presence="true">
              </mgt-person>
            </div>
            <div class="search-result-date">
              &nbsp; ${ir.modified} ${Object(Ke.l)(new Date(a.lastModifiedDateTime))}
            </div>
          </div>
          <div class="search-result-summary" .innerHTML="${Object(Ke.p)(e.summary)}"></div>
        </div>
        ${(null===(t=a.thumbnail)||void 0===t?void 0:t.url)&&u.c`
          <div class="search-result-thumbnail">
            <a href="${a.webUrl}" target="_blank"><img alt="${a.name}" src="${null===(n=a.thumbnail)||void 0===n?void 0:n.url}" /></a>
          </div>`}

      </div>
      <fluent-divider></fluent-divider>
    `}renderSite(e){const t=e.resource;return u.c`
      <div class="search-result-grid">
        <div class="search-result-icon">
          ${this.getResourceIcon(t)}
        </div>
        <div class="searc-result-content">
          <div class="search-result-name">
            <a href="${t.webUrl}" target="_blank">${t.displayName}</a>
          </div>
          <div class="search-result-info">
            <div class="search-result-url">
              <a href="${t.webUrl}" target="_blank">${t.webUrl}</a>
            </div>
          </div>
          <div class="search-result-summary" .innerHTML="${Object(Ke.p)(e.summary)}"></div>
        </div>
      </div>
      <fluent-divider></fluent-divider>
    `}renderList(e){const t=e.resource;return m.a`
      <div class="search-result-grid">
        <div class="search-result-icon">
          <mgt-file
            .fileDetails="${e.resource}"
            view="image">
          </mgt-file>
        </div>
        <div class="search-result-content">
          <div class="search-result-name">
            <a href="${t.webUrl}?Web=1" target="_blank">
              ${Object(Ke.q)(t.name||Object(Ke.k)(t.webUrl))}
            </a>
          </div>
          <div class="search-result-summary" .innerHTML="${Object(Ke.p)(e.summary)}"></div>
        </div>
      </div>
      <fluent-divider></fluent-divider>
    `}renderListItem(e){var t,n;const a=e.resource;return m.a`
      <div class="search-result-grid">
        <div class="search-result-icon">
          ${a.webUrl.endsWith(".aspx")?Object(S.b)(S.a.News):Object(S.b)(S.a.FileOuter)}
        </div>
        <div class="search-result-content">
          <div class="search-result-name">
            <a href="${a.webUrl}?Web=1" target="_blank">
              ${Object(Ke.q)(a.name||Object(Ke.k)(a.webUrl))}
            </a>
          </div>
          <div class="search-result-info">
            <div class="search-result-author">
              <mgt-person
                person-query=${a.lastModifiedBy.user.email}
                view="oneLine"
                person-card="hover"
                show-presence="true">
              </mgt-person>
            </div>
            <div class="search-result-date">
              &nbsp; ${ir.modified} ${Object(Ke.l)(new Date(a.lastModifiedDateTime))}
            </div>
          </div>
          <div class="search-result-summary" .innerHTML="${Object(Ke.p)(e.summary)}"></div>
        </div>
        ${(null===(t=a.thumbnail)||void 0===t?void 0:t.url)&&u.c`
          <div class="search-result-thumbnail">
            <a href="${a.webUrl}" target="_blank"><img alt="${Object(Ke.q)(a.name||Object(Ke.k)(a.webUrl))}" src="${(null===(n=a.thumbnail)||void 0===n?void 0:n.url)||u.d}" /></a>
          </div>`}
      </div>
      <fluent-divider></fluent-divider>
    `}renderPerson(e){const t=e.resource;return m.a`
      <div class="search-result">
        <mgt-person
          view="fourLines"
          person-query=${t.userPrincipalName}
          person-card="hover"
          show-presence="true">
        </mgt-person>
      </div>
      <fluent-divider></fluent-divider>
    `}renderBookmark(e){return this.renderAnswer(e,S.a.DoubleBookmark)}renderAcronym(e){return this.renderAnswer(e,S.a.BookOpen)}renderQnA(e){return this.renderAnswer(e,S.a.BookQuestion)}renderAnswer(e,t){const n=e.resource;return u.c`
      <div class="search-result-grid search-result-answer">
        <div class="search-result-icon">
          ${Object(S.b)(t)}
        </div>
        <div class="search-result-content">
          <div class="search-result-name">
            <a href="${this.getResourceUrl(n)}?Web=1" target="_blank">${n.displayName}</a>
          </div>
          <div class="search-result-summary">${n.description}</div>
        </div>
      </div>
      <fluent-divider></fluent-divider>
    `}renderDefault(e){const t=e.resource,n=this.getResourceUrl(t);return u.c`
      <div class="search-result-grid">
        <div class="search-result-icon">
          ${this.getResourceIcon(t)}
        </div>
        <div class="search-result-content">
          <div class="search-result-name">
            ${n?u.c`
                  <a href="${n}?Web=1" target="_blank">${this.getResourceName(t)}</a>
                `:u.c`
                  ${this.getResourceName(t)}
                `}
          </div>
          <div class="search-result-summary" .innerHTML="${this.getResultSummary(e)}"></div>
        </div>
      </div>
      <fluent-divider></fluent-divider>
    `}getResourceUrl(e){return e.webUrl||e.webLink||null}getResourceName(e){return e.displayName||e.subject||Object(Ke.q)(e.name)}getResultSummary(e){var t;return Object(Ke.p)(e.summary||(null===(t=e.resource)||void 0===t?void 0:t.description)||null)}getResourceIcon(e){switch(e["@odata.type"]){case"#microsoft.graph.site":return Object(S.b)(S.a.Globe);case"#microsoft.graph.message":return Object(S.b)(S.a.Email);case"#microsoft.graph.event":return Object(S.b)(S.a.Event);case"microsoft.graph.chatMessage":return Object(S.b)(S.a.SmallChat);default:return Object(S.b)(S.a.FileOuter)}}shouldRetrieveCache(){return Object(Ke.i)()&&this.cacheEnabled&&!this.isRefreshing}shouldUpdateCache(){return Object(Ke.i)()&&this.cacheEnabled}getRequestOptions(){const e={entityTypes:this.entityTypes,query:{queryString:this.queryString},from:this.from?this.from:void 0,size:this.size?this.size:void 0,fields:this.getFields(),enableTopResults:this.enableTopResults?this.enableTopResults:void 0};return this.entityTypes.includes("externalItem")&&(e.contentSources=this.contentSources),"beta"===this.version&&(e.query.queryTemplate=this.queryTemplate?this.queryTemplate:void 0),e}getFields(){if(this.fields)return this.defaultFields.concat(this.fields)}}dr([Object(f.b)({attribute:"query-string",type:String}),lr("design:type",String),lr("design:paramtypes",[String])],fr.prototype,"queryString",null),dr([Object(f.b)({attribute:"query-template",type:String}),lr("design:type",String)],fr.prototype,"queryTemplate",void 0),dr([Object(f.b)({attribute:"entity-types",converter:e=>e.split(",").map(e=>e.trim()),type:String}),lr("design:type",Array)],fr.prototype,"entityTypes",void 0),dr([Object(f.b)({attribute:"scopes",converter:(e,t)=>e?e.toLowerCase().split(","):null}),lr("design:type",Array)],fr.prototype,"scopes",void 0),dr([Object(f.b)({attribute:"content-sources",converter:(e,t)=>e?e.toLowerCase().split(","):null}),lr("design:type",Array)],fr.prototype,"contentSources",void 0),dr([Object(f.b)({attribute:"version",reflect:!0,type:String}),lr("design:type",Object)],fr.prototype,"version",void 0),dr([Object(f.b)({attribute:"size",reflect:!0,type:Number}),lr("design:type",Number),lr("design:paramtypes",[Object])],fr.prototype,"size",null),dr([Object(f.b)({attribute:"paging-max",reflect:!0,type:Number}),lr("design:type",Object)],fr.prototype,"pagingMax",void 0),dr([Object(f.b)({attribute:"fetch-thumbnail",type:Boolean}),lr("design:type",Boolean)],fr.prototype,"fetchThumbnail",void 0),dr([Object(f.b)({attribute:"fields",converter:e=>e.split(",").map(e=>e.trim()),type:String}),lr("design:type",Array)],fr.prototype,"fields",void 0),dr([Object(f.b)({attribute:"enable-top-results",reflect:!0,type:Boolean}),lr("design:type",Object)],fr.prototype,"enableTopResults",void 0),dr([Object(f.b)({attribute:"cache-enabled",reflect:!0,type:Boolean}),lr("design:type",Object)],fr.prototype,"cacheEnabled",void 0),dr([Object(f.b)({attribute:"cache-invalidation-period",reflect:!0,type:Number}),lr("design:type",Object)],fr.prototype,"cacheInvalidationPeriod",void 0),dr([Object(f.c)(),lr("design:type",Object)],fr.prototype,"response",void 0),dr([Object(f.c)(),lr("design:type",Number),lr("design:paramtypes",[Number])],fr.prototype,"currentPage",null);const pr=()=>{Object(d.c)(),Object(l.registerMgtPersonCardComponent)(),B(),X(),Fe(),et(),M(),Pn(),sa(),_i(),Object(bi.b)(),Object(gi.b)(),li(),Ii(),Ni(),nr(),ur(),Object(Ye.b)()};var mr=n("YZrM"),_r=n("icq5"),hr=n("1/1t"),br=n("tnKu"),gr=n("Ys1X"),vr=n("3ATu"),yr=n("fU9z"),Sr=n("qwWh")},o2vq:function(e,t,n){"use strict";n.d(t,"a",function(){return a});const a=e=>{const t=e;return t.statusCode&&"code"in t&&"body"in t&&t.date&&"message"in t&&"name"in t&&"requestId"in t}},o87Z:function(e,t,n){"use strict";n.d(t,"b",function(){return f}),n.d(t,"a",function(){return p});var a=n("oZuh"),i=n("5ZAu"),r=n("QBS5"),o=n("oePG"),s=n("Gy7L");const c="form-associated-proxy",d="ElementInternals",l=d in window&&"setFormValue"in window[d].prototype,u=new WeakMap;function f(e){const t=class extends e{constructor(...e){super(...e),this.dirtyValue=!1,this.disabled=!1,this.proxyEventsToBlock=["change","click"],this.proxyInitialized=!1,this.required=!1,this.initialValue=this.initialValue||"",this.elementInternals||(this.formResetCallback=this.formResetCallback.bind(this))}static get formAssociated(){return l}get validity(){return this.elementInternals?this.elementInternals.validity:this.proxy.validity}get form(){return this.elementInternals?this.elementInternals.form:this.proxy.form}get validationMessage(){return this.elementInternals?this.elementInternals.validationMessage:this.proxy.validationMessage}get willValidate(){return this.elementInternals?this.elementInternals.willValidate:this.proxy.willValidate}get labels(){if(this.elementInternals)return Object.freeze(Array.from(this.elementInternals.labels));if(this.proxy instanceof HTMLElement&&this.proxy.ownerDocument&&this.id){const e=this.proxy.labels,t=Array.from(this.proxy.getRootNode().querySelectorAll(`[for='${this.id}']`)),n=e?t.concat(Array.from(e)):t;return Object.freeze(n)}return a.d}valueChanged(e,t){this.dirtyValue=!0,this.proxy instanceof HTMLElement&&(this.proxy.value=this.value),this.currentValue=this.value,this.setFormValue(this.value),this.validate()}currentValueChanged(){this.value=this.currentValue}initialValueChanged(e,t){this.dirtyValue||(this.value=this.initialValue,this.dirtyValue=!1)}disabledChanged(e,t){this.proxy instanceof HTMLElement&&(this.proxy.disabled=this.disabled),i.a.queueUpdate(()=>this.classList.toggle("disabled",this.disabled))}nameChanged(e,t){this.proxy instanceof HTMLElement&&(this.proxy.name=this.name)}requiredChanged(e,t){this.proxy instanceof HTMLElement&&(this.proxy.required=this.required),i.a.queueUpdate(()=>this.classList.toggle("required",this.required)),this.validate()}get elementInternals(){if(!l)return null;let e=u.get(this);return e||(e=this.attachInternals(),u.set(this,e)),e}connectedCallback(){super.connectedCallback(),this.addEventListener("keypress",this._keypressHandler),this.value||(this.value=this.initialValue,this.dirtyValue=!1),this.elementInternals||(this.attachProxy(),this.form&&this.form.addEventListener("reset",this.formResetCallback))}disconnectedCallback(){this.proxyEventsToBlock.forEach(e=>this.proxy.removeEventListener(e,this.stopPropagation)),!this.elementInternals&&this.form&&this.form.removeEventListener("reset",this.formResetCallback)}checkValidity(){return this.elementInternals?this.elementInternals.checkValidity():this.proxy.checkValidity()}reportValidity(){return this.elementInternals?this.elementInternals.reportValidity():this.proxy.reportValidity()}setValidity(e,t,n){this.elementInternals?this.elementInternals.setValidity(e,t,n):"string"==typeof t&&this.proxy.setCustomValidity(t)}formDisabledCallback(e){this.disabled=e}formResetCallback(){this.value=this.initialValue,this.dirtyValue=!1}attachProxy(){var e;this.proxyInitialized||(this.proxyInitialized=!0,this.proxy.style.display="none",this.proxyEventsToBlock.forEach(e=>this.proxy.addEventListener(e,this.stopPropagation)),this.proxy.disabled=this.disabled,this.proxy.required=this.required,"string"==typeof this.name&&(this.proxy.name=this.name),"string"==typeof this.value&&(this.proxy.value=this.value),this.proxy.setAttribute("slot",c),this.proxySlot=document.createElement("slot"),this.proxySlot.setAttribute("name",c)),null===(e=this.shadowRoot)||void 0===e||e.appendChild(this.proxySlot),this.appendChild(this.proxy)}detachProxy(){var e;this.removeChild(this.proxy),null===(e=this.shadowRoot)||void 0===e||e.removeChild(this.proxySlot)}validate(e){this.proxy instanceof HTMLElement&&this.setValidity(this.proxy.validity,this.proxy.validationMessage,e)}setFormValue(e,t){this.elementInternals&&this.elementInternals.setFormValue(e,t||e)}_keypressHandler(e){if(e.key===s.g&&this.form instanceof HTMLFormElement){const e=this.form.querySelector("[type=submit]");null==e||e.click()}}stopPropagation(e){e.stopPropagation()}};return Object(r.c)({mode:"boolean"})(t.prototype,"disabled"),Object(r.c)({mode:"fromView",attribute:"value"})(t.prototype,"initialValue"),Object(r.c)({attribute:"current-value"})(t.prototype,"currentValue"),Object(r.c)(t.prototype,"name"),Object(r.c)({mode:"boolean"})(t.prototype,"required"),Object(o.d)(t.prototype,"value"),t}function p(e){class t extends(f(e)){}class n extends t{constructor(...e){super(e),this.dirtyChecked=!1,this.checkedAttribute=!1,this.checked=!1,this.dirtyChecked=!1}checkedAttributeChanged(){this.defaultChecked=this.checkedAttribute}defaultCheckedChanged(){this.dirtyChecked||(this.checked=this.defaultChecked,this.dirtyChecked=!1)}checkedChanged(e,t){this.dirtyChecked||(this.dirtyChecked=!0),this.currentChecked=this.checked,this.updateForm(),this.proxy instanceof HTMLInputElement&&(this.proxy.checked=this.checked),void 0!==e&&this.$emit("change"),this.validate()}currentCheckedChanged(e,t){this.checked=this.currentChecked}updateForm(){const e=this.checked?this.value:null;this.setFormValue(e,e)}connectedCallback(){super.connectedCallback(),this.updateForm()}formResetCallback(){super.formResetCallback(),this.checked=!!this.checkedAttribute,this.dirtyChecked=!1}}return Object(r.c)({attribute:"checked",mode:"boolean"})(n.prototype,"checkedAttribute"),Object(r.c)({attribute:"current-checked",converter:r.d})(n.prototype,"currentChecked"),Object(o.d)(n.prototype,"defaultChecked"),Object(o.d)(n.prototype,"checked"),n}},oZuh:function(e,t,n){"use strict";(function(e){n.d(t,"a",function(){return a}),n.d(t,"b",function(){return r}),n.d(t,"d",function(){return o}),n.d(t,"c",function(){return s});const a=function(){if("undefined"!=typeof globalThis)return globalThis;if(void 0!==e)return e;if("undefined"!=typeof self)return self;if("undefined"!=typeof window)return window;try{return new Function("return this")()}catch(e){return{}}}();void 0===a.trustedTypes&&(a.trustedTypes={createPolicy:(e,t)=>t});const i={configurable:!1,enumerable:!1,writable:!1};void 0===a.FAST&&Reflect.defineProperty(a,"FAST",Object.assign({value:Object.create(null)},i));const r=a.FAST;if(void 0===r.getById){const e=Object.create(null);Reflect.defineProperty(r,"getById",Object.assign({value(t,n){let a=e[t];return void 0===a&&(a=n?e[t]=n():null),a}},i))}const o=Object.freeze([]);function s(){const e=new WeakMap;return function(t){let n=e.get(t);if(void 0===n){let a=Reflect.getPrototypeOf(t);for(;void 0===n&&null!==a;)n=e.get(a),a=Reflect.getPrototypeOf(a);n=void 0===n?[]:n.slice(0),e.set(t,n)}return n}}}).call(this,n("EKCE"))},oaDa:function(e,t,n){"use strict";n.d(t,"a",function(){return a});class a{static renderTemplate(e,t,n){var a;let i;if(t.$parentTemplateContext&&(n=Object.assign(Object.assign({},n),{$parent:t.$parentTemplateContext})),null===(a=t.content)||void 0===a?void 0:a.childNodes.length){const a=t.content.cloneNode(!0);i=this.renderNode(a,e,n)}else if(t.childNodes.length){const a=document.createElement("div");for(const e of t.childNodes)a.appendChild(this.simpleCloneNode(e));i=this.renderNode(a,e,n)}i&&e.appendChild(i)}static setBindingSyntax(e,t){this._startExpression=e,this._endExpression=t;const n=this.escapeRegex(this._startExpression),a=this.escapeRegex(this._endExpression);this._expression=new RegExp(`${n}\\s*([$\\w\\.,'"\\s()\\[\\]]+)\\s*${a}`,"g")}static get globalContext(){return this._globalContext}static get expression(){return this._expression||this.setBindingSyntax("{{","}}"),this._expression}static escapeRegex(e){return e.replace(/[-/\\^$*+?.()|[\]{}]/g,"\\$&")}static simpleCloneNode(e){if(!e)return null;const t=e.cloneNode(!1);for(const n of e.childNodes){const e=this.simpleCloneNode(n);e&&t.appendChild(e)}return t}static expandExpressionsAsString(e,t){return e.replace(this.expression,(e,n)=>{const a=this.evalInContext(n||this.trimExpression(e),t);return a?"object"==typeof a?JSON.stringify(a):a.toString():""})}static renderNode(e,t,n){if("#text"===e.nodeName)return e.textContent=this.expandExpressionsAsString(e.textContent,n),e;if("TEMPLATE"===e.nodeName)return e.$parentTemplateContext=n,e;const a=e;if(a.attributes)for(const e of a.attributes)if("data-props"===e.name){const i=this.trimExpression(e.value);for(const e of i.split(",")){const i=e.trim().split(":");if(2===i.length){const e=i[0].trim(),r=this.evalInContext(i[1].trim(),n);e.startsWith("@")?"function"==typeof r&&a.addEventListener(e.substring(1),e=>r(e,n,t)):a[e]=r}}}else a.setAttribute(e.name,this.expandExpressionsAsString(e.value,n));const i=[],r=[];let o=!1;for(const a of e.childNodes){const e=a;let s=!1;if(e.dataset){let c=!1;if(e.dataset.if){const t=e.dataset.if;this.evalBoolInContext(this.trimExpression(t),n)?(e.removeAttribute("data-if"),o=!0,s=!0):(r.push(e),c=!0)}else void 0!==e.dataset.else&&(o?(r.push(e),c=!0):e.removeAttribute("data-else"));e.dataset.for&&!c?i.push(e):c||this.renderNode(a,t,n)}else this.renderNode(a,t,n);s||"#text"===a.nodeName||(o=!1)}for(const e of r)a.removeChild(e);for(const e of i){const i=e.dataset.for,r=this.trimExpression(i).split(/\s(in|of)\s/i);if(3===r.length){const i=r[0].trim(),o=r[2].trim(),s=this.evalInContext(o,n);if(Array.isArray(s)){e.removeAttribute("data-for");for(let r=0;r<s.length;r++){const o=Object.assign({$index:r},n);o[i]=s[r];const c=e.cloneNode(!0);this.renderNode(c,t,o),a.insertBefore(c,e)}}a.removeChild(e)}}return e}static evalBoolInContext(e,t){return t=Object.assign(Object.assign({},t),this.globalContext),new Function("with(this) { return !!("+e+")}").call(t)}static evalInContext(e,t){t=Object.assign(Object.assign({},t),this.globalContext);const n=new Function("with(this) { return "+e+";}");let a;try{a=n.call(t)}catch(e){}return a}static trimExpression(e){return(e=e.trim()).startsWith(this._startExpression)&&e.endsWith(this._endExpression)&&(e=(e=e.substr(this._startExpression.length,e.length-this._startExpression.length-this._endExpression.length)).trim()),e}}a._globalContext={}},oePG:function(e,t,n){"use strict";n.d(t,"b",function(){return o}),n.d(t,"d",function(){return s}),n.d(t,"e",function(){return c}),n.d(t,"a",function(){return l}),n.d(t,"c",function(){return u});var a=n("5ZAu"),i=n("oZuh"),r=n("O/9U");const o=i.b.getById(2,()=>{const e=/(:|&&|\|\||if)/,t=new WeakMap,n=a.a.queueUpdate;let o,s=e=>{throw new Error("Must call enableArrayObservation before observing arrays.")};function c(e){let n=e.$fastController||t.get(e);return void 0===n&&(Array.isArray(e)?n=s(e):t.set(e,n=new r.a(e))),n}const d=Object(i.c)();class l{constructor(e){this.name=e,this.field=`_${e}`,this.callback=`${e}Changed`}getValue(e){return void 0!==o&&o.watch(e,this.name),e[this.field]}setValue(e,t){const n=this.field,a=e[n];if(a!==t){e[n]=t;const i=e[this.callback];"function"==typeof i&&i.call(e,a,t),c(e).notify(this.name)}}}class u extends r.b{constructor(e,t,n=!1){super(e,t),this.binding=e,this.isVolatileBinding=n,this.needsRefresh=!0,this.needsQueue=!0,this.first=this,this.last=null,this.propertySource=void 0,this.propertyName=void 0,this.notifier=void 0,this.next=void 0}observe(e,t){this.needsRefresh&&null!==this.last&&this.disconnect();const n=o;o=this.needsRefresh?this:void 0,this.needsRefresh=this.isVolatileBinding;const a=this.binding(e,t);return o=n,a}disconnect(){if(null!==this.last){let e=this.first;for(;void 0!==e;)e.notifier.unsubscribe(this,e.propertyName),e=e.next;this.last=null,this.needsRefresh=this.needsQueue=!0}}watch(e,t){const n=this.last,a=c(e),i=null===n?this.first:{};if(i.propertySource=e,i.propertyName=t,i.notifier=a,a.subscribe(this,t),null!==n){if(!this.needsRefresh){let t;o=void 0,t=n.propertySource[n.propertyName],o=this,e===t&&(this.needsRefresh=!0)}n.next=i}this.last=i}handleChange(){this.needsQueue&&(this.needsQueue=!1,n(this))}call(){null!==this.last&&(this.needsQueue=!0,this.notify(this))}records(){let e=this.first;return{next:()=>{const t=e;return void 0===t?{value:void 0,done:!0}:(e=e.next,{value:t,done:!1})},[Symbol.iterator]:function(){return this}}}}return Object.freeze({setArrayObserverFactory(e){s=e},getNotifier:c,track(e,t){void 0!==o&&o.watch(e,t)},trackVolatile(){void 0!==o&&(o.needsRefresh=!0)},notify(e,t){c(e).notify(t)},defineProperty(e,t){"string"==typeof t&&(t=new l(t)),d(e).push(t),Reflect.defineProperty(e,t.name,{enumerable:!0,get:function(){return t.getValue(this)},set:function(e){t.setValue(this,e)}})},getAccessors:d,binding(e,t,n=this.isVolatileBinding(e)){return new u(e,t,n)},isVolatileBinding:t=>e.test(t.toString())})});function s(e,t){o.defineProperty(e,t)}function c(e,t,n){return Object.assign({},n,{get:function(){return o.trackVolatile(),n.get.apply(this)}})}const d=i.b.getById(3,()=>{let e=null;return{get:()=>e,set(t){e=t}}});class l{constructor(){this.index=0,this.length=0,this.parent=null,this.parentContext=null}get event(){return d.get()}get isEven(){return this.index%2==0}get isOdd(){return this.index%2!=0}get isFirst(){return 0===this.index}get isInMiddle(){return!this.isFirst&&!this.isLast}get isLast(){return this.index===this.length-1}static setEvent(e){d.set(e)}}o.defineProperty(l.prototype,"index"),o.defineProperty(l.prototype,"length");const u=Object.seal(new l)},olMv:function(e,t,n){"use strict";n.d(t,"a",function(){return a});class a{createCSS(){return""}createBehavior(){}}},oqLQ:function(e,t,n){"use strict";n.d(t,"a",function(){return r});class a{constructor(e){this.listenerCache=new WeakMap,this.query=e}bind(e){const{query:t}=this,n=this.constructListener(e);n.bind(t)(),t.addListener(n),this.listenerCache.set(e,n)}unbind(e){const t=this.listenerCache.get(e);t&&(this.query.removeListener(t),this.listenerCache.delete(e))}}class i extends a{constructor(e,t){super(e),this.styles=t}static with(e){return t=>new i(e,t)}constructListener(e){let t=!1;const n=this.styles;return function(){const{matches:a}=this;a&&!t?(e.$fastController.addStyles(n),t=a):!a&&t&&(e.$fastController.removeStyles(n),t=a)}}unbind(e){super.unbind(e),e.$fastController.removeStyles(this.styles)}}const r=i.with(window.matchMedia("(forced-colors)"));i.with(window.matchMedia("(prefers-color-scheme: dark)")),i.with(window.matchMedia("(prefers-color-scheme: light)"))},ox5k:function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("5rba");const i=e=>null!=e?e:a.d},pnUA:function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("SUtl");class i extends a.a{static fromGraph(e){const t=new i(e.client);return t.setComponent(e.componentName),t}constructor(e,t="beta"){super(e,t)}forComponent(e){const t=new i(this.client);return this.setComponent(e),t}}},ptXv:function(e,t,n){"use strict";n.d(t,"b",function(){return c}),n.d(t,"a",function(){return w}),n.d(t,"c",function(){return x}),n.d(t,"d",function(){return C});const a=globalThis,i=a.ShadowRoot&&(void 0===a.ShadyCSS||a.ShadyCSS.nativeShadow)&&"adoptedStyleSheets"in Document.prototype&&"replace"in CSSStyleSheet.prototype,r=Symbol(),o=new WeakMap;class s{constructor(e,t,n){if(this._$cssResult$=!0,n!==r)throw Error("CSSResult is not constructable. Use `unsafeCSS` or `css` instead.");this.cssText=e,this.t=t}get styleSheet(){let e=this.o;const t=this.t;if(i&&void 0===e){const n=void 0!==t&&1===t.length;n&&(e=o.get(t)),void 0===e&&((this.o=e=new CSSStyleSheet).replaceSync(this.cssText),n&&o.set(t,e))}return e}toString(){return this.cssText}}const c=(e,...t)=>{const n=1===e.length?e[0]:t.reduce((t,n,a)=>t+(e=>{if(!0===e._$cssResult$)return e.cssText;if("number"==typeof e)return e;throw Error("Value passed to 'css' function must be a 'css' function result: "+e+". Use 'unsafeCSS' to pass non-literal values, but take care to ensure page security.")})(n)+e[a+1],e[0]);return new s(n,e,r)},d=i?e=>e:e=>e instanceof CSSStyleSheet?(e=>{let t="";for(const n of e.cssRules)t+=n.cssText;return(e=>new s("string"==typeof e?e:e+"",void 0,r))(t)})(e):e;var l,u,f;const{is:p,defineProperty:m,getOwnPropertyDescriptor:_,getOwnPropertyNames:h,getOwnPropertySymbols:b,getPrototypeOf:g}=Object,v=globalThis,y=v.trustedTypes,S=y?y.emptyScript:"",D=v.reactiveElementPolyfillSupport,I=(e,t)=>e,x={toAttribute(e,t){switch(t){case Boolean:e=e?S:null;break;case Object:case Array:e=null==e?e:JSON.stringify(e)}return e},fromAttribute(e,t){let n=e;switch(t){case Boolean:n=null!==e;break;case Number:n=null===e?null:Number(e);break;case Object:case Array:try{n=JSON.parse(e)}catch(e){n=null}}return n}},C=(e,t)=>!p(e,t),O={attribute:!0,type:String,converter:x,reflect:!1,hasChanged:C};null!==(l=Symbol.metadata)&&void 0!==l||(Symbol.metadata=Symbol("metadata")),null!==(u=v.litPropertyMetadata)&&void 0!==u||(v.litPropertyMetadata=new WeakMap);class w extends HTMLElement{static addInitializer(e){var t;this._$Ei(),(null!==(t=this.l)&&void 0!==t?t:this.l=[]).push(e)}static get observedAttributes(){return this.finalize(),this._$Eh&&[...this._$Eh.keys()]}static createProperty(e,t=O){if(t.state&&(t.attribute=!1),this._$Ei(),this.elementProperties.set(e,t),!t.noAccessor){const n=Symbol(),a=this.getPropertyDescriptor(e,n,t);void 0!==a&&m(this.prototype,e,a)}}static getPropertyDescriptor(e,t,n){var a;const{get:i,set:r}=null!==(a=_(this.prototype,e))&&void 0!==a?a:{get(){return this[t]},set(e){this[t]=e}};return{get(){return null==i?void 0:i.call(this)},set(t){const a=null==i?void 0:i.call(this);r.call(this,t),this.requestUpdate(e,a,n)},configurable:!0,enumerable:!0}}static getPropertyOptions(e){var t;return null!==(t=this.elementProperties.get(e))&&void 0!==t?t:O}static _$Ei(){if(this.hasOwnProperty(I("elementProperties")))return;const e=g(this);e.finalize(),void 0!==e.l&&(this.l=[...e.l]),this.elementProperties=new Map(e.elementProperties)}static finalize(){if(this.hasOwnProperty(I("finalized")))return;if(this.finalized=!0,this._$Ei(),this.hasOwnProperty(I("properties"))){const e=this.properties,t=[...h(e),...b(e)];for(const n of t)this.createProperty(n,e[n])}const e=this[Symbol.metadata];if(null!==e){const t=litPropertyMetadata.get(e);if(void 0!==t)for(const[e,n]of t)this.elementProperties.set(e,n)}this._$Eh=new Map;for(const[e,t]of this.elementProperties){const n=this._$Eu(e,t);void 0!==n&&this._$Eh.set(n,e)}this.elementStyles=this.finalizeStyles(this.styles)}static finalizeStyles(e){const t=[];if(Array.isArray(e)){const n=new Set(e.flat(1/0).reverse());for(const e of n)t.unshift(d(e))}else void 0!==e&&t.push(d(e));return t}static _$Eu(e,t){const n=t.attribute;return!1===n?void 0:"string"==typeof n?n:"string"==typeof e?e.toLowerCase():void 0}constructor(){super(),this._$Ep=void 0,this.isUpdatePending=!1,this.hasUpdated=!1,this._$Em=null,this._$Ev()}_$Ev(){var e;this._$Eg=new Promise(e=>this.enableUpdating=e),this._$AL=new Map,this._$E_(),this.requestUpdate(),null===(e=this.constructor.l)||void 0===e||e.forEach(e=>e(this))}addController(e){var t,n;(null!==(t=this._$ES)&&void 0!==t?t:this._$ES=[]).push(e),void 0!==this.renderRoot&&this.isConnected&&(null===(n=e.hostConnected)||void 0===n||n.call(e))}removeController(e){var t;null===(t=this._$ES)||void 0===t||t.splice(this._$ES.indexOf(e)>>>0,1)}_$E_(){const e=new Map,t=this.constructor.elementProperties;for(const n of t.keys())this.hasOwnProperty(n)&&(e.set(n,this[n]),delete this[n]);e.size>0&&(this._$Ep=e)}createRenderRoot(){var e;const t=null!==(e=this.shadowRoot)&&void 0!==e?e:this.attachShadow(this.constructor.shadowRootOptions);return((e,t)=>{if(i)e.adoptedStyleSheets=t.map(e=>e instanceof CSSStyleSheet?e:e.styleSheet);else for(const n of t){const t=document.createElement("style"),i=a.litNonce;void 0!==i&&t.setAttribute("nonce",i),t.textContent=n.cssText,e.appendChild(t)}})(t,this.constructor.elementStyles),t}connectedCallback(){var e,t;null!==(e=this.renderRoot)&&void 0!==e||(this.renderRoot=this.createRenderRoot()),this.enableUpdating(!0),null===(t=this._$ES)||void 0===t||t.forEach(e=>{var t;return null===(t=e.hostConnected)||void 0===t?void 0:t.call(e)})}enableUpdating(e){}disconnectedCallback(){var e;null===(e=this._$ES)||void 0===e||e.forEach(e=>{var t;return null===(t=e.hostDisconnected)||void 0===t?void 0:t.call(e)})}attributeChangedCallback(e,t,n){this._$AK(e,n)}_$EO(e,t){const n=this.constructor.elementProperties.get(e),a=this.constructor._$Eu(e,n);if(void 0!==a&&!0===n.reflect){var i;const r=(void 0!==(null===(i=n.converter)||void 0===i?void 0:i.toAttribute)?n.converter:x).toAttribute(t,n.type);this._$Em=e,null==r?this.removeAttribute(a):this.setAttribute(a,r),this._$Em=null}}_$AK(e,t){const n=this.constructor,a=n._$Eh.get(e);if(void 0!==a&&this._$Em!==a){var i;const e=n.getPropertyOptions(a),r="function"==typeof e.converter?{fromAttribute:e.converter}:void 0!==(null===(i=e.converter)||void 0===i?void 0:i.fromAttribute)?e.converter:x;this._$Em=a,this[a]=r.fromAttribute(t,e.type),this._$Em=null}}requestUpdate(e,t,n,a=!1,i){if(void 0!==e){var r,o;if(null!==(r=n)&&void 0!==r||(n=this.constructor.getPropertyOptions(e)),!(null!==(o=n.hasChanged)&&void 0!==o?o:C)(a?i:this[e],t))return;this.C(e,t,n)}!1===this.isUpdatePending&&(this._$Eg=this._$EP())}C(e,t,n){var a;this._$AL.has(e)||this._$AL.set(e,t),!0===n.reflect&&this._$Em!==e&&(null!==(a=this._$Ej)&&void 0!==a?a:this._$Ej=new Set).add(e)}async _$EP(){this.isUpdatePending=!0;try{await this._$Eg}catch(e){Promise.reject(e)}const e=this.scheduleUpdate();return null!=e&&await e,!this.isUpdatePending}scheduleUpdate(){return this.performUpdate()}performUpdate(){if(!this.isUpdatePending)return;if(!this.hasUpdated){if(this._$Ep){for(const[e,t]of this._$Ep)this[e]=t;this._$Ep=void 0}const e=this.constructor.elementProperties;if(e.size>0)for(const[t,n]of e)!0!==n.wrapped||this._$AL.has(t)||void 0===this[t]||this.C(t,this[t],n)}let e=!1;const t=this._$AL;try{var n;e=this.shouldUpdate(t),e?(this.willUpdate(t),null!==(n=this._$ES)&&void 0!==n&&n.forEach(e=>{var t;return null===(t=e.hostUpdate)||void 0===t?void 0:t.call(e)}),this.update(t)):this._$ET()}catch(t){throw e=!1,this._$ET(),t}e&&this._$AE(t)}willUpdate(e){}_$AE(e){var t;null!==(t=this._$ES)&&void 0!==t&&t.forEach(e=>{var t;return null===(t=e.hostUpdated)||void 0===t?void 0:t.call(e)}),this.hasUpdated||(this.hasUpdated=!0,this.firstUpdated(e)),this.updated(e)}_$ET(){this._$AL=new Map,this.isUpdatePending=!1}get updateComplete(){return this.getUpdateComplete()}getUpdateComplete(){return this._$Eg}shouldUpdate(e){return!0}update(e){this._$Ej&&(this._$Ej=this._$Ej.forEach(e=>this._$EO(e,this[e]))),this._$ET()}updated(e){}firstUpdated(e){}}w.elementStyles=[],w.shadowRootOptions={mode:"open"},w[I("elementProperties")]=new Map,w[I("finalized")]=new Map,null!=D&&D({ReactiveElement:w}),(null!==(f=v.reactiveElementVersions)&&void 0!==f?f:v.reactiveElementVersions=[]).push("2.0.0")},"q/PQ":function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("/49y");const i=e=>{try{const t=new URL(e).origin;if(a.b.has(t))return t}catch(e){return}}},q5bN:function(e,t,n){"use strict";n.d(t,"a",function(){return p}),n.d(t,"b",function(){return M});var a=n("P4Ao"),i=n("oZuh");const r=new Map;"metadata"in Reflect||(Reflect.metadata=function(e,t){return function(n){Reflect.defineMetadata(e,t,n)}},Reflect.defineMetadata=function(e,t,n){let a=r.get(n);void 0===a&&r.set(n,a=new Map),a.set(e,t)},Reflect.getOwnMetadata=function(e,t){const n=r.get(t);if(void 0!==n)return n.get(e)});class o{constructor(e,t){this.container=e,this.key=t}instance(e){return this.registerResolver(0,e)}singleton(e){return this.registerResolver(1,e)}transient(e){return this.registerResolver(2,e)}callback(e){return this.registerResolver(3,e)}cachedCallback(e){return this.registerResolver(3,k(e))}aliasTo(e){return this.registerResolver(5,e)}registerResolver(e,t){const{container:n,key:a}=this;return this.container=this.key=void 0,n.registerResolver(a,new v(a,e,t))}}function s(e){const t=e.slice(),n=Object.keys(e),a=n.length;let i;for(let r=0;r<a;++r)i=n[r],N(i)||(t[i]=e[i]);return t}const c=Object.freeze({none(e){throw Error(`${e.toString()} not registered, did you forget to add @singleton()?`)},singleton:e=>new v(e,1,e),transient:e=>new v(e,2,e)}),d=Object.freeze({default:Object.freeze({parentLocator:()=>null,responsibleForOwnerRequests:!1,defaultResolver:c.singleton})}),l=new Map;function u(e){return t=>Reflect.getOwnMetadata(e,t)}let f=null;const p=Object.freeze({createContainer:e=>new A(null,Object.assign({},d.default,e)),findResponsibleContainer(e){const t=e.$$container$$;return t&&t.responsibleForOwnerRequests?t:p.findParentContainer(e)},findParentContainer(e){const t=new CustomEvent(w,{bubbles:!0,composed:!0,cancelable:!0,detail:{container:void 0}});return e.dispatchEvent(t),t.detail.container||p.getOrCreateDOMContainer()},getOrCreateDOMContainer:(e,t)=>e?e.$$container$$||new A(e,Object.assign({},d.default,t,{parentLocator:p.findParentContainer})):f||(f=new A(null,Object.assign({},d.default,t,{parentLocator:()=>null}))),getDesignParamtypes:u("design:paramtypes"),getAnnotationParamtypes:u("di:paramtypes"),getOrCreateAnnotationParamTypes(e){let t=this.getAnnotationParamtypes(e);return void 0===t&&Reflect.defineMetadata("di:paramtypes",t=[],e),t},getDependencies(e){let t=l.get(e);if(void 0===t){const n=e.inject;if(void 0===n){const n=p.getDesignParamtypes(e),a=p.getAnnotationParamtypes(e);if(void 0===n)if(void 0===a){const n=Object.getPrototypeOf(e);t="function"==typeof n&&n!==Function.prototype?s(p.getDependencies(n)):[]}else t=s(a);else if(void 0===a)t=s(n);else{t=s(n);let e,i=a.length;for(let n=0;n<i;++n)e=a[n],void 0!==e&&(t[n]=e);const r=Object.keys(a);let o;i=r.length;for(let e=0;e<i;++e)o=r[e],N(o)||(t[o]=a[o])}}else t=s(n);l.set(e,t)}return t},defineProperty(e,t,n,i=!1){const r=`$di_${t}`;Reflect.defineProperty(e,t,{get:function(){let e=this[r];if(void 0===e){const o=this instanceof HTMLElement?p.findResponsibleContainer(this):p.getOrCreateDOMContainer();if(e=o.get(n),this[r]=e,i&&this instanceof a.a){const a=this.$fastController,i=()=>{p.findResponsibleContainer(this).get(n)!==this[r]&&(this[r]=e,a.notify(t))};a.subscribe({handleChange:i},"isConnected")}}return e}})},createInterface(e,t){const n="function"==typeof e?e:t,a="string"==typeof e?e:e&&"friendlyName"in e&&e.friendlyName||U,i="string"!=typeof e&&(e&&"respectConnection"in e&&e.respectConnection||!1),r=function(e,t,n){if(null==e||void 0!==new.target)throw new Error(`No registration for interface: '${r.friendlyName}'`);t?p.defineProperty(e,t,r,i):p.getOrCreateAnnotationParamTypes(e)[n]=r};return r.$isInterface=!0,r.friendlyName=null==a?"(anonymous)":a,null!=n&&(r.register=function(e,t){return n(new o(e,null!=t?t:r))}),r.toString=function(){return`InterfaceSymbol<${r.friendlyName}>`},r},inject:(...e)=>function(t,n,a){if("number"==typeof a){const n=p.getOrCreateAnnotationParamTypes(t),i=e[0];void 0!==i&&(n[a]=i)}else if(n)p.defineProperty(t,n,e[0]);else{const n=a?p.getOrCreateAnnotationParamTypes(a.value):p.getOrCreateAnnotationParamTypes(t);let i;for(let t=0;t<e.length;++t)i=e[t],void 0!==i&&(n[t]=i)}},transient:e=>(e.register=function(t){return M.transient(e,e).register(t)},e.registerInRequestor=!1,e),singleton:(e,t=h)=>(e.register=function(t){return M.singleton(e,e).register(t)},e.registerInRequestor=t.scoped,e)}),m=p.createInterface("Container");function _(e){return function(t){const n=function(e,t,a){p.inject(n)(e,t,a)};return n.$isResolver=!0,n.resolve=function(n,a){return e(t,n,a)},n}}p.inject;const h={scoped:!1};function b(e,t,n){p.inject(b)(e,t,n)}function g(e,t){return t.getFactory(e).construct(t)}_((e,t,n)=>()=>n.get(e)),_((e,t,n)=>n.has(e,!0)?n.get(e):void 0),b.$isResolver=!0,b.resolve=()=>{},_((e,t,n)=>{const a=g(e,t),i=new v(e,0,a);return n.registerResolver(e,i),a}),_((e,t,n)=>g(e,t));class v{constructor(e,t,n){this.key=e,this.strategy=t,this.state=n,this.resolving=!1}get $isResolver(){return!0}register(e){return e.registerResolver(this.key,this)}resolve(e,t){switch(this.strategy){case 0:return this.state;case 1:if(this.resolving)throw new Error(`Cyclic dependency found: ${this.state.name}`);return this.resolving=!0,this.state=e.getFactory(this.state).construct(t),this.strategy=0,this.resolving=!1,this.state;case 2:{const n=e.getFactory(this.state);if(null===n)throw new Error(`Resolver for ${String(this.key)} returned a null factory`);return n.construct(t)}case 3:return this.state(e,t,this);case 4:return this.state[0].resolve(e,t);case 5:return t.get(this.state);default:throw new Error(`Invalid resolver strategy specified: ${this.strategy}.`)}}getFactory(e){var t,n,a;switch(this.strategy){case 1:case 2:return e.getFactory(this.state);case 5:return null!==(a=null===(n=null===(t=e.getResolver(this.state))||void 0===t?void 0:t.getFactory)||void 0===n?void 0:n.call(t,e))&&void 0!==a?a:null;default:return null}}}function y(e){return this.get(e)}function S(e,t){return t(e)}class D{constructor(e,t){this.Type=e,this.dependencies=t,this.transformers=null}construct(e,t){let n;return n=void 0===t?new this.Type(...this.dependencies.map(y,e)):new this.Type(...this.dependencies.map(y,e),...t),null==this.transformers?n:this.transformers.reduce(S,n)}registerTransformer(e){(this.transformers||(this.transformers=[])).push(e)}}const I={$isResolver:!0,resolve:(e,t)=>t};function x(e){return"function"==typeof e.register}function C(e){return function(e){return x(e)&&"boolean"==typeof e.registerInRequestor}(e)&&e.registerInRequestor}const O=new Set(["Array","ArrayBuffer","Boolean","DataView","Date","Error","EvalError","Float32Array","Float64Array","Function","Int8Array","Int16Array","Int32Array","Map","Number","Object","Promise","RangeError","ReferenceError","RegExp","Set","SharedArrayBuffer","String","SyntaxError","TypeError","Uint8Array","Uint8ClampedArray","Uint16Array","Uint32Array","URIError","WeakMap","WeakSet"]),w="__DI_LOCATE_PARENT__",E=new Map;class A{constructor(e,t){this.owner=e,this.config=t,this._parent=void 0,this.registerDepth=0,this.context=null,null!==e&&(e.$$container$$=this),this.resolvers=new Map,this.resolvers.set(m,I),e instanceof Node&&e.addEventListener(w,e=>{e.composedPath()[0]!==this.owner&&(e.detail.container=this,e.stopImmediatePropagation())})}get parent(){return void 0===this._parent&&(this._parent=this.config.parentLocator(this.owner)),this._parent}get depth(){return null===this.parent?0:this.parent.depth+1}get responsibleForOwnerRequests(){return this.config.responsibleForOwnerRequests}registerWithContext(e,...t){return this.context=e,this.register(...t),this.context=null,this}register(...e){if(100==++this.registerDepth)throw new Error("Unable to autoregister dependency");let t,n,a,i,r;const o=this.context;for(let s=0,c=e.length;s<c;++s)if(t=e[s],F(t))if(x(t))t.register(this,o);else if(void 0!==t.prototype)M.singleton(t,t).register(this);else for(n=Object.keys(t),i=0,r=n.length;i<r;++i)a=t[n[i]],F(a)&&(x(a)?a.register(this,o):this.register(a));return--this.registerDepth,this}registerResolver(e,t){P(e);const n=this.resolvers,a=n.get(e);return null==a?n.set(e,t):a instanceof v&&4===a.strategy?a.state.push(t):n.set(e,new v(e,4,[a,t])),t}registerTransformer(e,t){const n=this.getResolver(e);if(null==n)return!1;if(n.getFactory){const e=n.getFactory(this);return null!=e&&(e.registerTransformer(t),!0)}return!1}getResolver(e,t=!0){if(P(e),void 0!==e.resolve)return e;let n,a=this;for(;null!=a;){if(n=a.resolvers.get(e),null!=n)return n;if(null==a.parent){const n=C(e)?this:a;return t?this.jitRegister(e,n):null}a=a.parent}return null}has(e,t=!1){return!!this.resolvers.has(e)||!(!t||null==this.parent)&&this.parent.has(e,!0)}get(e){if(P(e),e.$isResolver)return e.resolve(this,this);let t,n=this;for(;null!=n;){if(t=n.resolvers.get(e),null!=t)return t.resolve(n,this);if(null==n.parent){const a=C(e)?this:n;return t=this.jitRegister(e,a),t.resolve(n,this)}n=n.parent}throw new Error(`Unable to resolve key: ${String(e)}`)}getAll(e,t=!1){P(e);const n=this;let a,r=n;if(t){let t=i.d;for(;null!=r;)a=r.resolvers.get(e),null!=a&&(t=t.concat(T(a,r,n))),r=r.parent;return t}for(;null!=r;){if(a=r.resolvers.get(e),null!=a)return T(a,r,n);if(r=r.parent,null==r)return i.d}return i.d}getFactory(e){let t=E.get(e);if(void 0===t){if(H(e))throw new Error(`${e.name} is a native function and therefore cannot be safely constructed by DI. If this is intentional, please use a callback or cachedCallback resolver.`);E.set(e,t=new D(e,p.getDependencies(e)))}return t}registerFactory(e,t){E.set(e,t)}createChild(e){return new A(null,Object.assign({},this.config,e,{parentLocator:()=>this}))}jitRegister(e,t){if("function"!=typeof e)throw new Error(`Attempted to jitRegister something that is not a constructor: '${e}'. Did you forget to register this dependency?`);if(O.has(e.name))throw new Error(`Attempted to jitRegister an intrinsic type: ${e.name}. Did you forget to add @inject(Key)`);if(x(e)){const n=e.register(t);if(!(n instanceof Object)||null==n.resolve){const n=t.resolvers.get(e);if(null!=n)return n;throw new Error("A valid resolver was not returned from the static register method")}return n}if(e.$isInterface)throw new Error(`Attempted to jitRegister an interface: ${e.friendlyName}`);{const n=this.config.defaultResolver(e,t);return t.resolvers.set(e,n),n}}}const L=new WeakMap;function k(e){return function(t,n,a){if(L.has(a))return L.get(a);const i=e(t,n,a);return L.set(a,i),i}}const M=Object.freeze({instance:(e,t)=>new v(e,0,t),singleton:(e,t)=>new v(e,1,t),transient:(e,t)=>new v(e,2,t),callback:(e,t)=>new v(e,3,t),cachedCallback:(e,t)=>new v(e,3,k(t)),aliasTo:(e,t)=>new v(t,5,e)});function P(e){if(null==e)throw new Error("key/value cannot be null or undefined. Are you trying to inject/register something that doesn't exist with DI?")}function T(e,t,n){if(e instanceof v&&4===e.strategy){const a=e.state;let i=a.length;const r=new Array(i);for(;i--;)r[i]=a[i].resolve(t,n);return r}return[e.resolve(t,n)]}const U="(anonymous)";function F(e){return"object"==typeof e&&null!==e||"function"==typeof e}const H=function(){const e=new WeakMap;let t=!1,n="",a=0;return function(i){return t=e.get(i),void 0===t&&(n=i.toString(),a=n.length,t=a>=29&&a<=100&&125===n.charCodeAt(a-1)&&n.charCodeAt(a-2)<=32&&93===n.charCodeAt(a-3)&&101===n.charCodeAt(a-4)&&100===n.charCodeAt(a-5)&&111===n.charCodeAt(a-6)&&99===n.charCodeAt(a-7)&&32===n.charCodeAt(a-8)&&101===n.charCodeAt(a-9)&&118===n.charCodeAt(a-10)&&105===n.charCodeAt(a-11)&&116===n.charCodeAt(a-12)&&97===n.charCodeAt(a-13)&&110===n.charCodeAt(a-14)&&88===n.charCodeAt(a-15),e.set(i,t)),t}}(),R={};function N(e){switch(typeof e){case"number":return e>=0&&(0|e)===e;case"string":{const t=R[e];if(void 0!==t)return t;const n=e.length;if(0===n)return R[e]=!1;let a=0;for(let t=0;t<n;++t)if(a=e.charCodeAt(t),0===t&&48===a&&n>1||a<48||a>57)return R[e]=!1;return R[e]=!0}default:return!1}}},qibw:function(e,t,n){"use strict";var a;n.d(t,"a",function(){return a}),function(e){e.ltr="ltr",e.rtl="rtl"}(a||(a={}))},qqMp:function(e,t,n){"use strict";n.d(t,"b",function(){return r.b}),n.d(t,"c",function(){return o.b}),n.d(t,"d",function(){return o.d}),n.d(t,"a",function(){return s});var a,i,r=n("ptXv"),o=n("5rba");class s extends r.a{constructor(){super(...arguments),this.renderOptions={host:this},this._$Do=void 0}createRenderRoot(){var e,t;const n=super.createRenderRoot();return null!==(t=(e=this.renderOptions).renderBefore)&&void 0!==t||(e.renderBefore=n.firstChild),n}update(e){const t=this.render();this.hasUpdated||(this.renderOptions.isConnected=this.isConnected),super.update(e),this._$Do=Object(o.e)(t,this.renderRoot,this.renderOptions)}connectedCallback(){var e;super.connectedCallback(),null===(e=this._$Do)||void 0===e||e.setConnected(!0)}disconnectedCallback(){var e;super.disconnectedCallback(),null===(e=this._$Do)||void 0===e||e.setConnected(!1)}render(){return o.c}}s._$litElement$=!0,s.finalized=!0,null===(a=globalThis.litElementHydrateSupport)||void 0===a||a.call(globalThis,{LitElement:s});const c=globalThis.litElementPolyfillSupport;null==c||c({LitElement:s}),(null!==(i=globalThis.litElementVersions)&&void 0!==i?i:globalThis.litElementVersions=[]).push("4.0.0")},qwWh:function(e,t,n){"use strict";n.d(t,"a",function(){return a.a}),n.d(t,"l",function(){return i.a}),n.d(t,"m",function(){return i.b}),n.d(t,"g",function(){return r.a}),n.d(t,"B",function(){return r.b}),n.d(t,"b",function(){return o.a}),n.d(t,"e",function(){return s.a}),n.d(t,"n",function(){return s.b}),n.d(t,"o",function(){return u}),n.d(t,"p",function(){return f.a}),n.d(t,"D",function(){return p.a}),n.d(t,"i",function(){return m.a}),n.d(t,"k",function(){return m.b}),n.d(t,"s",function(){return m.c}),n.d(t,"t",function(){return _.a}),n.d(t,"u",function(){return _.b}),n.d(t,"v",function(){return h}),n.d(t,"E",function(){return b.b}),n.d(t,"c",function(){return b.a}),n.d(t,"d",function(){return g.a}),n.d(t,"f",function(){return v.a}),n.d(t,"F",function(){return y.b}),n.d(t,"y",function(){return y.a}),n.d(t,"A",function(){return S.a}),n.d(t,"J",function(){return D.a}),n.d(t,"L",function(){return I.a}),n.d(t,"w",function(){return x.a}),n.d(t,"x",function(){return C.a}),n.d(t,"h",function(){return O.a}),n.d(t,"j",function(){return w.a}),n.d(t,"I",function(){return E.a}),n.d(t,"C",function(){return L}),n.d(t,"H",function(){return A.b}),n.d(t,"M",function(){return A.c}),n.d(t,"G",function(){return A.a}),n.d(t,"z",function(){return k.a}),n.d(t,"K",function(){return k.b}),n.d(t,"r",function(){return M.a}),n.d(t,"q",function(){return N});var a=n("4Eu1"),i=n("/49y"),r=n("SUtl"),o=n("pnUA"),s=n("dblZ"),c=n("ZgG/"),d=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},l=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)};class u extends s.b{constructor(){super(...arguments),this.stateChangedHandler=()=>{this.fireCustomEvent("onStateChanged",this.provider.state)}}get provider(){return this._provider}set provider(e){this._provider&&this.provider.removeStateChangedHandler(()=>this.stateChangedHandler),this._provider=e,this._provider&&this.provider.onStateChanged(()=>this.stateChangedHandler)}get isAvailable(){return!0}firstUpdated(e){super.firstUpdated(e);let t=!1;if(this.dependsOn){let e=this.dependsOn;for(;e;){if(e.isAvailable){t=!0;break}e=e.dependsOn}}!t&&this.isAvailable&&this.initializeProvider()}initializeProvider(){}}d([Object(c.b)({attribute:"depends-on",converter:e=>document.querySelector(e),type:String}),l("design:type",u)],u.prototype,"dependsOn",void 0),d([Object(c.b)({attribute:"base-url",type:String}),l("design:type",String)],u.prototype,"baseUrl",void 0),d([Object(c.b)({attribute:"custom-hosts",type:String,converter:e=>e.split(",").map(e=>e.trim())}),l("design:type",Array)],u.prototype,"customHosts",void 0);var f=n("Y1A4"),p=n("HN6m"),m=n("/i08"),_=n("zFbe");class h extends m.a{get name(){return"MgtSimpleProvider"}constructor(e,t,n){super(),this._getAccessTokenHandler=e,this._loginHandler=t,this._logoutHandler=n,this.graph=Object(r.b)(this)}getAccessToken(e){return this._getAccessTokenHandler(null==e?void 0:e.scopes)}login(){return this._loginHandler()}logout(){return this._logoutHandler()}}var b=n("DNu6"),g=n("d3jT"),v=n("x+GM"),y=n("ylMe"),S=n("E7Zv"),D=n("RIOo"),I=n("q/PQ"),x=n("Yn9m"),C=n("oaDa"),O=n("WHff"),w=n("zb6c"),E=n("cBsD"),A=n("R2TB");const L=e=>{const t=`${p.a.prefix}-${e}`,n=customElements.get(t),a=e=>e.packageVersion||" Unknown likely <3.0.0";return n?e=>(Object(A.a)(`Tag name ${t} is already defined using class ${n.name} version ${a(n)}\n`,`Currently registering class ${e.name} with version ${a(e)}\n`,"Please use the disambiguation feature to define a unique tag name for this component see: https://github.com/microsoftgraph/microsoft-graph-toolkit/tree/main/packages/mgt-components#disambiguation"),e):Object(c.a)(t)};var k=n("0mOt"),M=n("2oY9"),P=n("EYXE");class T{constructor(){this.session=window.sessionStorage}setItem(e,t){this.session.setItem(e,t)}getItem(e){return this.session.getItem(e)}clear(){this.session.clear()}}var U=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};class F{static get _sessionCache(){return!this._cache&&(e=>{let t;try{t=window.sessionStorage;const e="__storage_test__";return t.setItem(e,e),t.removeItem(e),!0}catch(e){return e instanceof DOMException&&(22===e.code||1014===e.code||"QuotaExceededError"===e.name||"NS_ERROR_DOM_QUOTA_REACHED"===e.name)&&t&&0!==t.length}})()&&(this._cache=new T),this._cache}execute(e){return U(this,void 0,void 0,function*(){try{const t=yield F.getBaseUrl();e.request=t+encodeURIComponent(e.request)}catch(e){}return yield this._nextMiddleware.execute(e)})}setNext(e){this._nextMiddleware=e}static getBaseUrl(){var e,t;return U(this,void 0,void 0,function*(){if(!this._baseUrl){const n=null===(e=this._sessionCache)||void 0===e?void 0:e.getItem("endpointURL");if(n)this._baseUrl=n;else{try{const e=yield fetch("https://cdn.graph.office.net/en-us/graph/api/proxy/endpoint"),t=yield e.json();"string"!=typeof t?F.setBaseFallbackUrl():this._baseUrl=t+"?url="}catch(e){F.setBaseFallbackUrl()}null===(t=this._sessionCache)||void 0===t||t.setItem("endpointURL",this._baseUrl)}}return this._baseUrl})}static setBaseFallbackUrl(){this._baseUrl="https://proxy.apisandbox.msdn.microsoft.com/svc?url="}}class H extends r.a{static create(e){return t=this,void 0,a=function*(){const t=[new P.AuthenticationHandler(e),new P.RetryHandler(new P.RetryHandlerOptions),new P.TelemetryHandler,new F,new P.HTTPMessageHandler];return new H(P.Client.initWithMiddleware({middleware:Object(S.a)(...t),customHosts:new Set([new URL(yield F.getBaseUrl()).hostname])}))},new((n=void 0)||(n=Promise))(function(e,i){function r(e){try{s(a.next(e))}catch(e){i(e)}}function o(e){try{s(a.throw(e))}catch(e){i(e)}}function s(t){var a;t.done?e(t.value):(a=t.value,a instanceof n?a:new n(function(e){e(a)})).then(r,o)}s((a=a.apply(t,[])).next())});var t,n,a}forComponent(e){return this}}var R=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};class N extends m.a{constructor(e=!1,t=[]){super(),this._accounts=[],this._mockGraphPromise=H.create(this);const n=Boolean(t.length);this.isMultipleAccountSupported=n,this.isMultipleAccountDisabled=!n,this._accounts=t,this.initializeMockGraph(e)}get isMultiAccountSupportedAndEnabled(){return!this.isMultipleAccountDisabled&&this.isMultipleAccountSupported}getAllAccounts(){return this._accounts}getActiveAccount(){if(this._accounts.length)return this._accounts[0]}login(){return R(this,void 0,void 0,function*(){this.setState(m.c.Loading),yield this._mockGraphPromise,yield new Promise(e=>setTimeout(e,3e3)),this.setState(m.c.SignedIn)})}logout(){return R(this,void 0,void 0,function*(){this.setState(m.c.Loading),yield this._mockGraphPromise,yield new Promise(e=>setTimeout(e,3e3)),this.setState(m.c.SignedOut)})}getAccessToken(){return Promise.resolve("{token:https://graph.microsoft.com/}")}get name(){return"MgtMockProvider"}initializeMockGraph(e=!1){return R(this,void 0,void 0,function*(){this.graph=yield this._mockGraphPromise,e?this.setState(m.c.SignedIn):this.setState(m.c.SignedOut)})}}},s2f2:function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("Kct4");function i(e){return Object(a.a)(e)?-1:1}},s9K1:function(e,t,n){"use strict";n.d(t,"g",function(){return c}),n.d(t,"d",function(){return d}),n.d(t,"h",function(){return l}),n.d(t,"e",function(){return u}),n.d(t,"f",function(){return f}),n.d(t,"j",function(){return p}),n.d(t,"i",function(){return m}),n.d(t,"b",function(){return _}),n.d(t,"a",function(){return h}),n.d(t,"c",function(){return b});var a=n("DNu6"),i=n("RIOo"),r=n("imsm"),o=n("z0DP"),s=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const c=()=>a.a.config.users.invalidationPeriod||a.a.config.defaultInvalidationPeriod,d=()=>a.a.config.users.isEnabled&&a.a.config.isEnabled,l=(e,t="",n=10)=>s(void 0,void 0,void 0,function*(){let o;const s=`${""===t?"*":t}:${n}`,l={maxResults:n,results:null};if(d()){o=a.a.getCache(r.a.users,r.a.users.stores.userFilters);const e=yield o.getValue(s);if(e&&c()>Date.now()-e.timeCached)return e.results.map(e=>JSON.parse(e))}const u=e.api("/users").top(n);t&&u.filter(t);try{const e=yield u.middlewareOptions(Object(i.a)("user.read")).get();return d()&&e&&(l.results=e.value.map(e=>JSON.stringify(e)),yield o.putValue(t,l)),e.value}catch(e){}}),u=(e,t)=>s(void 0,void 0,void 0,function*(){let n;if(d()){n=a.a.getCache(r.a.users,r.a.users.stores.users);const e=yield n.getValue("me");if(e&&c()>Date.now()-e.timeCached){const n=JSON.parse(e.user),a=t?t.filter(e=>!Object.keys(n).includes(e)):null;if(!a||a.length<=1)return n}}let o="me";t&&(o=o+"?$select="+t.toString());const s=yield e.api(o).middlewareOptions(Object(i.a)("user.read")).get();return d()&&(yield n.putValue("me",{user:JSON.stringify(s)})),s}),f=(e,t,n)=>s(void 0,void 0,void 0,function*(){let o;if(d()){o=a.a.getCache(r.a.users,r.a.users.stores.users);const e=yield o.getValue(t);if(e&&c()>Date.now()-e.timeCached){const t=e.user?JSON.parse(e.user):null,a=n&&t?n.filter(e=>!Object.keys(t).includes(e)):null;if(!a||a.length<=1)return t}}let s,l=`/users/${t}`;n&&(l=l+"?$select="+n.toString());try{s=yield e.api(l).middlewareOptions(Object(i.a)("user.readbasic.all")).get()}catch(e){}return d()&&(yield o.putValue(t,{user:JSON.stringify(s)})),s}),p=(e,t,n="",i="",o)=>s(void 0,void 0,void 0,function*(){if(!t||0===t.length)return[];const l=e.createBatch(),p={},m={},_=[];let h;n=n.toLowerCase(),d()&&(h=a.a.getCache(r.a.users,r.a.users.stores.users));for(const a of t){p[a]=null;let t,r,o=`/users/${a}`;if(d()&&(r=yield h.getValue(a)),(null==r?void 0:r.user)&&c()>Date.now()-r.timeCached)if(t=JSON.parse(null==r?void 0:r.user),n){if(t){const e=t.displayName;(null==e?void 0:e.toLowerCase().includes(n))&&(m[a]=t)}}else t?p[a]=t:(l.get(a,o,["user.readbasic.all"]),_.push(a));else""!==a&&("me"===a.toString()?p[a]=yield u(e):(o=`/users/${a}`,i&&(o+=`${o}?$filter=${i}`),l.get(a,o,["user.readbasic.all"]),_.push(a)))}try{if(l.hasRequests){const e=yield l.executeAll();for(const a of t){const t=e.get(a);if(null==t?void 0:t.content){const e=t.content;n?((null==e?void 0:e.displayName.toLowerCase())||"").includes(n)&&(m[a]=e):p[a]=e,d()&&(yield h.putValue(a,{user:JSON.stringify(e)}))}else{const e=o.find(e=>Object.values(e).includes(a));e&&(p[a]=e)}}}return n&&Object.keys(m).length?Promise.all(Object.values(m)):Promise.all(Object.values(p))}catch(n){try{return t.filter(e=>_.includes(e)).forEach(t=>{p[t]=f(e,t)}),d()&&(yield Promise.all(t.filter(e=>_.includes(e)).map(e=>s(void 0,void 0,void 0,function*(){return yield h.putValue(e,{user:JSON.stringify(yield p[e])})})))),Promise.all(Object.values(p))}catch(e){return[]}}}),m=(e,t,n)=>s(void 0,void 0,void 0,function*(){var i;if(!t||0===t.length)return[];const l=e.createBatch(),u=[];let f,p;d()&&(p=a.a.getCache(r.a.users,r.a.users.stores.usersQuery));for(const e of t)if(d()&&(f=yield p.getValue(e)),d()&&(null==f?void 0:f.results[0])&&c()>Date.now()-f.timeCached){const e=JSON.parse(f.results[0]);u.push(e)}else l.get(e,`/me/people?$search="${e}"`,["people.read"],{"X-PeopleQuery-QuerySources":"Mailbox,Directory"});if(l.hasRequests)try{const e=yield l.executeAll();for(const a of t){const t=e.get(a);if((null===(i=null==t?void 0:t.content)||void 0===i?void 0:i.value)&&t.content.value.length>0)u.push(t.content.value[0]),d()&&(yield p.putValue(a,{maxResults:1,results:[JSON.stringify(t.content.value[0])]}));else{const e=n.find(e=>Object.values(e).includes(a));e&&u.push(e)}}return u}catch(n){try{return Promise.all(t.filter(e=>e&&""!==e).map(t=>s(void 0,void 0,void 0,function*(){const n=yield Object(o.d)(e,t,1);if(null==n?void 0:n.length)return d()&&(yield p.putValue(t,{maxResults:1,results:[JSON.stringify(n[0])]})),n[0]})))}catch(e){return[]}}return u}),_=(e,t,n=10,o="")=>s(void 0,void 0,void 0,function*(){const s={maxResults:n,results:null},l=`${t}:${n}:${o}`;let u;if(d()){u=a.a.getCache(r.a.users,r.a.users.stores.usersQuery);const e=yield u.getValue(l);if(e&&c()>Date.now()-e.timeCached)return e.results.map(e=>JSON.parse(e))}const f=`${t.replace(/#/g,"%2523")}`,p=e.api("users").header("ConsistencyLevel","eventual").count(!0).search(`"displayName:${f}" OR "mail:${f}"`);let m;""!==o&&p.filter(o);try{m=yield p.top(n).middlewareOptions(Object(i.a)("User.ReadBasic.All")).get()}catch(e){}return d()&&m&&(s.results=m.value.map(e=>JSON.stringify(e)),yield u.putValue(t,s)),m?m.value:null}),h=(e,t,n,l=10,u=o.a.person,f=!1,p="",m="")=>s(void 0,void 0,void 0,function*(){const s={maxResults:l,results:null};let _;const h=`${n||"*"}:${t||"*"}:${l}:${u}:${f}:${p}`;if(d()){_=a.a.getCache(r.a.users,r.a.users.stores.usersQuery);const e=yield _.getValue(h);if(e&&c()>Date.now()-e.timeCached)return e.results.map(e=>JSON.parse(e))}let b="";t&&(b=`startswith(displayName,'${t}') or startswith(givenName,'${t}') or startswith(surname,'${t}') or startswith(mail,'${t}') or startswith(userPrincipalName,'${t}')`);let g=`/groups/${n}/${f?"transitiveMembers":"members"}`;u===o.a.person?g+="/microsoft.graph.user":u===o.a.group&&(g+="/microsoft.graph.group",t&&(b=`startswith(displayName,'${t}') or startswith(mail,'${t}')`)),p&&(b+=t?` and ${p}`:p),m&&(b+=t?` and ${m}`:m);const v=yield e.api(g).count(!0).top(l).filter(b).header("ConsistencyLevel","eventual").middlewareOptions(Object(i.a)("GroupMember.Read.All")).get();return d()&&v&&(s.results=v.value.map(e=>JSON.stringify(e)),yield _.putValue(h,s)),v?v.value:null}),b=(e,t,n,a=10,i=o.a.person,r=!1,c="")=>s(void 0,void 0,void 0,function*(){const o=[];for(const s of n)try{const n=yield h(e,t,s,a,i,r,c);o.push(...n)}catch(e){continue}return o})},swXE:function(e,t,n){"use strict";n.d(t,"j",function(){return l}),n.d(t,"a",function(){return p}),n.d(t,"f",function(){return m}),n.d(t,"b",function(){return _}),n.d(t,"g",function(){return h}),n.d(t,"c",function(){return b}),n.d(t,"k",function(){return g}),n.d(t,"l",function(){return v}),n.d(t,"h",function(){return y}),n.d(t,"d",function(){return S}),n.d(t,"i",function(){return D}),n.d(t,"e",function(){return I});var a=n("PT2o"),i=n("tq/8"),r=n("RN7+"),o=n("SiT+"),s=n("6eT7"),c=n("xAa8"),d=n("0q6d");function l(e){function t(e){return e<=.03928?e/12.92:Math.pow((e+.055)/1.055,2.4)}return function(e){return.2126*e.r+.7152*e.g+.0722*e.b}(new s.a(t(e.r),t(e.g),t(e.b),1))}function u(e,t,n){return n-t==0?0:(e-t)/(n-t)}function f(e,t,n){return(u(e.r,t.r,n.r)+u(e.g,t.g,n.g)+u(e.b,t.b,n.b))/3}function p(e,t,n=null){let a=0,i=n;return null!==i?a=f(e,t,i):(i=new s.a(0,0,0,1),a=f(e,t,i),a<=0&&(i=new s.a(1,1,1,1),a=f(e,t,i))),a=Math.round(1e3*a)/1e3,new s.a(i.r,i.g,i.b,a)}function m(e){const t=Math.max(e.r,e.g,e.b),n=Math.min(e.r,e.g,e.b),i=t-n;let r=0;0!==i&&(r=t===e.r?(e.g-e.b)/i%6*60:t===e.g?60*((e.b-e.r)/i+2):60*((e.r-e.g)/i+4)),r<0&&(r+=360);const o=(t+n)/2;let s=0;return 0!==i&&(s=i/(1-Math.abs(2*o-1))),new a.a(r,s,o)}function _(e,t=1){const n=(1-Math.abs(2*e.l-1))*e.s,a=n*(1-Math.abs(e.h/60%2-1)),i=e.l-n/2;let r=0,o=0,c=0;return e.h<60?(r=n,o=a,c=0):e.h<120?(r=a,o=n,c=0):e.h<180?(r=0,o=n,c=a):e.h<240?(r=0,o=a,c=n):e.h<300?(r=a,o=0,c=n):e.h<360&&(r=n,o=0,c=a),new s.a(r+i,o+i,c+i,t)}function h(e){const t=Math.max(e.r,e.g,e.b),n=t-Math.min(e.r,e.g,e.b);let a=0;0!==n&&(a=t===e.r?(e.g-e.b)/n%6*60:t===e.g?60*((e.b-e.r)/n+2):60*((e.r-e.g)/n+4)),a<0&&(a+=360);let r=0;return 0!==t&&(r=n/t),new i.a(a,r,t)}function b(e,t=1){const n=e.s*e.v,a=n*(1-Math.abs(e.h/60%2-1)),i=e.v-n;let r=0,o=0,c=0;return e.h<60?(r=n,o=a,c=0):e.h<120?(r=a,o=n,c=0):e.h<180?(r=0,o=n,c=a):e.h<240?(r=0,o=a,c=n):e.h<300?(r=a,o=0,c=n):e.h<360&&(r=n,o=0,c=a),new s.a(r+i,o+i,c+i,t)}function g(e){function t(e){return e<=.04045?e/12.92:Math.pow((e+.055)/1.055,2.4)}const n=t(e.r),a=t(e.g),i=t(e.b),r=.4124564*n+.3575761*a+.1804375*i,o=.2126729*n+.7151522*a+.072175*i,s=.0193339*n+.119192*a+.9503041*i;return new c.a(r,o,s)}function v(e,t=1){function n(e){return e<=.0031308?12.92*e:1.055*Math.pow(e,1/2.4)-.055}const a=n(3.2404542*e.x-1.5371385*e.y-.4985314*e.z),i=n(-.969266*e.x+1.8760108*e.y+.041556*e.z),r=n(.0556434*e.x-.2040259*e.y+1.0572252*e.z);return new s.a(a,i,r,t)}function y(e){return function(e){function t(e){return e>r.a.epsilon?Math.pow(e,1/3):(r.a.kappa*e+16)/116}const n=t(e.x/c.a.whitePoint.x),a=t(e.y/c.a.whitePoint.y),i=116*a-16,o=500*(n-a),s=200*(a-t(e.z/c.a.whitePoint.z));return new r.a(i,o,s)}(g(e))}function S(e,t=1){return v(function(e){const t=(e.l+16)/116,n=t+e.a/500,a=t-e.b/200,i=Math.pow(n,3),o=Math.pow(t,3),s=Math.pow(a,3);let d=0;d=i>r.a.epsilon?i:(116*n-16)/r.a.kappa;let l=0;l=e.l>r.a.epsilon*r.a.kappa?o:e.l/r.a.kappa;let u=0;return u=s>r.a.epsilon?s:(116*a-16)/r.a.kappa,d=c.a.whitePoint.x*d,l=c.a.whitePoint.y*l,u=c.a.whitePoint.z*u,new c.a(d,l,u)}(e),t)}function D(e){return function(e){let t=0;(Math.abs(e.b)>.001||Math.abs(e.a)>.001)&&(t=Object(d.h)(Math.atan2(e.b,e.a))),t<0&&(t+=360);const n=Math.sqrt(e.a*e.a+e.b*e.b);return new o.a(e.l,n,t)}(y(e))}function I(e,t=1){return S(function(e){let t=0,n=0;return 0!==e.h&&(t=Math.cos(Object(d.b)(e.h))*e.c,n=Math.sin(Object(d.b)(e.h))*e.c),new r.a(e.l,t,n)}(e),t)}},tezw:function(e,t,n){"use strict";n.d(t,"a",function(){return s});var a=n("eftJ"),i=n("QBS5"),r=n("oePG"),o=n("TDEi");class s extends o.a{constructor(){super(...arguments),this.percentComplete=0}valueChanged(){this.$fastController.isConnected&&this.updatePercentComplete()}minChanged(){this.$fastController.isConnected&&this.updatePercentComplete()}maxChanged(){this.$fastController.isConnected&&this.updatePercentComplete()}connectedCallback(){super.connectedCallback(),this.updatePercentComplete()}updatePercentComplete(){const e="number"==typeof this.min?this.min:0,t="number"==typeof this.max?this.max:100,n="number"==typeof this.value?this.value:0,a=t-e;this.percentComplete=0===a?0:Math.fround((n-e)/a*100)}}Object(a.a)([Object(i.c)({converter:i.e})],s.prototype,"value",void 0),Object(a.a)([Object(i.c)({converter:i.e})],s.prototype,"min",void 0),Object(a.a)([Object(i.c)({converter:i.e})],s.prototype,"max",void 0),Object(a.a)([Object(i.c)({mode:"boolean"})],s.prototype,"paused",void 0),Object(a.a)([r.d],s.prototype,"percentComplete",void 0)},tnKu:function(e,t,n){"use strict";n.d(t,"b",function(){return l}),n.d(t,"a",function(){return u});var a=n("qqMp"),i=n("C001"),r=n("7Cdu"),o=n("/4Fm");const s=[a.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host .loading,:host .no-data{margin:0 20px;display:flex;justify-content:center}:host .no-data{font-style:normal;font-weight:600;font-size:14px;color:var(--font-color,#323130);line-height:19px}:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{position:relative;user-select:none}:host .root.compact{padding:0}:host .root .message{padding:8px 20px;display:flex;align-items:center;justify-content:space-between}:host .root .message:hover{background-color:var(--message-hover-color,var(--neutral-fill-hover));cursor:pointer}:host .root .message:last-child{margin-bottom:unset}:host .root .message .message__detail{min-width:0;line-height:normal}:host .root .message .message__detail .message__subject{color:var(--message-subject-color,var(--neutral-foreground-color));font-size:var(--message-subject-font-size,14px);font-weight:var(--message-subject-font-weight,600);line-height:var(--message-subject-line-height,20px)}:host .root .message .message__detail .message__from{font-size:var(--message-subject-font-size,12px);color:var(--message-from-color,var(--neutral-foreground-color));line-height:var(--message-subject-line-height,16px);font-weight:var(--message-subject-font-weight,400)}:host .root .message .message__detail .message__message{font-size:var(--message-subject-font-size,12px);color:var(--message-color,var(--neutral-foreground-hint));line-height:var(--message-subject-line-height,16px);font-weight:var(--message-subject-font-weight,400)}:host .root .message .message__detail>div{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}:host .root .message .message__date{margin-top:8px;font-size:12px;color:var(--message-date,var(--neutral-foreground-hint));margin-left:10px;white-space:nowrap}
`],c={emailsSectionTitle:"Emails"};var d=n("0mOt");const l=()=>Object(d.b)("messages",u);class u extends i.a{static get styles(){return s}get strings(){return c}constructor(e){super(),this._messages=e}get displayName(){return this.strings.emailsSectionTitle}get cardTitle(){return this.strings.emailsSectionTitle}clearState(){super.clearState(),this._messages=[]}renderIcon(){return Object(r.b)(r.a.Messages)}renderCompactView(){var e;let t;if(this.isLoadingState)t=this.renderLoading();else if(null===(e=this._messages)||void 0===e?void 0:e.length){const e=this._messages?this._messages.slice(0,3).map(e=>this.renderMessage(e)):[];t=a.c`
         ${e}
       `}else t=this.renderNoData();return a.c`
       <div class="root compact">
         ${t}
       </div>
     `}renderFullView(){var e;let t;return t=this.isLoadingState?this.renderLoading():(null===(e=this._messages)||void 0===e?void 0:e.length)?a.c`
         ${this._messages.slice(0,5).map(e=>this.renderMessage(e))}
       `:this.renderNoData(),a.c`
       <div class="root">
         ${t}
       </div>
     `}renderMessage(e){return a.c`
       <div class="message" @click=${()=>this.handleMessageClick(e)}>
         <div class="message__detail">
           <div class="message__subject">${e.subject}</div>
           <div class="message__from">${e.from.emailAddress.name}</div>
           <div class="message__message">${e.bodyPreview}</div>
         </div>
         <div class="message__date">${Object(o.l)(new Date(e.receivedDateTime))}</div>
       </div>
     `}handleMessageClick(e){window.open(e.webLink,"_blank","noreferrer")}}},"tq/8":function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("0q6d");class i{constructor(e,t,n){this.h=e,this.s=t,this.v=n}static fromObject(e){return!e||isNaN(e.h)||isNaN(e.s)||isNaN(e.v)?null:new i(e.h,e.s,e.v)}equalValue(e){return this.h===e.h&&this.s===e.s&&this.v===e.v}roundToPrecision(e){return new i(Object(a.i)(this.h,e),Object(a.i)(this.s,e),Object(a.i)(this.v,e))}toObject(){return{h:this.h,s:this.s,v:this.v}}}},uVfA:function(e,t,n){"use strict";n.r(t),n.d(t,"registerMgtPersonCardComponent",function(){return he}),n.d(t,"MgtPersonCard",function(){return be});var a=n("qqMp"),i=n("ZgG/"),r=n("h2QR"),o=n("Y1A4"),s=n("Yn9m"),c=n("cBsD"),d=n("zFbe"),l=n("/i08"),u=n("HN6m"),f=n("z0DP"),p=n("ZzBS"),m=n("glp4"),_=n("hgjj"),h=n("7Cdu"),b=n("zlIh"),g=n("RIOo"),v=n("YZrM"),y=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const S="businessPhones,companyName,department,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,id,accountEnabled";var D=n("7SX6");const I=[a.b`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size,14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{box-shadow:var(--person-card-box-shadow,var(--elevation-shadow-card-rest));width:var(--mgt-flyout-set-width,375px);overflow:hidden;user-select:none;border-radius:8px;background-color:var(--person-card-background-color,var(--neutral-layer-floating));--file-list-background-color:transparent;--file-list-box-shadow:none;--file-item-background-color:transparent}:host .small{max-height:100vh;overflow:hidden auto}:host .nav{height:0;position:relative;z-index:100}:host .nav__back{padding-top:18px;padding-inline-start:12px;height:32px;width:32px}:host .nav__back svg{fill:var(--person-card-icon-color,var(--accent-foreground-rest))}:host .nav__back:hover{cursor:pointer}:host .nav__back:hover svg{fill:var(--person-card-nav-back-arrow-hover-color,var(--neutral-foreground-rest))}:host .close-card-container{position:relative;z-index:100}:host .close-card-container .close-button{position:absolute;right:10px;margin-top:9px;z-index:1;border:1px solid transparent}:host .close-card-container .close-button svg path{fill:var(--person-card-icon-color,var(--accent-foreground-rest))}:host .close-card-container .close-button:hover svg path{fill:var(--person-card-close-button-hover-color,var(--neutral-foreground-rest))}:host .person-details-container{display:flex;flex-direction:column;padding-inline-start:22px;padding-block:36px;border-bottom:1px solid var(--person-card-border-color,var(--neutral-stroke-rest))}:host .person-details-container .person-image{--person-avatar-top-spacing:var(--person-card-avatar-top-spacing, 15px);--person-details-left-spacing:var(--person-card-details-left-spacing, 15px);--person-details-bottom-spacing:var(--person-card-details-bottom-spacing, 0);--person-background-color:var(--person-card-background-color, var(--neutral-layer-floating));--person-line1-font-size:var(--person-card-line1-font-size, 20px);--person-line1-font-weight:var(--person-card-line1-font-weight, 600);--person-line1-text-line-height:var(--person-card-line1-line-height, 28px);--person-line1-text-color:var(--person-card-line1-text-color, var(--neutral-foreground-rest));--person-line2-font-size:var(--person-card-line2-font-size, 14px);--person-line2-font-weight:var(--person-card-line2-font-weight, 400);--person-line2-text-line-height:var(--person-card-line2-line-height, 20px);--person-line2-text-color:var(--person-card-line2-text-color, var(--neutral-foreground-hint));--person-line3-font-size:var(--person-card-line3-font-size, 14px);--person-line3-font-weight:var(--person-card-line3-font-weight, 400);--person-line3-text-line-height:var(--person-card-line3-line-height, 19px);--person-line3-text-color:var(--person-card-line3-text-color, var(--neutral-foreground-hint));--person-avatar-size:var(--person-card-avatar-size, 75px);--person-presence-wrapper-bottom:-15px}:host .person-details-container .base-icons{display:flex;align-items:center;margin-inline-start:var(--person-card-base-icons-left-spacing,calc(var(--person-card-avatar-size,75px) + var(--person-card-details-left-spacing,15px) - 8px));z-index:1}:host .person-details-container .base-icons .icon{display:flex;align-items:center;font-size:13px;white-space:nowrap}:host .person-details-container .base-icons .icon:not(:last-child){margin-inline-end:28px}:host .person-details-container .base-icons .icon svg .filled{display:none}:host .person-details-container .base-icons .icon svg .regular{display:block}:host .person-details-container .base-icons .icon svg path{fill:var(--person-card-icon-color,var(--accent-foreground-rest))}:host .person-details-container .base-icons .icon:active svg .filled,:host .person-details-container .base-icons .icon:hover svg .filled{display:block}:host .person-details-container .base-icons .icon:active svg .regular,:host .person-details-container .base-icons .icon:hover svg .regular{display:none}:host .person-details-container .base-icons .icon:active svg path,:host .person-details-container .base-icons .icon:hover svg path{fill:var(--person-card-icon-hover-color,var(--accent-foreground-hover))}:host .expanded-details-container{display:flex;flex-direction:column;position:relative}:host .expanded-details-container .expanded-details-button{display:flex;justify-content:center;height:32px}:host .expanded-details-container .expanded-details-button svg path{stroke:var(--person-card-icon-color,var(--accent-foreground-rest))}:host .expanded-details-container .expanded-details-button:hover{cursor:pointer;background-color:var(--person-card-expanded-background-color-hover,var(--neutral-fill-hover))}:host .expanded-details-container .sections .section{padding:20px 0;display:flex;flex-direction:column;position:relative}:host .expanded-details-container .sections .section:not(:last-child)::after{position:absolute;content:"";width:90%;transform:translateX(-50%);border-bottom:1px solid var(--person-card-border-color,var(--neutral-stroke-rest));left:50%;bottom:0}:host .expanded-details-container .sections .section__header{display:flex;flex-direction:row;padding:0 20px}:host .expanded-details-container .sections .section__title{flex-grow:1;color:var(--person-card-line1-text-color,var(--neutral-foreground-rest));font-size:14px;font-weight:600;line-height:19px}:host .expanded-details-container .sections .section__show-more{font-size:12px;font-weight:600;align-self:center;vertical-align:top;--base-height-multiplier:6}:host .expanded-details-container .sections .section__show-more:hover{background-color:var(--person-card-fluent-background-color-hover,var(--neutral-fill-hover))}:host .expanded-details-container .sections .section__content{margin-top:14px}:host .expanded-details-container .sections .section .additional-details{padding:0 20px}:host .expanded-details-container .divider{z-index:100;position:relative;width:375px;margin:0;border-style:none;border-bottom:1px solid var(--person-card-border-color,var(--neutral-stroke-rest))}:host .expanded-details-container .section-nav{height:35px}:host .expanded-details-container .section-nav fluent-tabs{grid-template-columns:minmax(1px,.1fr)}:host .expanded-details-container .section-nav fluent-tabs.horizontal::part(activeIndicator){width:10px}:host .expanded-details-container .section-nav fluent-tabs fluent-tab-panel{max-height:360px;min-height:360px;overflow:hidden auto;padding:0;scrollbar-width:thin}:host .expanded-details-container .section-nav fluent-tabs fluent-tab-panel .inserted{padding:20px 0;max-width:var(--mgt-flyout-set-width,375px);min-width:var(--mgt-flyout-set-width,360px);--file-list-box-shadow:none;--file-list-padding:0}:host .expanded-details-container .section-nav fluent-tabs fluent-tab-panel .inserted .title{font-size:14px;font-weight:600;color:var(--title-color-subtitle,var(--neutral-foreground-rest,#1a1a1a));margin:0 20px 20px;line-height:19px}:host .expanded-details-container .section-nav fluent-tabs fluent-tab-panel .overview-panel{max-width:var(--mgt-flyout-set-width,375px)}:host .expanded-details-container .section-nav fluent-tabs fluent-tab-panel::-webkit-scrollbar{height:4em;width:4px;border-radius:11px}:host .expanded-details-container .section-nav fluent-tabs fluent-tab-panel::-webkit-scrollbar-button{height:1px}:host .expanded-details-container .section-nav fluent-tabs fluent-tab-panel::-webkit-scrollbar-track{border-radius:10px}:host .expanded-details-container .section-nav fluent-tabs fluent-tab-panel::-webkit-scrollbar-thumb{background:grey;border-radius:10px;height:4px}:host .expanded-details-container .section-nav fluent-tabs fluent-tab{padding-bottom:1px!important;border:1px solid transparent!important}:host .expanded-details-container .section-nav fluent-tabs fluent-tab:focus-visible{border:1px solid #2b2b2b!important}:host .expanded-details-container .section-nav fluent-tabs fluent-tab.section-nav__icon{cursor:pointer;box-sizing:border-box;width:53px;height:36px;display:flex;align-items:center;justify-content:center}:host .expanded-details-container .section-nav fluent-tabs fluent-tab.section-nav__icon svg{fill:var(--person-card-fluent-background-color,transparent)}:host .expanded-details-container .section-nav fluent-tabs fluent-tab.section-nav__icon:hover{background:var(--person-card-fluent-background-color-hover,var(--neutral-fill-hover));border-radius:4px;z-index:0;position:relative}:host .expanded-details-container .section-host{min-height:360px;overflow:hidden auto}:host .expanded-details-container .section-host::-webkit-scrollbar{height:4em;width:4px;border-radius:11px}:host .expanded-details-container .section-host::-webkit-scrollbar-button{background:#fff}:host .expanded-details-container .section-host::-webkit-scrollbar-track{background:#fff;border-radius:10px}:host .expanded-details-container .section-host::-webkit-scrollbar-track-piece{background:#fff}:host .expanded-details-container .section-host::-webkit-scrollbar-thumb{background:grey;border-radius:10px;height:4px}:host .expanded-details-container .section-host.small{overflow-y:hidden}:host .loading{margin:40px 20px;display:flex;justify-content:center;height:360px}:host .message-section{border-bottom:1px solid var(--person-card-chat-input-border-color,var(--neutral-foreground-hint));display:flex}:host .message-section fluent-text-field{margin:10px 10px 10px 16px;--neutral-fill-input-rest:$person-card-background-color;--neutral-fill-input-hover:$person-card-chat-input-hover-color;--neutral-fill-input-focus:$person-card-chat-input-focus-color;width:300px;margin-inline-start:16px;border-radius:4px;border:1px solid var(--person-card-chat-input-border-color,var(--neutral-foreground-hint))}:host .message-section .send-message-icon{max-width:15px;margin-top:10px}:host .message-section svg{height:17px;width:16px;fill:var(--person-card-chat-input-border-color,var(--neutral-foreground-hint))}:host .message-section svg:hover{filter:brightness(.5)}:host .message-section svg:disabled{filter:none}:host .person-root.large,:host .person-root.threelines{--person-avatar-size-3-lines:75px}:host fluent-button.close-button.lightweight::part(control),:host fluent-button.expanded-details-button::part(control),:host fluent-button.section__show-more.lightweight::part(control){background:var(--person-card-fluent-background-color,transparent)}:host fluent-button.icon::part(control),:host fluent-button.nav__back::part(control){border:none;padding:0;background:0 0}:host fluent-button.icon::part(control) :hover,:host fluent-button.nav__back::part(control) :hover{background:0 0}[dir=rtl] .base-icons{right:91px}[dir=rtl] .nav__back{width:20px!important;transform:scaleX(-1);filter:fliph;filter:FlipH}[dir=rtl] .close-card-container .close-button{right:auto;left:10px}[dir=rtl] .message-section svg{transform:scale(-1,1)}@media (forced-colors:active) and (prefers-color-scheme:dark){.root{border:1px solid #fff;border-radius:inherit}.root svg,.root svg>path{fill:#fff!important;fill-rule:nonzero!important;clip-rule:nonzero!important}.expanded-details-container>svg,.expanded-details-container>svg>path,svg,svg>path{fill:#fff!important;fill-rule:nonzero!important;clip-rule:nonzero!important}}@media (forced-colors:active) and (prefers-color-scheme:light){.root{border:1px solid #000;border-radius:inherit}.root svg,.root svg>path{fill:#000!important;fill-rule:nonzero!important;clip-rule:nonzero!important}.expanded-details-container>svg,.expanded-details-container>svg>path,svg,svg>path{fill:#000!important;fill-rule:nonzero!important;clip-rule:nonzero!important}}
`];var x=n("1/1t"),C=n("dupE"),O=n("tnKu"),w=n("Ys1X"),E=n("3ATu"),A=n("icq5");const L={showMoreSectionButton:"Show more",endOfCard:"End of the card",quickMessage:"Send a quick message",expandDetailsLabel:"Expand details",sendMessageLabel:"Send message",emailButtonLabel:"Email",callButtonLabel:"Call",chatButtonLabel:"Chat",closeCardLabel:"Close card",videoButtonLabel:"Video",goBackLabel:"Go Back"};var k=n("fU9z"),M=n("bica"),P=n("eftJ"),T=n("QBS5"),U=n("oePG"),F=n("Gy7L"),H=n("kpPY"),R=n("VRJB"),N=n("C5kU"),B=n("6fxl"),j=n("TDEi");const V="horizontal";class z extends j.a{constructor(){super(...arguments),this.orientation=V,this.activeindicator=!0,this.showActiveIndicator=!0,this.prevActiveTabIndex=0,this.activeTabIndex=0,this.ticking=!1,this.change=()=>{this.$emit("change",this.activetab)},this.isDisabledElement=e=>"true"===e.getAttribute("aria-disabled"),this.isHiddenElement=e=>e.hasAttribute("hidden"),this.isFocusableElement=e=>!this.isDisabledElement(e)&&!this.isHiddenElement(e),this.setTabs=()=>{const e="gridColumn",t="gridRow",n=this.isHorizontal()?e:t;this.activeTabIndex=this.getActiveIndex(),this.showActiveIndicator=!1,this.tabs.forEach((a,i)=>{if("tab"===a.slot){const e=this.activeTabIndex===i&&this.isFocusableElement(a);this.activeindicator&&this.isFocusableElement(a)&&(this.showActiveIndicator=!0);const t=this.tabIds[i],n=this.tabpanelIds[i];a.setAttribute("id",t),a.setAttribute("aria-selected",e?"true":"false"),a.setAttribute("aria-controls",n),a.addEventListener("click",this.handleTabClick),a.addEventListener("keydown",this.handleTabKeyDown),a.setAttribute("tabindex",e?"0":"-1"),e&&(this.activetab=a,this.activeid=t)}a.style[e]="",a.style[t]="",a.style[n]=`${i+1}`,this.isHorizontal()?a.classList.remove("vertical"):a.classList.add("vertical")})},this.setTabPanels=()=>{this.tabpanels.forEach((e,t)=>{const n=this.tabIds[t],a=this.tabpanelIds[t];e.setAttribute("id",a),e.setAttribute("aria-labelledby",n),this.activeTabIndex!==t?e.setAttribute("hidden",""):e.removeAttribute("hidden")})},this.handleTabClick=e=>{const t=e.currentTarget;1===t.nodeType&&this.isFocusableElement(t)&&(this.prevActiveTabIndex=this.activeTabIndex,this.activeTabIndex=this.tabs.indexOf(t),this.setComponent())},this.handleTabKeyDown=e=>{if(this.isHorizontal())switch(e.key){case F.c:e.preventDefault(),this.adjustBackward(e);break;case F.d:e.preventDefault(),this.adjustForward(e)}else switch(e.key){case F.e:e.preventDefault(),this.adjustBackward(e);break;case F.b:e.preventDefault(),this.adjustForward(e)}switch(e.key){case F.j:e.preventDefault(),this.adjust(-this.activeTabIndex);break;case F.f:e.preventDefault(),this.adjust(this.tabs.length-this.activeTabIndex-1)}},this.adjustForward=e=>{const t=this.tabs;let n=0;for(n=this.activetab?t.indexOf(this.activetab)+1:1,n===t.length&&(n=0);n<t.length&&t.length>1;){if(this.isFocusableElement(t[n])){this.moveToTabByIndex(t,n);break}if(this.activetab&&n===t.indexOf(this.activetab))break;n+1>=t.length?n=0:n+=1}},this.adjustBackward=e=>{const t=this.tabs;let n=0;for(n=this.activetab?t.indexOf(this.activetab)-1:0,n=n<0?t.length-1:n;n>=0&&t.length>1;){if(this.isFocusableElement(t[n])){this.moveToTabByIndex(t,n);break}n-1<0?n=t.length-1:n-=1}},this.moveToTabByIndex=(e,t)=>{const n=e[t];this.activetab=n,this.prevActiveTabIndex=this.activeTabIndex,this.activeTabIndex=t,n.focus(),this.setComponent()}}orientationChanged(){this.$fastController.isConnected&&(this.setTabs(),this.setTabPanels(),this.handleActiveIndicatorPosition())}activeidChanged(e,t){this.$fastController.isConnected&&this.tabs.length<=this.tabpanels.length&&(this.prevActiveTabIndex=this.tabs.findIndex(t=>t.id===e),this.setTabs(),this.setTabPanels(),this.handleActiveIndicatorPosition())}tabsChanged(){this.$fastController.isConnected&&this.tabs.length<=this.tabpanels.length&&(this.tabIds=this.getTabIds(),this.tabpanelIds=this.getTabPanelIds(),this.setTabs(),this.setTabPanels(),this.handleActiveIndicatorPosition())}tabpanelsChanged(){this.$fastController.isConnected&&this.tabpanels.length<=this.tabs.length&&(this.tabIds=this.getTabIds(),this.tabpanelIds=this.getTabPanelIds(),this.setTabs(),this.setTabPanels(),this.handleActiveIndicatorPosition())}getActiveIndex(){return void 0!==this.activeid?-1===this.tabIds.indexOf(this.activeid)?0:this.tabIds.indexOf(this.activeid):0}getTabIds(){return this.tabs.map(e=>{var t;return null!==(t=e.getAttribute("id"))&&void 0!==t?t:`tab-${Object(H.a)()}`})}getTabPanelIds(){return this.tabpanels.map(e=>{var t;return null!==(t=e.getAttribute("id"))&&void 0!==t?t:`panel-${Object(H.a)()}`})}setComponent(){this.activeTabIndex!==this.prevActiveTabIndex&&(this.activeid=this.tabIds[this.activeTabIndex],this.focusTab(),this.change())}isHorizontal(){return this.orientation===V}handleActiveIndicatorPosition(){this.showActiveIndicator&&this.activeindicator&&this.activeTabIndex!==this.prevActiveTabIndex&&(this.ticking?this.ticking=!1:(this.ticking=!0,this.animateActiveIndicator()))}animateActiveIndicator(){this.ticking=!0;const e=this.isHorizontal()?"gridColumn":"gridRow",t=this.isHorizontal()?"translateX":"translateY",n=this.isHorizontal()?"offsetLeft":"offsetTop",a=this.activeIndicatorRef[n];this.activeIndicatorRef.style[e]=`${this.activeTabIndex+1}`;const i=this.activeIndicatorRef[n];this.activeIndicatorRef.style[e]=`${this.prevActiveTabIndex+1}`;const r=i-a;this.activeIndicatorRef.style.transform=`${t}(${r}px)`,this.activeIndicatorRef.classList.add("activeIndicatorTransition"),this.activeIndicatorRef.addEventListener("transitionend",()=>{this.ticking=!1,this.activeIndicatorRef.style[e]=`${this.activeTabIndex+1}`,this.activeIndicatorRef.style.transform=`${t}(0px)`,this.activeIndicatorRef.classList.remove("activeIndicatorTransition")})}adjust(e){const t=this.tabs.filter(e=>this.isFocusableElement(e)),n=t.indexOf(this.activetab),a=Object(R.b)(0,t.length-1,n+e),i=this.tabs.indexOf(t[a]);i>-1&&this.moveToTabByIndex(this.tabs,i)}focusTab(){this.tabs[this.activeTabIndex].focus()}connectedCallback(){super.connectedCallback(),this.tabIds=this.getTabIds(),this.tabpanelIds=this.getTabPanelIds(),this.activeTabIndex=this.getActiveIndex()}}Object(P.a)([T.c],z.prototype,"orientation",void 0),Object(P.a)([T.c],z.prototype,"activeid",void 0),Object(P.a)([U.d],z.prototype,"tabs",void 0),Object(P.a)([U.d],z.prototype,"tabpanels",void 0),Object(P.a)([Object(T.c)({mode:"boolean"})],z.prototype,"activeindicator",void 0),Object(P.a)([U.d],z.prototype,"activeIndicatorRef",void 0),Object(P.a)([U.d],z.prototype,"showActiveIndicator",void 0),Object(B.a)(z,N.a);var G=n("6BDD"),K=n("UauI"),W=n("6vBc"),q=n("+53S"),Q=n("4X57"),Y=n("j9Xn"),J=n("xY0q"),X=n("oqLQ"),Z=n("8hiW"),$=n("QkLF"),ee=n("NLdk");const te=z.compose({baseName:"tabs",template:(e,t)=>G.a`
    <template class="${e=>e.orientation}">
        ${Object(N.d)(e,t)}
        <div class="tablist" part="tablist" role="tablist">
            <slot class="tab" name="tab" part="tab" ${Object(K.a)("tabs")}></slot>

            ${Object(W.a)(e=>e.showActiveIndicator,G.a`
                    <div
                        ${Object(q.a)("activeIndicatorRef")}
                        class="activeIndicator"
                        part="activeIndicator"
                    ></div>
                `)}
        </div>
        ${Object(N.b)(e,t)}
        <div class="tabpanel" part="tabpanel">
            <slot name="tabpanel" ${Object(K.a)("tabpanels")}></slot>
        </div>
    </template>
`,styles:(e,t)=>Q.a`
      ${Object(J.a)("grid")} :host {
        box-sizing: border-box;
        ${ee.a}
        color: ${Z.fb};
        grid-template-columns: auto 1fr auto;
        grid-template-rows: auto 1fr;
      }

      .tablist {
        display: grid;
        grid-template-rows: calc(${$.a} * 1px); auto;
        grid-template-columns: auto;
        position: relative;
        width: max-content;
        align-self: end;
      }

      .start,
      .end {
        align-self: center;
      }

      .activeIndicator {
        grid-row: 2;
        grid-column: 1;
        width: 20px;
        height: 3px;
        border-radius: calc(${Z.q} * 1px);
        justify-self: center;
        background: ${Z.e};
      }

      .activeIndicatorTransition {
        transition: transform 0.2s ease-in-out;
      }

      .tabpanel {
        grid-row: 2;
        grid-column-start: 1;
        grid-column-end: 4;
        position: relative;
      }

      :host(.vertical) {
        grid-template-rows: auto 1fr auto;
        grid-template-columns: auto 1fr;
      }

      :host(.vertical) .tablist {
        grid-row-start: 2;
        grid-row-end: 2;
        display: grid;
        grid-template-rows: auto;
        grid-template-columns: auto 1fr;
        position: relative;
        width: max-content;
        justify-self: end;
        align-self: flex-start;
        width: 100%;
      }

      :host(.vertical) .tabpanel {
        grid-column: 2;
        grid-row-start: 1;
        grid-row-end: 4;
      }

      :host(.vertical) .end {
        grid-row: 3;
      }

      :host(.vertical) .activeIndicator {
        grid-column: 1;
        grid-row: 1;
        width: 3px;
        height: 20px;
        margin-inline-start: calc(${Z.y} * 1px);
        border-radius: calc(${Z.q} * 1px);
        align-self: center;
        background: ${Z.e};
      }

      :host(.vertical) .activeIndicatorTransition {
        transition: transform 0.2s linear;
      }
    `.withBehaviors(Object(X.a)(Q.a`
        .activeIndicator,
        :host(.vertical) .activeIndicator {
          background: ${Y.a.Highlight};
        }
      `))});class ne extends j.a{}Object(P.a)([Object(T.c)({mode:"boolean"})],ne.prototype,"disabled",void 0);var ae=n("wHpb"),ie=n("FVZ7");const re=ne.compose({baseName:"tab",template:(e,t)=>G.a`
    <template slot="tab" role="tab" aria-disabled="${e=>e.disabled}">
        <slot></slot>
    </template>
`,styles:(e,t)=>Q.a`
      ${Object(J.a)("inline-flex")} :host {
        box-sizing: border-box;
        ${ee.a}
        height: calc((${$.a} + (${Z.s} * 2)) * 1px);
        padding: 0 calc((6 + (${Z.s} * 2 * ${Z.r})) * 1px);
        color: ${Z.fb};
        border-radius: calc(${Z.q} * 1px);
        border: calc(${Z.vb} * 1px) solid transparent;
        align-items: center;
        justify-content: center;
        grid-row: 1 / 3;
        cursor: pointer;
      }

      :host([aria-selected='true']) {
        z-index: 2;
      }

      :host(:hover),
      :host(:active) {
        color: ${Z.fb};
      }

      :host(:${ae.a}) {
        ${ie.a}
      }

      :host(.vertical) {
        justify-content: start;
        grid-column: 1 / 3;
      }

      :host(.vertical[aria-selected='true']) {
        z-index: 2;
      }

      :host(.vertical:hover),
      :host(.vertical:active) {
        color: ${Z.fb};
      }

      :host(.vertical:hover[aria-selected='true']) {
      }
    `.withBehaviors(Object(X.a)(Q.a`
          :host {
            forced-color-adjust: none;
            border-color: transparent;
            color: ${Y.a.ButtonText};
            fill: currentcolor;
          }
          :host(:hover),
          :host(.vertical:hover),
          :host([aria-selected='true']:hover) {
            background: transparent;
            color: ${Y.a.Highlight};
            fill: currentcolor;
          }
          :host([aria-selected='true']) {
            background: transparent;
            color: ${Y.a.Highlight};
            fill: currentcolor;
          }
          :host(:${ae.a}) {
            background: transparent;
            outline-color: ${Y.a.ButtonText};
          }
        `))});class oe extends j.a{}const se=oe.compose({baseName:"tab-panel",template:(e,t)=>G.a`
    <template slot="tabpanel" role="tabpanel">
        <slot></slot>
    </template>
`,styles:(e,t)=>Q.a`
  ${Object(J.a)("block")} :host {
    box-sizing: border-box;
    ${ee.a}
    padding: 0 calc((6 + (${Z.s} * 2 * ${Z.r})) * 1px);
  }
`});var ce=n("m1Vi"),de=n("P2Ap"),le=n("VDxL"),ue=n("0mOt"),fe=n("D4ZN"),pe=function(e,t,n,a){var i,r=arguments.length,o=r<3?t:null===a?a=Object.getOwnPropertyDescriptor(t,n):a;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)o=Reflect.decorate(e,t,n,a);else for(var s=e.length-1;s>=0;s--)(i=e[s])&&(o=(r<3?i(o):r>3?i(t,n,o):i(t,n))||o);return r>3&&o&&Object.defineProperty(t,n,o),o},me=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)},_e=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const he=()=>{Object(le.a)(M.a,te,re,se,ce.a,de.a),Object(ue.b)("person-card",be),Object(fe.b)(),Object(x.b)(),Object(w.b)(),Object(O.b)(),Object(C.b)(),Object(E.b)(),customElements.get(Object(ue.a)("person"))||Object(D.c)()};class be extends o.a{static get styles(){return I}get strings(){return L}static get requiredScopes(){return Object(A.a)()}get personDetails(){return this._personDetails}set personDetails(e){this._personDetails!==e&&(this._personDetails=e,this.personImage=this.getImage(),this.requestStateUpdate())}get userId(){return this._userId}set userId(e){e!==this._userId&&(this._userId=e,this.personDetails=null,this._cardState=null,this.requestStateUpdate())}get internalPersonDetails(){var e;return(null===(e=this._cardState)||void 0===e?void 0:e.person)||this.personDetails}constructor(){super(),this.isSending=!1,this.goBack=()=>{var e;if(!(null===(e=this._history)||void 0===e?void 0:e.length))return;const t=this._history.pop();this._currentSection=null;const n=this.renderRoot.querySelector("fluent-tab");n&&n.click(),this._cardState=t.state,this._personDetails=t.state.person,this.personImage=t.personImage,this.loadSections()},this.handleEndOfCard=e=>{if(e&&"Tab"===e.code){const e=this.renderRoot.querySelector("#end-of-container");if(e){e.blur();const t=this.renderRoot.querySelector("mgt-person");t&&t.focus()}}},this.closeCard=()=>{this.updateCurrentSection(null),this.isExpanded=!1},this.sendQuickMessage=()=>_e(this,void 0,void 0,function*(){const e=this._chatInput.trim();if(!(null==e?void 0:e.length))return;const t=this.personDetails,n=this._me.userPrincipalName;this.isSending=!0;const a=yield((e,t,n)=>y(void 0,void 0,void 0,function*(){const a={chatType:"oneonOne",members:[{"@odata.type":"#microsoft.graph.aadUserConversationMember",roles:["owner"],"user@odata.bind":`https://graph.microsoft.com/v1.0/users('${n}')`},{"@odata.type":"#microsoft.graph.aadUserConversationMember",roles:["owner"],"user@odata.bind":`https://graph.microsoft.com/v1.0/users('${t}')`}]};return yield e.api("/chats").header("Cache-Control","no-store").middlewareOptions(Object(g.a)("Chat.Create","Chat.ReadWrite")).post(a)}))(this._graph,t.userPrincipalName,n),i={body:{content:e}};yield((e,t,n)=>y(void 0,void 0,void 0,function*(){return yield e.api(`/chats/${t}/messages`).header("Cache-Control","no-store").middlewareOptions(Object(g.a)("Chat.ReadWrite","ChatMessage.Send")).post(n)}))(this._graph,a.id,i),this.isSending=!1,this.clearInputData()}),this.emailUser=()=>{const e=this.internalPersonDetails;if(e){const t=Object(f.e)(e);t&&window.open("mailto:"+t,"_blank","noreferrer")}},this.callUser=()=>{var e,t;const n=this.personDetails,a=this.personDetails;if(null===(e=null==n?void 0:n.businessPhones)||void 0===e?void 0:e.length){const e=n.businessPhones[0];e&&window.open("tel:"+e,"_blank","noreferrer")}else if(null===(t=null==a?void 0:a.phones)||void 0===t?void 0:t.length){const e=this.getPersonBusinessPhones(a)[0];e&&window.open("tel:"+e,"_blank","noreferrer")}},this.chatUser=(e=null)=>{const t=this.personDetails;if(null==t?void 0:t.userPrincipalName){let n=`https://teams.microsoft.com/l/chat/0/0?users=${t.userPrincipalName}`;(null==e?void 0:e.length)&&(n+=`&message=${e}`);const a=()=>window.open(n,"_blank","noreferrer");s.a.isAvailable?s.a.executeDeepLink(n,e=>{e||a()}):a()}},this.videoCallUser=()=>{const e=this.personDetails;if(null==e?void 0:e.userPrincipalName){const t=`https://teams.microsoft.com/l/call/0/0?users=${e.userPrincipalName}&withVideo=true`,n=()=>window.open(t,"_blank");s.a.isAvailable?s.a.executeDeepLink(t,e=>{e||n()}):n()}},this.showExpandedDetails=()=>{const e=this.renderRoot.querySelector(".root");(null==e?void 0:e.animate)&&e.animate([{height:"auto",transformOrigin:"top left"},{height:"auto",transformOrigin:"top left"}],{duration:1e3,easing:"ease-in-out",fill:"both"}),this.isExpanded=!0,this.fireCustomEvent("expanded",null,!0)},this.sendQuickMessageOnEnter=e=>{"Enter"===e.code&&this.sendQuickMessage()},this.handleGoBack=e=>{"Enter"===e.code&&this.goBack()},this._chatInput="",this._currentSection=null,this._history=[],this.sections=[],this._graph=null}attributeChangedCallback(e,t,n){super.attributeChangedCallback(e,t,n),t!==n&&"person-query"===e&&(this.personDetails=null,this._cardState=null,this.requestStateUpdate())}navigate(e){this._history.push({personDetails:this.personDetails,personImage:this.getImage(),state:this._cardState}),this._personDetails=e,this._cardState=null,this.personImage=null,this._currentSection=null,this.sections=[],this._chatInput="",this.requestStateUpdate()}clearHistory(){var e;if(this._currentSection=null,!(null===(e=this._history)||void 0===e?void 0:e.length))return;const t=this._history[0];this._history=[],this._cardState=t.state,this._personDetails=t.personDetails,this.personImage=t.personImage,this.loadSections()}render(){var e;if(!this.internalPersonDetails)return this.renderNoData();const t=this.internalPersonDetails,n=this.getImage();if(this.hasTemplate("default"))return this.renderTemplate("default",{person:this.internalPersonDetails,personImage:n});let i;i=this.strings.closeCardLabel;const o=this.isExpanded?a.c`
           <div class="close-card-container">
             <fluent-button 
              appearance="lightweight" 
              class="close-button" 
              aria-label=${i}
              @click=${this.closeCard} >
               ${Object(h.b)(h.a.Close)}
             </fluent-button>
           </div>
         `:null;i=this.strings.goBackLabel;const s=(null===(e=this._history)||void 0===e?void 0:e.length)?a.c`
            <div class="nav">
              <fluent-button 
                appearance="lightweight"
                class="nav__back" 
                aria-label=${i} 
                @keydown=${this.handleGoBack}
                @click=${this.goBack}>${Object(h.b)(h.a.Back)}
               </fluent-button>
            </div>
          `:null;let c=this.renderTemplate("person-details",{person:this.internalPersonDetails,personImage:n});if(!c){const e=this.renderPerson(),n=this.renderContactIcons(t);c=a.c`
         ${e} ${n}
       `}const d=this.isExpanded?this.renderExpandedDetails():this.renderExpandedDetailsButton();this._windowHeight=window.innerHeight&&document.documentElement.clientHeight?Math.min(window.innerHeight,document.documentElement.clientHeight):window.innerHeight||document.documentElement.clientHeight,this._windowHeight<250&&(this._smallView=!0);const l=this.lockTabNavigation?a.c`<div @keydown=${this.handleEndOfCard} aria-label=${this.strings.endOfCard} tabindex="0" id="end-of-container"></div>`:a.c``;return a.c`
      <div class="root" dir=${this.direction}>
        <div class=${Object(r.a)({small:this._smallView})}>
          ${s}
          ${o}
          <div class="person-details-container">${c}</div>
          <div class="expanded-details-container">${d}</div>
          ${l}
        </div>
      </div>
     `}renderNoData(){return this.renderTemplate("no-data",null)||a.c``}renderPerson(){return c.a`
      <mgt-person
        tabindex="0"
        class="person-image"
        .personDetails=${this.internalPersonDetails}
        .personImage=${this.getImage()}
        .personPresence=${this.personPresence}
        .showPresence=${this.showPresence}
        .view=${p.a.threelines}
      ></mgt-person>
    `}renderPersonSubtitle(e){if(e=e||this.internalPersonDetails,Object(k.c)(e)&&e.department)return a.c`
       <div class="department">${e.department}</div>
     `}renderContactIcons(e){const t=e=e||this.internalPersonDetails;let n,i,r;Object(f.e)(e)&&(n=`${this.strings.emailButtonLabel} ${e.displayName}`,i=a.c`
        <fluent-button class="icon"
          aria-label=${n}
          @click=${this.emailUser}>
          ${Object(h.b)(h.a.SmallEmail)}
        </fluent-button>
      `),(null==t?void 0:t.userPrincipalName)&&(n=`${this.strings.chatButtonLabel} ${e.displayName}`,r=a.c`
        <fluent-button class="icon"
          aria-label=${n}
          @click=${this.chatUser}>
          ${Object(h.b)(h.a.SmallChat)}
        </fluent-button>
       `),n=`${this.strings.videoButtonLabel} ${e.displayName}`;const o=a.c`
      <fluent-button class="icon"
        aria-label=${n}
        @click=${this.videoCallUser}>
        ${Object(h.b)(h.a.Video)}
      </fluent-button>
    `;let s;return t.userPrincipalName&&(n=`${this.strings.callButtonLabel} ${e.displayName}`,s=a.c`
         <fluent-button class="icon"
          aria-label=${n}
          @click=${this.callUser}>
          ${Object(h.b)(h.a.Call)}
        </fluent-button>
       `),a.c`
       <div class="base-icons">
         ${i} ${r} ${o} ${s}
       </div>
     `}renderExpandedDetailsButton(){return a.c`
      <fluent-button
        aria-label=${this.strings.expandDetailsLabel}
        class="expanded-details-button"
        @click=${this.showExpandedDetails}
      >
        ${Object(h.b)(h.a.ExpandDown)}
      </fluent-button>
     `}renderExpandedDetails(){if(!this._cardState&&this._isStateLoading)return c.a`
         <div class="loading">
           <mgt-spinner></mgt-spinner>
         </div>
       `;d.a.globalProvider.state===l.c.SignedOut&&this.loadSections();const e=this.renderSectionNavigation();return a.c`
      <div class="section-nav">
        ${e}
      </div>
      <hr class="divider"/>
      <div
        class="section-host ${this._smallView?"small":""} ${this._smallView?"small":""}"
        @wheel=${e=>this.handleSectionScroll(e)}
        tabindex=0
      ></div>
    `}renderSectionNavigation(){if(!this.sections||this.sections.length<2&&!this.hasTemplate("additional-details"))return;const e=this._currentSection?this.sections.indexOf(this._currentSection):-1,t=this.sections.map((t,n)=>{const i=t.tagName.toLowerCase(),o=Object(r.a)({active:n===e,"section-nav__icon":!0});return a.c`
        <fluent-tab
          id="${i}-Tab"
          class=${o}
          slot="tab"
          @keyup="${()=>this.updateCurrentSection(t)}"
          @click=${()=>this.updateCurrentSection(t)}
        >
          ${t.renderIcon()}
        </fluent-tab>
      `}),n=this.sections.map(e=>a.c`
        <fluent-tab-panel slot="tabpanel">
          <div class="inserted">
            <div class="title">${e.cardTitle}</div>
            ${this._currentSection?e.asFullView():null}
          </div>
        </fluent-tab-panel>
      `),i=Object(r.a)({active:-1===e,"section-nav__icon":!0,overviewTab:!0});return a.c`
      <fluent-tabs
        orientation="horizontal"
        activeindicator
        @wheel=${e=>this.handleSectionScroll(e)}
      >
        <fluent-tab
          class="${i}"
          slot="tab"
          @keyup="${()=>this.updateCurrentSection(null)}"
          @click=${()=>this.updateCurrentSection(null)}
        >
          <div>${Object(h.b)(h.a.Overview)}</div>
        </fluent-tab>
        ${t}
        <fluent-tab-panel slot="tabpanel" >
          <div class="overview-panel">${this._currentSection?null:this.renderOverviewSection()}</div>
        </fluent-tab-panel>
        ${n}
      </fluent-tabs>
    `}renderOverviewSection(){const e=this.sections.map(e=>a.c`
        <div class="section">
          <div class="section__header">
            <div class="section__title" tabindex=0>${e.displayName}</div>
              <fluent-button
                appearance="lightweight"
                class="section__show-more"
                @click=${()=>this.updateCurrentSection(e)}
              >
                ${this.strings.showMoreSectionButton}
              </fluent-button>
          </div>
          <div class="section__content">${e.asCompactView()}</div>
        </div>
      `),t=this.renderTemplate("additional-details",{person:this.internalPersonDetails,personImage:this.getImage(),state:this._cardState});return t&&e.splice(1,0,a.c`
           <div class="section">
             <div class="additional-details">${t}</div>
           </div>
         `),a.c`
       <div class="sections">
          ${this.renderMessagingSection()}
          ${e}
       </div>
     `}renderCurrentSection(){var e;if((null===(e=this.sections)||void 0===e?void 0:e.length)||this.hasTemplate("additional-details"))return 1!==this.sections.length||this.hasTemplate("additional-details")?this._currentSection?a.c`
       ${this._currentSection.asFullView()}
     `:this.renderOverviewSection():a.c`
         ${this.sections[0].asFullView()}
       `}renderMessagingSection(){const e=this.personDetails,t=this._me.userPrincipalName,n=this._chatInput;return(null==e?void 0:e.userPrincipalName)===t?void 0:v.a.isSendMessageVisible?a.c`
      <div class="message-section">
        <fluent-text-field
          autocomplete="off"
          appearance="outline"
          placeholder="${this.strings.quickMessage}"
          .value=${n}
          @input=${e=>{this._chatInput=e.target.value,this.requestUpdate()}}
          @keydown="${e=>this.sendQuickMessageOnEnter(e)}">
        </fluent-text-field>
        <fluent-button class="send-message-icon"
          aria-label=${this.strings.sendMessageLabel}
          @click=${this.sendQuickMessage}
          ?disabled=${this.isSending}>
          ${this.isSending?Object(h.b)(h.a.Confirmation):Object(h.b)(h.a.Send)}
        </fluent-button>
      </div>
      `:a.d}loadState(){var e,t;return _e(this,void 0,void 0,function*(){if(this._cardState)return;if(!this.personDetails&&this.inheritDetails){let e=this.parentElement;for(;e&&e.tagName!==`${u.a.prefix}-PERSON`.toUpperCase();)e=e.parentElement;const t=e.personDetails||e.personDetailsInternal;e&&t&&(this.personDetails=t,this.personImage=e.personImage)}const n=d.a.globalProvider;if(!n||n.state!==l.c.SignedIn)return;const a=n.graph.forComponent(this);if(this._graph=a,this._isStateLoading=!0,this._me||(this._me=yield d.a.me()),this.personDetails){const e=this.personDetails;let t;if(Object(k.c)(e)&&(t=e.userPrincipalName||e.id),t&&!Object(f.e)(e)){const e=yield Object(_.a)(a,t);this.personDetails=e,this.personImage=this.getImage()}}else if(this.userId||"me"===this.personQuery){const e=yield Object(_.a)(a,this.userId);this.personDetails=e,this.personImage=this.getImage()}else if(this.personQuery){const e=yield Object(f.d)(a,this.personQuery,1);(null==e?void 0:e.length)&&(this.personDetails=e[0],yield Object(m.d)(a,this.personDetails,v.a.useContactApis).then(e=>{e&&(this.personDetails.personImage=e,this.personImage=e)}))}const i={activity:"Offline",availability:"Offline",id:null};if(!this.personPresence&&this.showPresence)try{(null===(e=this.personDetails)||void 0===e?void 0:e.id)?this.personPresence=yield Object(b.a)(a,this.personDetails.id):this.personPresence=i}catch(e){this.personPresence=i}(null===(t=this.personDetails)||void 0===t?void 0:t.id)&&(this._cardState=yield((e,t,n)=>y(void 0,void 0,void 0,function*(){var a;const i=t.id,r=Object(f.e)(t),o="classification"in t||"personType"in t&&("PersonalContact"===t.personType.subclass||"Group"===t.personType.class),s=e.createBatch();let c;o||v.a.sections.organization&&(((e,t)=>{const n=`manager($levels=max;$select=${S})`;e.get("person",`users/${t}?$expand=${n}&$select=${S}&$count=true`,["user.read.all"],{ConsistencyLevel:"eventual"}),e.get("directReports",`users/${t}/directReports?$select=${S}`)})(s,i),v.a.sections.organization.showWorksWith&&((e,t)=>{e.get("people",`users/${t}/people?$filter=personType/class eq 'Person'`,["People.Read.All"])})(s,i)),v.a.sections.mailMessages&&r&&((e,t)=>{e.get("messages",`me/messages?$search="from:${t}"`,["Mail.ReadBasic"])})(s,r),v.a.sections.files&&((e,t)=>{let n;n=t?`me/insights/shared?$filter=lastshared/sharedby/address eq '${t}'`:"me/insights/used",e.get("files",n,["Sites.Read.All"])})(s,n?null:r);const d={};try{c=yield s.executeAll()}catch(e){}if(c)for(const[e,t]of c)d[e]=(null===(a=t.content)||void 0===a?void 0:a.value)||t.content;if(!o&&v.a.sections.profile)try{const t=yield((e,t)=>y(void 0,void 0,void 0,function*(){return yield e.api(`/users/${t}/profile`).version("beta").get()}))(e,i);t&&(d.profile=t)}catch(e){}return d.directReports&&d.directReports.length>0&&(d.directReports=d.directReports.filter(e=>e.accountEnabled)),d}))(a,this.personDetails,this._me===this.personDetails.id)),this.loadSections(),this._isStateLoading=!1})}get hasPhone(){var e,t;const n=this.personDetails,a=this.personDetails;return Boolean(null===(e=null==n?void 0:n.businessPhones)||void 0===e?void 0:e.length)||Boolean(null===(t=null==a?void 0:a.phones)||void 0===t?void 0:t.length)}loadSections(){if(this.sections=[],!this.internalPersonDetails)return;const e=new x.a(this.internalPersonDetails);if(e.hasData&&this.sections.push(e),!this._cardState)return;const{person:t,directReports:n,messages:a,files:i,profile:r}=this._cardState;if(v.a.sections.organization&&((null==t?void 0:t.manager)||(null==n?void 0:n.length))&&this.sections.push(new w.a(this._cardState,this._me)),v.a.sections.mailMessages&&(null==a?void 0:a.length)&&this.sections.push(new O.a(a)),v.a.sections.files&&(null==i?void 0:i.length)&&this.sections.push(new C.a),v.a.sections.profile&&r){const e=new E.a(r);e.hasData&&this.sections.push(e)}}getImage(){if(this.personImage)return this.personImage;const e=this.personDetails;return(null==e?void 0:e.personImage)?e.personImage:null}clearInputData(){this._chatInput="",this.requestUpdate()}getPersonBusinessPhones(e){const t=e.phones,n=[];for(const e of t)"business"===e.type&&n.push(e.number);return n}updateCurrentSection(e){if(e){const t=e.tagName.toLowerCase();this.renderRoot.querySelector(`#${t}-Tab`).click()}const t=this.renderRoot.querySelectorAll("fluent-tab-panel");for(const e of t)e.scrollTop=0;this._currentSection=e,this.requestUpdate()}handleSectionScroll(e){const t=this.renderRoot.querySelectorAll("fluent-tab-panel");for(const n of t)n&&(e.deltaY<0&&0===n.scrollTop||e.deltaY>0&&n.clientHeight+n.scrollTop>=n.scrollHeight-1||e.stopPropagation())}}pe([Object(i.b)({attribute:"person-details",type:Object}),me("design:type",Object),me("design:paramtypes",[Object])],be.prototype,"personDetails",null),pe([Object(i.b)({attribute:"person-query"}),me("design:type",String)],be.prototype,"personQuery",void 0),pe([Object(i.b)({attribute:"lock-tab-navigation",type:Boolean}),me("design:type",Boolean)],be.prototype,"lockTabNavigation",void 0),pe([Object(i.b)({attribute:"user-id"}),me("design:type",String),me("design:paramtypes",[String])],be.prototype,"userId",null),pe([Object(i.b)({attribute:"person-image",type:String}),me("design:type",String)],be.prototype,"personImage",void 0),pe([Object(i.b)({attribute:"fetch-image",type:Boolean}),me("design:type",Boolean)],be.prototype,"fetchImage",void 0),pe([Object(i.b)({attribute:"is-expanded",type:Boolean}),me("design:type",Boolean)],be.prototype,"isExpanded",void 0),pe([Object(i.b)({attribute:"inherit-details",type:Boolean}),me("design:type",Boolean)],be.prototype,"inheritDetails",void 0),pe([Object(i.b)({attribute:"show-presence",type:Boolean}),me("design:type",Boolean)],be.prototype,"showPresence",void 0),pe([Object(i.b)({attribute:"person-presence",type:Object}),me("design:type",Object)],be.prototype,"personPresence",void 0),pe([Object(i.c)(),me("design:type",Object)],be.prototype,"isSending",void 0),pe([Object(i.c)(),me("design:type",Object)],be.prototype,"_cardState",void 0),pe([Object(i.c)(),me("design:type",Boolean)],be.prototype,"_isStateLoading",void 0)},uXNP:function(e,t,n){"use strict";n.d(t,"a",function(){return r});var a=n("eftJ"),i=n("QBS5");class r{}Object(a.a)([Object(i.c)({attribute:"aria-atomic"})],r.prototype,"ariaAtomic",void 0),Object(a.a)([Object(i.c)({attribute:"aria-busy"})],r.prototype,"ariaBusy",void 0),Object(a.a)([Object(i.c)({attribute:"aria-controls"})],r.prototype,"ariaControls",void 0),Object(a.a)([Object(i.c)({attribute:"aria-current"})],r.prototype,"ariaCurrent",void 0),Object(a.a)([Object(i.c)({attribute:"aria-describedby"})],r.prototype,"ariaDescribedby",void 0),Object(a.a)([Object(i.c)({attribute:"aria-details"})],r.prototype,"ariaDetails",void 0),Object(a.a)([Object(i.c)({attribute:"aria-disabled"})],r.prototype,"ariaDisabled",void 0),Object(a.a)([Object(i.c)({attribute:"aria-errormessage"})],r.prototype,"ariaErrormessage",void 0),Object(a.a)([Object(i.c)({attribute:"aria-flowto"})],r.prototype,"ariaFlowto",void 0),Object(a.a)([Object(i.c)({attribute:"aria-haspopup"})],r.prototype,"ariaHaspopup",void 0),Object(a.a)([Object(i.c)({attribute:"aria-hidden"})],r.prototype,"ariaHidden",void 0),Object(a.a)([Object(i.c)({attribute:"aria-invalid"})],r.prototype,"ariaInvalid",void 0),Object(a.a)([Object(i.c)({attribute:"aria-keyshortcuts"})],r.prototype,"ariaKeyshortcuts",void 0),Object(a.a)([Object(i.c)({attribute:"aria-label"})],r.prototype,"ariaLabel",void 0),Object(a.a)([Object(i.c)({attribute:"aria-labelledby"})],r.prototype,"ariaLabelledby",void 0),Object(a.a)([Object(i.c)({attribute:"aria-live"})],r.prototype,"ariaLive",void 0),Object(a.a)([Object(i.c)({attribute:"aria-owns"})],r.prototype,"ariaOwns",void 0),Object(a.a)([Object(i.c)({attribute:"aria-relevant"})],r.prototype,"ariaRelevant",void 0),Object(a.a)([Object(i.c)({attribute:"aria-roledescription"})],r.prototype,"ariaRoledescription",void 0)},wHpb:function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("FGLN");const i=Object(a.a)()?"focus-visible":"focus"},"x+GM":function(e,t,n){"use strict";n.d(t,"a",function(){return a});class a{constructor(){this.eventHandlers=[]}fire(e){for(const t of this.eventHandlers)t(e)}add(e){this.eventHandlers.push(e)}remove(e){for(let t=0;t<this.eventHandlers.length;t++)this.eventHandlers[t]===e&&(this.eventHandlers.splice(t,1),t--)}}},xAa8:function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("0q6d");class i{constructor(e,t,n){this.x=e,this.y=t,this.z=n}static fromObject(e){return!e||isNaN(e.x)||isNaN(e.y)||isNaN(e.z)?null:new i(e.x,e.y,e.z)}equalValue(e){return this.x===e.x&&this.y===e.y&&this.z===e.z}roundToPrecision(e){return new i(Object(a.i)(this.x,e),Object(a.i)(this.y,e),Object(a.i)(this.z,e))}toObject(){return{x:this.x,y:this.y,z:this.z}}}i.whitePoint=new i(.95047,1,1.08883)},xY0q:function(e,t,n){"use strict";n.d(t,"a",function(){return i});const a=":host([hidden]){display:none}";function i(e){return`${a}:host{display:${e}}`}},yQ0v:function(e,t,n){"use strict";n.d(t,"b",function(){return r}),n.d(t,"a",function(){return a});var a,i=n("i2HT");function r(e){return i.a.create(e,e,e)}!function(e){e[e.LightMode=.98]="LightMode",e[e.DarkMode=.15]="DarkMode"}(a||(a={}))},ylMe:function(e,t,n){"use strict";n.d(t,"b",function(){return a}),n.d(t,"a",function(){return r});const a=(e,t)=>i(e,t,new Set),i=(e,t,n)=>{const a=Object.prototype.toString.call(e),r=Object.prototype.toString.call(t);if("object"==typeof e&&"object"==typeof t&&a===r&&"[object Object]"===a&&!n.has(e)){n.add(e);for(const a in e)if(!i(e[a],t[a],n))return!1;for(const n in t)if(!Object.prototype.hasOwnProperty.call(e,n))return!1;return!0}if(Array.isArray(e)&&Array.isArray(t)&&!n.has(e)){if(n.add(e),e.length!==t.length)return!1;for(let a=0;a<e.length;a++)if(!i(e[a],t[a],n))return!1;return!0}return e===t},r=(e,t)=>{if(e===t)return!0;if(!e||!t)return!1;if(e.length!==t.length)return!1;if(0===e.length)return!0;const n=new Set(e);for(const e of t)if(!n.has(e))return!1;return!0}},z0DP:function(e,t,n){"use strict";n.d(t,"a",function(){return a}),n.d(t,"b",function(){return i}),n.d(t,"d",function(){return f}),n.d(t,"f",function(){return p}),n.d(t,"e",function(){return m}),n.d(t,"c",function(){return _}),n.d(t,"g",function(){return h});var a,i,r=n("DNu6"),o=n("RIOo"),s=n("/4Fm"),c=n("imsm"),d=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};!function(e){e[e.any=0]="any",e.person="person",e.group="group"}(a||(a={})),function(e){e.any="any",e.user="user",e.contact="contact"}(i||(i={}));const l=()=>r.a.config.people.invalidationPeriod||r.a.config.defaultInvalidationPeriod,u=()=>r.a.config.people.isEnabled&&r.a.config.isEnabled,f=(e,t,n=10,a=i.any,s="")=>d(void 0,void 0,void 0,function*(){const d=`${t}:${n}:${a}`;let f;if(u()){const e=c.a.people,t=c.a.people.stores.peopleQuery;f=r.a.getCache(e,t);const n=u()?yield f.getValue(d):null;if(n&&l()>Date.now()-n.timeCached)return n.results.map(e=>JSON.parse(e))}let p,m="personType/class eq 'Person'";a!==i.any&&(a===i.user?m+="and personType/subclass eq 'OrganizationUser'":m+="and (personType/subclass eq 'ImplicitContact' or personType/subclass eq 'PersonalContact')"),""!==s&&(m+=` and  ${s}`);try{let r=e.api("/me/people").search('"'+t+'"').top(n).filter(m).middlewareOptions(Object(o.a)("people.read"));if(a!==i.contact&&(r=r.header("X-PeopleQuery-QuerySources","Mailbox,Directory")),p=yield r.get(),u()&&p){const e={maxResults:n,results:null};e.results=p.value.map(e=>JSON.stringify(e)),yield f.putValue(d,e)}}catch(e){}return null==p?void 0:p.value}),p=(e,t=i.any,n="",a=10)=>d(void 0,void 0,void 0,function*(){let s;const d=`${n||`*:${t}`}:${a}`;if(u()){s=r.a.getCache(c.a.people,c.a.people.stores.peopleQuery);const e=yield s.getValue(d);if(e&&l()>Date.now()-e.timeCached)return e.results.map(e=>JSON.parse(e))}let f,p="personType/class eq 'Person'";t!==i.any&&(t===i.user?p+="and personType/subclass eq 'OrganizationUser'":p+="and (personType/subclass eq 'ImplicitContact' or personType/subclass eq 'PersonalContact')"),n&&(p+=` and ${n}`);try{let n=e.api("/me/people").middlewareOptions(Object(o.a)("people.read")).top(a).filter(p);t!==i.contact&&(n=n.header("X-PeopleQuery-QuerySources","Mailbox,Directory")),f=yield n.get(),u()&&f&&(yield s.putValue(d,{maxResults:10,results:f.value.map(e=>JSON.stringify(e))}))}catch(e){}return f?f.value:null}),m=e=>{var t,n;const a=e,i=e,r=e;return i.mail?Object(s.c)(i.mail):(null===(t=a.scoredEmailAddresses)||void 0===t?void 0:t.length)?Object(s.c)(a.scoredEmailAddresses[0].address):(null===(n=r.emailAddresses)||void 0===n?void 0:n.length)?Object(s.c)(r.emailAddresses[0].address):null},_=(e,t)=>d(void 0,void 0,void 0,function*(){let n;if(u()){n=r.a.getCache(c.a.people,c.a.people.stores.contacts);const e=yield n.getValue(t);if(e&&l()>Date.now()-e.timeCached)return JSON.parse(e.person)}const a=`${t.replace(/#/g,"%2523")}`,i=yield e.api("/me/contacts").filter(`emailAddresses/any(a:a/address eq '${a}')`).middlewareOptions(Object(o.a)("contacts.read")).get();return u()&&i&&(yield n.putValue(t,{person:JSON.stringify(i.value)})),i?i.value:null}),h=(e,t,n,a)=>d(void 0,void 0,void 0,function*(){var i;let s;const d=`${t}${n}`;if(u()){s=r.a.getCache(c.a.people,c.a.people.stores.peopleQuery);const e=yield s.getValue(d);if(e&&l()>Date.now()-e.timeCached)return e.results.map(e=>JSON.parse(e))}let f=e.api(n).version(t);(null==a?void 0:a.length)&&(f=f.middlewareOptions(Object(o.a)(...a)));let p=yield f.get();if(p&&Array.isArray(p.value)&&p["@odata.nextLink"]){let n=p;for(;null==n?void 0:n["@odata.nextLink"];){const a=n["@odata.nextLink"].split(t)[1];n=yield e.client.api(a).version(t).get(),(null===(i=null==n?void 0:n.value)||void 0===i?void 0:i.length)&&(n.value=p.value.concat(n.value),p=n)}}if(u()&&p){const e={results:null};Array.isArray(p.value)?e.results=p.value.map(e=>JSON.stringify(e)):e.results=[JSON.stringify(p)],yield s.putValue(d,e)}return null==p?void 0:p.value})},zCAG:function(e,t,n){"use strict";var a;n.d(t,"a",function(){return a}),function(e){e[e.none=0]="none",e[e.hover=1]="hover",e[e.click=2]="click"}(a||(a={}))},zFbe:function(e,t,n){"use strict";n.d(t,"a",function(){return s}),n.d(t,"b",function(){return a});var a,i=n("x+GM"),r=n("/i08"),o=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};class s{static get globalProvider(){return this._globalProvider}static set globalProvider(e){e!==this._globalProvider&&(this._globalProvider&&(this._globalProvider.removeStateChangedHandler(this.handleProviderStateChanged),this._globalProvider.isMultiAccountSupportedAndEnabled&&this._globalProvider.removeActiveAccountChangedHandler(this.handleActiveAccountChanged)),e&&(e.onStateChanged(this.handleProviderStateChanged),e.isMultiAccountSupportedAndEnabled&&e.onActiveAccountChanged(this.handleActiveAccountChanged)),this._globalProvider=e,this._eventDispatcher.fire(a.ProviderChanged))}static onProviderUpdated(e){this._eventDispatcher.add(e)}static removeProviderUpdatedListener(e){this._eventDispatcher.remove(e)}static onActiveAccountChanged(e){this._activeAccountChangedDispatcher.add(e)}static removeActiveAccountChangedListener(e){this._activeAccountChangedDispatcher.remove(e)}static me(){return this.client?(this._mePromise||(this._mePromise=this.getMe()),this._mePromise):(this._mePromise=null,null)}static getMe(){return o(this,void 0,void 0,function*(){try{const e=yield this.client.api("me").get();if(null==e?void 0:e.id)return e}catch(e){}return null})}static getCacheId(){var e;return o(this,void 0,void 0,function*(){if(this._cacheId)return this._cacheId;if((null===(e=s.globalProvider)||void 0===e?void 0:e.state)===r.c.SignedIn&&!this._cacheId&&this.client)try{this._cacheId=yield this.createCacheId()}catch(e){}return this._cacheId})}static unsetCacheId(){this._cacheId=null,this._mePromise=null}static createCacheId(){return o(this,void 0,void 0,function*(){if(s.globalProvider.isMultiAccountSupportedAndEnabled){const e=this.createCacheIdWithAccountDetails();if(e)return e}return yield this.createCacheIdWithUserDetails()})}static createCacheIdWithUserDetails(){return o(this,void 0,void 0,function*(){const e=yield this.me();return(null==e?void 0:e.id)?e.id+"-"+e.userPrincipalName:null})}static createCacheIdWithAccountDetails(){const e=s.globalProvider.getActiveAccount();return e.tenantId&&e.id?e.tenantId+e.id:null}static get client(){return s.globalProvider&&s.globalProvider.state===r.c.SignedIn?s.globalProvider.graph.client:null}}s._eventDispatcher=new i.a,s._activeAccountChangedDispatcher=new i.a,s.handleProviderStateChanged=()=>{s.globalProvider&&s.globalProvider.state===r.c.SignedIn||(s._mePromise=null),s._eventDispatcher.fire(a.ProviderStateChanged)},s.handleActiveAccountChanged=()=>{s.unsetCacheId(),s._activeAccountChangedDispatcher.fire(null)},function(e){e[e.ProviderChanged=0]="ProviderChanged",e[e.ProviderStateChanged=1]="ProviderStateChanged"}(a||(a={}))},zb6c:function(e,t,n){"use strict";n.d(t,"a",function(){return i});var a=n("x+GM");class i{static get strings(){return this._strings}static set strings(e){this._strings=e,this._stringsEventDispatcher.fire(null)}static getDocumentDirection(){var e,t;switch((null===(e=document.body)||void 0===e?void 0:e.getAttribute("dir"))||(null===(t=document.documentElement)||void 0===t?void 0:t.getAttribute("dir"))){case"rtl":return"rtl";case"auto":return"auto";default:return"ltr"}}static onStringsUpdated(e){this._stringsEventDispatcher.add(e)}static removeOnStringsUpdated(e){this._stringsEventDispatcher.remove(e)}static onDirectionUpdated(e){this._directionEventDispatcher.add(e),this.initDirection()}static removeOnDirectionUpdated(e){this._directionEventDispatcher.remove(e)}static initDirection(){if(this._isDirectionInit)return;this._isDirectionInit=!0,this.mutationObserver=new MutationObserver(e=>{e.forEach(e=>{"dir"===e.attributeName&&this._directionEventDispatcher.fire(null)})});const e={attributes:!0,attributeFilter:["dir"]};this.mutationObserver.observe(document.body,e),this.mutationObserver.observe(document.documentElement,e)}static updateStringsForTag(e,t){var n;if((e=e.toLowerCase()).startsWith("mgt-")&&(e=e.substring(4)),this._strings&&t){for(const e of Object.entries(t)){const n=this._strings[e[0]];"string"==typeof n&&(t[e[0]]=n)}if(null===(n=this._strings._components)||void 0===n?void 0:n[e]){const n=this._strings._components[e];for(const e of Object.keys(n))t[e]&&(t[e]=n[e])}}return t}}i._stringsEventDispatcher=new a.a,i._directionEventDispatcher=new a.a,i._isDirectionInit=!1},zlIh:function(e,t,n){"use strict";n.d(t,"a",function(){return d}),n.d(t,"b",function(){return l});var a=n("DNu6"),i=n("RIOo"),r=n("imsm"),o=function(e,t,n,a){return new(n||(n=Promise))(function(i,r){function o(e){try{c(a.next(e))}catch(e){r(e)}}function s(e){try{c(a.throw(e))}catch(e){r(e)}}function c(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(o,s)}c((a=a.apply(e,t||[])).next())})};const s=()=>a.a.config.presence.invalidationPeriod||a.a.config.defaultInvalidationPeriod,c=()=>a.a.config.presence.isEnabled&&a.a.config.isEnabled,d=(e,t)=>o(void 0,void 0,void 0,function*(){let n;if(c()){n=a.a.getCache(r.a.presence,r.a.presence.stores.presence);const e=yield n.getValue(t||"me");if(e&&s()>Date.now()-e.timeCached)return JSON.parse(e.presence)}const o=t?["presence.read.all"]:["presence.read"],d=t?`/users/${t}/presence`:"/me/presence",l=yield e.api(d).middlewareOptions(Object(i.a)(...o)).get();return c()&&(yield n.putValue(t||"me",{presence:JSON.stringify(l)})),l}),l=(e,t,n=!1)=>o(void 0,void 0,void 0,function*(){if(!t||0===t.length)return{};const o={},l=[],u=["presence.read.all"];let f;c()&&(f=a.a.getCache(r.a.presence,r.a.presence.stores.presence));for(const e of t)if(null==e?void 0:e.id){const t=e.id;let a;o[t]=null,!n&&c()&&(a=yield f.getValue(t)),!n&&c()&&a&&s()>Date.now()-a.timeCached?o[t]=JSON.parse(a.presence):l.push(t)}try{if(l.length>0){const t=yield e.api("/communications/getPresencesByUserId").middlewareOptions(Object(i.a)(...u)).post({ids:l});for(const e of t.value)o[e.id]=e,c()&&(yield f.putValue(e.id,{presence:JSON.stringify(e)}))}return o}catch(n){try{const n=yield Promise.all(t.filter(e=>(null==e?void 0:e.id)&&!o[e.id]&&"personType"in e&&"OrganizationUser"===e.personType.subclass).map(t=>d(e,t.id)));for(const e of n)o[e.id]=e;return o}catch(e){return null}}})}})});