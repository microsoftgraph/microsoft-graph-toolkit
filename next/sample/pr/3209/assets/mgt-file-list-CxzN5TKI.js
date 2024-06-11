import{m as _,w as H,aE as ae,F as se,cs as re,p as ne,D as de,O as ce,_ as w,b as U,x as P,ct as ue,ca as pe,M as he,I as C,y as fe,cu as ge,C as $,L as I,K as M,P as me,Q as q,cv as ve,W,Z as X,$ as be,a1 as Z,cw as ye,Y as v,a2 as S,a3 as k,a6 as L,a4 as E,cx as N,a0 as f,X as xe,a8 as Y,cy as Fe,cz as Se}from"./App-Ofsx7rcY.js";import{c as ke}from"./repeat-BfXRkkP9.js";import{d as we,a as Ue,s as $e,b as Ce,c as _e,i as Le,e as J,f as Te,h as Ie,j as Ee,k as De,l as ze,m as Pe,n as Be,o as Re,p as Ae,q as Oe,r as Me,t as qe,u as Ne,w as je,x as Qe,y as j}from"./graph.files-C7QBtkFu.js";import{r as ee,M as Ge}from"./mgt-file-CUcuJr7p.js";import{a as Ke}from"./index-CBUvQWBQ.js";import{b as D,a as Q}from"./index-BTg_gX7d.js";const Ve=(s,e)=>_`
    <div class="positioning-region" part="positioning-region">
        ${H(t=>t.modal,_`
                <div
                    class="overlay"
                    part="overlay"
                    role="presentation"
                    @click="${t=>t.dismiss()}"
                ></div>
            `)}
        <div
            role="dialog"
            tabindex="-1"
            class="control"
            part="control"
            aria-modal="${t=>t.modal}"
            aria-describedby="${t=>t.ariaDescribedby}"
            aria-labelledby="${t=>t.ariaLabelledby}"
            aria-label="${t=>t.ariaLabel}"
            ${ae("dialog")}
        >
            <slot></slot>
        </div>
    </div>
`;/*!
* tabbable 5.3.3
* @license MIT, https://github.com/focus-trap/tabbable/blob/master/LICENSE
*/var He=["input","select","textarea","a[href]","button","[tabindex]:not(slot)","audio[controls]","video[controls]",'[contenteditable]:not([contenteditable="false"])',"details>summary:first-of-type","details"],We=He.join(","),te=typeof Element>"u",T=te?function(){}:Element.prototype.matches||Element.prototype.msMatchesSelector||Element.prototype.webkitMatchesSelector,B=!te&&Element.prototype.getRootNode?function(s){return s.getRootNode()}:function(s){return s.ownerDocument},Xe=function(e,t){return e.tabIndex<0&&(t||/^(AUDIO|VIDEO|DETAILS)$/.test(e.tagName)||e.isContentEditable)&&isNaN(parseInt(e.getAttribute("tabindex"),10))?0:e.tabIndex},ie=function(e){return e.tagName==="INPUT"},Ze=function(e){return ie(e)&&e.type==="hidden"},Ye=function(e){var t=e.tagName==="DETAILS"&&Array.prototype.slice.apply(e.children).some(function(i){return i.tagName==="SUMMARY"});return t},Je=function(e,t){for(var i=0;i<e.length;i++)if(e[i].checked&&e[i].form===t)return e[i]},et=function(e){if(!e.name)return!0;var t=e.form||B(e),i=function(r){return t.querySelectorAll('input[type="radio"][name="'+r+'"]')},l;if(typeof window<"u"&&typeof window.CSS<"u"&&typeof window.CSS.escape=="function")l=i(window.CSS.escape(e.name));else try{l=i(e.name)}catch(a){return console.error("Looks like you have a radio button with a name attribute containing invalid CSS selector characters and need the CSS.escape polyfill: %s",a.message),!1}var o=Je(l,e.form);return!o||o===e},tt=function(e){return ie(e)&&e.type==="radio"},it=function(e){return tt(e)&&!et(e)},G=function(e){var t=e.getBoundingClientRect(),i=t.width,l=t.height;return i===0&&l===0},ot=function(e,t){var i=t.displayCheck,l=t.getShadowRoot;if(getComputedStyle(e).visibility==="hidden")return!0;var o=T.call(e,"details>summary:first-of-type"),a=o?e.parentElement:e;if(T.call(a,"details:not([open]) *"))return!0;var r=B(e).host,n=(r==null?void 0:r.ownerDocument.contains(r))||e.ownerDocument.contains(e);if(!i||i==="full"){if(typeof l=="function"){for(var c=e;e;){var d=e.parentElement,h=B(e);if(d&&!d.shadowRoot&&l(d)===!0)return G(e);e.assignedSlot?e=e.assignedSlot:!d&&h!==e.ownerDocument?e=h.host:e=d}e=c}if(n)return!e.getClientRects().length}else if(i==="non-zero-area")return G(e);return!1},lt=function(e){if(/^(INPUT|BUTTON|SELECT|TEXTAREA)$/.test(e.tagName))for(var t=e.parentElement;t;){if(t.tagName==="FIELDSET"&&t.disabled){for(var i=0;i<t.children.length;i++){var l=t.children.item(i);if(l.tagName==="LEGEND")return T.call(t,"fieldset[disabled] *")?!0:!l.contains(e)}return!0}t=t.parentElement}return!1},at=function(e,t){return!(t.disabled||Ze(t)||ot(t,e)||Ye(t)||lt(t))},st=function(e,t){return!(it(t)||Xe(t)<0||!at(e,t))},K=function(e,t){if(t=t||{},!e)throw new Error("No node provided");return T.call(e,We)===!1?!1:st(t,e)};class y extends se{constructor(){super(...arguments),this.modal=!0,this.hidden=!1,this.trapFocus=!0,this.trapFocusChanged=()=>{this.$fastController.isConnected&&this.updateTrapFocus()},this.isTrappingFocus=!1,this.handleDocumentKeydown=e=>{if(!e.defaultPrevented&&!this.hidden)switch(e.key){case ne:this.dismiss(),e.preventDefault();break;case re:this.handleTabKeyDown(e);break}},this.handleDocumentFocus=e=>{!e.defaultPrevented&&this.shouldForceFocus(e.target)&&(this.focusFirstElement(),e.preventDefault())},this.handleTabKeyDown=e=>{if(!this.trapFocus||this.hidden)return;const t=this.getTabQueueBounds();if(t.length!==0){if(t.length===1){t[0].focus(),e.preventDefault();return}e.shiftKey&&e.target===t[0]?(t[t.length-1].focus(),e.preventDefault()):!e.shiftKey&&e.target===t[t.length-1]&&(t[0].focus(),e.preventDefault())}},this.getTabQueueBounds=()=>{const e=[];return y.reduceTabbableItems(e,this)},this.focusFirstElement=()=>{const e=this.getTabQueueBounds();e.length>0?e[0].focus():this.dialog instanceof HTMLElement&&this.dialog.focus()},this.shouldForceFocus=e=>this.isTrappingFocus&&!this.contains(e),this.shouldTrapFocus=()=>this.trapFocus&&!this.hidden,this.updateTrapFocus=e=>{const t=e===void 0?this.shouldTrapFocus():e;t&&!this.isTrappingFocus?(this.isTrappingFocus=!0,document.addEventListener("focusin",this.handleDocumentFocus),de.queueUpdate(()=>{this.shouldForceFocus(document.activeElement)&&this.focusFirstElement()})):!t&&this.isTrappingFocus&&(this.isTrappingFocus=!1,document.removeEventListener("focusin",this.handleDocumentFocus))}}dismiss(){this.$emit("dismiss"),this.$emit("cancel")}show(){this.hidden=!1}hide(){this.hidden=!0,this.$emit("close")}connectedCallback(){super.connectedCallback(),document.addEventListener("keydown",this.handleDocumentKeydown),this.notifier=ce.getNotifier(this),this.notifier.subscribe(this,"hidden"),this.updateTrapFocus()}disconnectedCallback(){super.disconnectedCallback(),document.removeEventListener("keydown",this.handleDocumentKeydown),this.updateTrapFocus(!1),this.notifier.unsubscribe(this,"hidden")}handleChange(e,t){switch(t){case"hidden":this.updateTrapFocus();break}}static reduceTabbableItems(e,t){return t.getAttribute("tabindex")==="-1"?e:K(t)||y.isFocusableFastElement(t)&&y.hasTabbableShadow(t)?(e.push(t),e):t.childElementCount?e.concat(Array.from(t.children).reduce(y.reduceTabbableItems,[])):e}static isFocusableFastElement(e){var t,i;return!!(!((i=(t=e.$fastController)===null||t===void 0?void 0:t.definition.shadowOptions)===null||i===void 0)&&i.delegatesFocus)}static hasTabbableShadow(e){var t,i;return Array.from((i=(t=e.shadowRoot)===null||t===void 0?void 0:t.querySelectorAll("*"))!==null&&i!==void 0?i:[]).some(l=>K(l))}}w([U({mode:"boolean"})],y.prototype,"modal",void 0);w([U({mode:"boolean"})],y.prototype,"hidden",void 0);w([U({attribute:"trap-focus",mode:"boolean"})],y.prototype,"trapFocus",void 0);w([U({attribute:"aria-describedby"})],y.prototype,"ariaDescribedby",void 0);w([U({attribute:"aria-labelledby"})],y.prototype,"ariaLabelledby",void 0);w([U({attribute:"aria-label"})],y.prototype,"ariaLabel",void 0);const rt=(s,e)=>_`
    <template
        role="progressbar"
        aria-valuenow="${t=>t.value}"
        aria-valuemin="${t=>t.min}"
        aria-valuemax="${t=>t.max}"
        class="${t=>t.paused?"paused":""}"
    >
        ${H(t=>typeof t.value=="number",_`
                <div class="progress" part="progress" slot="determinate">
                    <div
                        class="determinate"
                        part="determinate"
                        style="width: ${t=>t.percentComplete}%"
                    ></div>
                </div>
            `,_`
                <div class="progress" part="progress" slot="indeterminate">
                    <slot class="indeterminate" name="indeterminate">
                        ${e.indeterminateIndicator1||""}
                        ${e.indeterminateIndicator2||""}
                    </slot>
                </div>
            `)}
    </template>
`,nt=(s,e)=>P`
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
    box-shadow: ${ue};
    margin-top: auto;
    margin-bottom: auto;
    border-radius: calc(${pe} * 1px);
    width: var(--dialog-width);
    height: var(--dialog-height);
    background: ${he};
    z-index: 1;
    border: calc(${C} * 1px) solid transparent;
  }
`,dt=y.compose({baseName:"dialog",template:Ve,styles:nt}),ct=(s,e)=>P`
    ${fe("flex")} :host {
      align-items: center;
      height: calc((${C} * 3) * 1px);
    }

    .progress {
      background-color: ${ge};
      border-radius: calc(${$} * 1px);
      width: 100%;
      height: calc(${C} * 1px);
      display: flex;
      align-items: center;
      position: relative;
    }

    .determinate {
      background-color: ${I};
      border-radius: calc(${$} * 1px);
      height: calc((${C} * 3) * 1px);
      transition: all 0.2s ease-in-out;
      display: flex;
    }

    .indeterminate {
      height: calc((${C} * 3) * 1px);
      border-radius: calc(${$} * 1px);
      display: flex;
      width: 100%;
      position: relative;
      overflow: hidden;
    }

    .indeterminate-indicator-1 {
      position: absolute;
      opacity: 0;
      height: 100%;
      background-color: ${I};
      border-radius: calc(${$} * 1px);
      animation-timing-function: cubic-bezier(0.4, 0, 0.6, 1);
      width: 40%;
      animation: indeterminate-1 2s infinite;
    }

    .indeterminate-indicator-2 {
      position: absolute;
      opacity: 0;
      height: 100%;
      background-color: ${I};
      border-radius: calc(${$} * 1px);
      animation-timing-function: cubic-bezier(0.4, 0, 0.6, 1);
      width: 60%;
      animation: indeterminate-2 2s infinite;
    }

    :host(.paused) .indeterminate-indicator-1,
    :host(.paused) .indeterminate-indicator-2 {
      animation: none;
      background-color: ${M};
      width: 100%;
      opacity: 1;
    }

    :host(.paused) .determinate {
      background-color: ${M};
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
  `.withBehaviors(me(P`
        .indeterminate-indicator-1,
        .indeterminate-indicator-2,
        .determinate,
        .progress {
          background-color: ${q.ButtonText};
        }
        :host(.paused) .indeterminate-indicator-1,
        :host(.paused) .indeterminate-indicator-2,
        :host(.paused) .determinate {
          background-color: ${q.GrayText};
        }
      `));class ut extends ve{}const pt=ut.compose({baseName:"progress",template:rt,styles:ct,indeterminateIndicator1:`
    <span class="indeterminate-indicator-1" part="indeterminate-indicator-1"></span>
  `,indeterminateIndicator2:`
    <span class="indeterminate-indicator-2" part="indeterminate-indicator-2"></span>
  `}),ht=[W`
:host .file-upload-area-button{width:auto;display:flex;align-items:end;justify-content:end;margin-inline-end:36px;margin-top:30px}:host .focus,:host :focus{outline:0}:host fluent-button .upload-icon path{fill:var(--file-upload-button-text-color,var(--foreground-on-accent-rest))}:host fluent-button.file-upload-button::part(control){border:var(--file-upload-button-border,none);background:var(--file-upload-button-background-color,var(--accent-fill-rest))}:host fluent-button.file-upload-button::part(control):hover{background:var(--file-upload-button-background-color-hover,var(--accent-fill-hover))}:host fluent-button.file-upload-button .upload-text{color:var(--file-upload-button-text-color,var(--foreground-on-accent-rest));font-weight:400;line-height:20px}:host input{display:none}:host fluent-progress.file-upload-bar{width:180px;margin-top:10px}:host fluent-dialog::part(overlay){opacity:.5}:host fluent-dialog::part(control){--dialog-width:$file-upload-dialog-width;--dialog-height:$file-upload-dialog-height;padding:var(--file-upload-dialog-padding,24px);border:var(--file-upload-dialog-border,1px solid var(--neutral-fill-rest))}:host fluent-dialog .file-upload-dialog-ok{background:var(--file-upload-dialog-keep-both-button-background-color,var(--accent-fill-rest));border:var(--file-upload-dialog-keep-both-button-border,none);color:var(--file-upload-dialog-keep-both-button-text-color,var(--foreground-on-accent-rest))}:host fluent-dialog .file-upload-dialog-ok:hover{background:var(--file-upload-dialog-keep-both-button-background-color-hover,var(--accent-fill-hover))}:host fluent-dialog .file-upload-dialog-cancel{background:var(--file-upload-dialog-replace-button-background-color,var(--accent-fill-rest));border:var(--file-upload-dialog-replace-button-border,1px solid var(--neutral-foreground-rest));color:var(--file-upload-dialog-replace-button-text-color,var(--neutral-foreground-rest))}:host fluent-dialog .file-upload-dialog-cancel:hover{background:var(--file-upload-dialog-replace-button-background-color-hover,var(--accent-fill-hover))}:host fluent-checkbox{margin-top:12px}:host fluent-checkbox .file-upload-dialog-check{color:var(--file-upload-dialog-text-color,--foreground-on-accent-rest)}:host .file-upload-table{display:flex}:host .file-upload-table.upload{display:flex}:host .file-upload-table .file-upload-cell{padding:1px 0 1px 1px;display:table-cell;vertical-align:middle;position:relative}:host .file-upload-table .file-upload-cell.percent-indicator{padding-inline-start:10px}:host .file-upload-table .file-upload-cell .description{opacity:.5;position:relative}:host .file-upload-table .file-upload-cell .file-upload-filename{max-width:250px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}:host .file-upload-table .file-upload-cell .file-upload-status{position:absolute;left:28px}:host .file-upload-table .file-upload-cell .file-upload-cancel{cursor:pointer;margin-inline-start:20px}:host .file-upload-table .file-upload-cell .file-upload-name{width:auto}:host .file-upload-table .file-upload-cell .cancel-icon{fill:var(--file-upload-dialog-text-color,var(--neutral-foreground-rest))}:host .mgt-file-item{--file-background-color:transparent;--file-padding:0 12px;--file-padding-inline-start:24px}:host .file-upload-template{clear:both}:host .file-upload-template .file-upload-folder-tab{padding-inline-start:20px}:host .file-upload-dialog{display:none}:host .file-upload-dialog .file-upload-dialog-content{background-color:var(--file-upload-dialog-background-color,var(--accent-fill-rest));color:var(--file-upload-dialog-text-color,var(--neutral-foreground-rest))}:host .file-upload-dialog .file-upload-dialog-content-text{margin-bottom:36px}:host .file-upload-dialog .file-upload-dialog-title{margin-top:0}:host .file-upload-dialog .file-upload-dialog-editor{display:flex;align-items:end;justify-content:end;gap:5px}:host .file-upload-dialog .file-upload-dialog-close{float:right;cursor:pointer}:host .file-upload-dialog .file-upload-dialog-close svg{fill:var(--file-upload-dialog-text-color,var(--neutral-foreground-rest));padding-right:5px}:host .file-upload-dialog.visible{display:block}:host fluent-checkbox.file-upload-dialog-check.hide{display:none}:host .file-upload-dialog-success{cursor:pointer;opacity:.5}:host #file-upload-border{display:none}:host #file-upload-border.visible{border:var(--file-upload-border-drag,1px dashed #0078d4);background-color:var(--file-upload-background-color-drag,rgba(0,120,212,.1));position:absolute;inset:0;z-index:1;display:inline-block}[dir=rtl] :host .file-upload-status{left:0;right:28px}@media (forced-colors:active){:host fluent-button .upload-icon path{fill:highlighttext}:host fluent-button.file-upload-button::part(control){border-width:1px;border-style:solid;border-color:buttontext;background:highlight}:host fluent-button.file-upload-button::part(control):hover{background:highlighttext;border-color:highlight}:host fluent-button.file-upload-button .upload-text{color:highlighttext}:host fluent-button:hover .upload-icon path{fill:highlight}:host fluent-button:hover.file-upload-button::part(control){border-color:highlight;background:highlighttext}:host fluent-button:hover.file-upload-button .upload-text{color:highlight}}
`],u={failUploadFile:"File upload failed",successUploadFile:"File upload succeeded",cancelUploadFile:"File cancel.",buttonUploadFile:"Upload Files",maximumFilesTitle:"Maximum files",maximumFiles:"Sorry, the maximum number of files you can upload at once is {MaxNumber}. Do you want to upload the first {MaxNumber} files or reselect?",maximumFileSizeTitle:"Maximum files size",maximumFileSize:'Sorry, the maximum file size to upload is {FileSize}. The file "{FileName}" has ',fileTypeTitle:"File type",fileType:'Sorry, the format of following file "{FileName}" cannot be uploaded.',checkAgain:"Don't show again",checkApplyAll:"Apply to all",buttonOk:"OK",buttonCancel:"Cancel",buttonUpload:"Upload",buttonKeep:"Keep both",buttonReplace:"Replace",buttonReselect:"Reselect",fileReplaceTitle:"Replace file",fileReplace:'Do you want to replace the file "{FileName}" or keep both files?',uploadButtonLabel:"File upload button"};var oe=function(s,e,t,i){var l=arguments.length,o=l<3?e:i===null?i=Object.getOwnPropertyDescriptor(e,t):i,a;if(typeof Reflect=="object"&&typeof Reflect.decorate=="function")o=Reflect.decorate(s,e,t,i);else for(var r=s.length-1;r>=0;r--)(a=s[r])&&(o=(l<3?a(o):l>3?a(e,t,o):a(e,t))||o);return l>3&&o&&Object.defineProperty(e,t,o),o},le=function(s,e){if(typeof Reflect=="object"&&typeof Reflect.metadata=="function")return Reflect.metadata(s,e)},F=function(s,e,t,i){function l(o){return o instanceof t?o:new t(function(a){a(o)})}return new(t||(t=Promise))(function(o,a){function r(d){try{c(i.next(d))}catch(h){a(h)}}function n(d){try{c(i.throw(d))}catch(h){a(h)}}function c(d){d.done?o(d.value):l(d.value).then(r,n)}c((i=i.apply(s,e||[])).next())})};const z=s=>s.isDirectory,ft=s=>s.isFile,gt=s=>"getAsEntry"in s&&typeof s.getAsEntry=="function",mt=s=>"getAsFile"in s&&typeof s.getAsFile=="function"||"webkitGetAsEntry"in s&&typeof s.webkitGetAsEntry=="function",vt=()=>{X(pt,be,Ke,dt),ee(),Z("file-upload",R)},bt=s=>(s==null?void 0:s.length)>1?s[1]==="replace"?"replace":"rename":null;class R extends ye{static get styles(){return ht}get strings(){return u}static get requiredScopes(){return[...new Set(["files.readwrite","files.readwrite.all","sites.readwrite.all"])]}get _dropEffect(){return"copy"}constructor(){super(),this._dragCounter=0,this._maxChunkSize=4*1024*1024,this._dialogTitle="",this._dialogContent="",this._dialogPrimaryButton="",this._dialogSecondaryButton="",this._dialogCheckBox="",this._applyAll=!1,this._applyAllConflictBehavior=null,this._maximumFileSize=!1,this._excludedFileType=!1,this.onFileUploadChange=e=>{const t=e.target;!e||t.files.length<1||this.readUploadedFiles(t.files,()=>t.value=null)},this.onFileUploadClick=()=>{this.renderRoot.querySelector("#file-upload-input").click()},this.handleonDragOver=e=>{e.preventDefault(),e.stopPropagation(),e.dataTransfer.items&&e.dataTransfer.items.length>0&&(e.dataTransfer.dropEffect=e.dataTransfer.dropEffect=this._dropEffect)},this.handleonDragEnter=e=>{e.preventDefault(),e.stopPropagation(),this._dragCounter++,e.dataTransfer.items&&e.dataTransfer.items.length>0&&(e.dataTransfer.dropEffect=this._dropEffect,this.renderRoot.querySelector("#file-upload-border").classList.add("visible"))},this.handleonDragLeave=e=>{e.preventDefault(),e.stopPropagation(),this._dragCounter--,this._dragCounter===0&&this.renderRoot.querySelector("#file-upload-border").classList.remove("visible")},this.handleonDrop=e=>{var t;e.preventDefault(),e.stopPropagation();const i=()=>{e.dataTransfer.clearData()};this.renderRoot.querySelector("#file-upload-border").classList.remove("visible"),!((t=e.dataTransfer)===null||t===void 0)&&t.items&&this.readUploadedFiles(e.dataTransfer.items,i),this._dragCounter=0},this.filesToUpload=[],this.addEventListener("__uploadfailed",this.focusOnUpload),this.addEventListener("__uploadsuccess",this.focusOnUpload)}focusOnUpload(){const e=this.renderRoot.querySelector('mgt-file[part="upload"]');e&&(e.setAttribute("tabindex","0"),e.classList.add("upload"),e.focus())}render(){if(this.parentElement!==null){const e=this.parentElement;e.addEventListener("dragenter",this.handleonDragEnter),e.addEventListener("dragleave",this.handleonDragLeave),e.addEventListener("dragover",this.handleonDragOver),e.addEventListener("drop",this.handleonDrop)}return v`
        <div id="file-upload-dialog" class="file-upload-dialog">
          <!-- Modal content -->
          <fluent-dialog modal="true" class="file-upload-dialog-content">
            <span
              class="file-upload-dialog-close"
              id="file-upload-dialog-close">
                ${S(k.Cancel)}
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
              <span slot="start">${S(k.Upload)}</span>
              <span class="upload-text">${this.strings.buttonUploadFile}</span>
          </fluent-button>
        </div>
        <div class="file-upload-template">
          ${this.renderFolderTemplate(this.filesToUpload)}
        </div>
       `}renderFolderTemplate(e){const t=[];if(e.length>0){const i=e.map(l=>t.indexOf(l.fullPath.substring(0,l.fullPath.lastIndexOf("/")))===-1?l.fullPath.endsWith("/")?v`${this.renderFileTemplate(l,"")}`:(t.push(l.fullPath.substring(0,l.fullPath.lastIndexOf("/"))),L`
            <div class='file-upload-table'>
              <div class='file-upload-cell'>
                <mgt-file
                  .fileDetails=${{name:l.fullPath.substring(1,l.fullPath.lastIndexOf("/")),folder:"Folder"}}
                  view="oneline"
                  class="mgt-file-item">
                </mgt-file>
              </div>
            </div>
            ${this.renderFileTemplate(l,"file-upload-folder-tab")}`):v`${this.renderFileTemplate(l,"file-upload-folder-tab")}`);return v`${i}`}return v``}renderFileTemplate(e,t){const i=E({"file-upload-table":!0,upload:e.completed}),l=t+(e.fieldUploadResponse==="lastModifiedDateTime"?" file-upload-dialog-success":""),o=e.fieldUploadResponse==="description",a=E({description:o}),r=e.completed?v``:this.renderFileUploadTemplate(e),n=o?u.failUploadFile:u.successUploadFile;return L`
        <div class="${i}">
          <div class="${l}">
            <div class='file-upload-cell'>
              <div class="${a}">
                <div class="file-upload-status">
                  ${e.iconStatus}
                </div>
                <mgt-file
                  .fileDetails=${e.driveItem}
                  .view=${e.view}
                  .line2Property=${e.fieldUploadResponse}
                  aria-label="${e.driveItem.name} ${n}"
                  part="upload"
                  class="mgt-file-item">
                </mgt-file>
              </div>
            </div>
              ${r}
            </div>
          </div>
        </div>`}renderFileUploadTemplate(e){const t=E({"file-upload-table":!0,upload:e.completed});return v`
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
                ${S(k.Cancel)}
              </span>
            </div>
          </div>
        </div>
      </div>
    </div>
    `}deleteFileUploadSession(e){return F(this,void 0,void 0,function*(){try{e.uploadUrl!==void 0?(yield we(this.fileUploadList.graph,e.uploadUrl),e.uploadUrl=void 0,e.completed=!0,this.setUploadFail(e,u.cancelUploadFile)):(e.uploadUrl=void 0,e.completed=!0,this.setUploadFail(e,u.cancelUploadFile))}catch{e.uploadUrl=void 0,e.completed=!0,this.setUploadFail(e,u.cancelUploadFile)}})}readUploadedFiles(e,t){return F(this,void 0,void 0,function*(){const i=yield this.getFilesFromUploadArea(e);yield this.getSelectedFiles(i),t()})}getSelectedFiles(e){return F(this,void 0,void 0,function*(){let t=[];const i=[];this._applyAll=!1,this._applyAllConflictBehavior=null,this._maximumFileSize=!1,this._excludedFileType=!1,this.filesToUpload.forEach(o=>{o.completed?i.push(o):t.push(o)});for(const o of e){const a=o.fullPath===""?"/"+o.name:o.fullPath;if(t.filter(r=>r.fullPath===a).length===0){let r=!0;if(this.fileUploadList.maxFileSize!==void 0&&r&&o.size>this.fileUploadList.maxFileSize*1024&&(r=!1,this._maximumFileSize===!1)){const n=yield this.getFileUploadStatus(o,a,"MaxFileSize",this.fileUploadList);n!==null&&n[0]===1&&(this._maximumFileSize=!0)}if(this.fileUploadList.excludedFileExtensions!==void 0&&this.fileUploadList.excludedFileExtensions.length>0&&r&&this.fileUploadList.excludedFileExtensions.filter(n=>o.name.toLowerCase().indexOf(n.toLowerCase())>-1).length>0&&(r=!1,this._excludedFileType===!1)){const n=yield this.getFileUploadStatus(o,a,"ExcludedFileType",this.fileUploadList);n!==null&&n[0]===1&&(this._excludedFileType=!0)}if(r){const n=yield this.getFileUploadStatus(o,a,"Upload",this.fileUploadList);let c=!1;n!==null&&(n[0]===-1?c=!0:(this._applyAll=!!n[0],this._applyAllConflictBehavior=n[1]?1:0)),t.push({file:o,driveItem:{name:o.name},fullPath:a,conflictBehavior:bt(n),iconStatus:null,percent:1,view:"image",completed:c,maxSize:this._maxChunkSize,minSize:0})}}}t=t.sort((o,a)=>o.fullPath.substring(0,o.fullPath.lastIndexOf("/")).localeCompare(a.fullPath.substring(0,a.fullPath.lastIndexOf("/")))),t.forEach(o=>{if(i.filter(a=>a.fullPath===o.fullPath).length!==0){const a=i.findIndex(r=>r.fullPath===o.fullPath);i.splice(a,1)}}),t.push(...i),this.filesToUpload=t;const l=this.filesToUpload.map(o=>this.sendFileItemGraph(o));yield Promise.all(l)})}getFileUploadStatus(e,t,i,l){const o=Object.create(null,{requestStateUpdate:{get:()=>super.requestStateUpdate}});return F(this,void 0,void 0,function*(){const a=this.renderRoot.querySelector("#file-upload-dialog");switch(i){case"Upload":return(yield Ue(this.fileUploadList.graph,`${this.getGrapQuery(t)}?$select=id`))!==null?this._applyAll===!0?[this._applyAll,this._applyAllConflictBehavior]:(a.classList.add("visible"),this._dialogTitle=u.fileReplaceTitle,this._dialogContent=u.fileReplace.replace("{FileName}",e.name),this._dialogCheckBox=u.checkApplyAll,this._dialogPrimaryButton=u.buttonReplace,this._dialogSecondaryButton=u.buttonKeep,yield o.requestStateUpdate.call(this,!0),new Promise(n=>{const c=this.renderRoot.querySelector(".file-upload-dialog-close"),d=this.renderRoot.querySelector(".file-upload-dialog-ok"),h=this.renderRoot.querySelector(".file-upload-dialog-cancel"),x=this.renderRoot.querySelector("#file-upload-dialog-check");x.checked=!1,x.classList.remove("hide");const b=()=>{a.classList.remove("visible"),n([x.checked?1:0,"replace"])},A=()=>{a.classList.remove("visible"),n([x.checked?1:0,"rename"])},O=()=>{a.classList.remove("visible"),n([-1])};d.removeEventListener("click",b),h.removeEventListener("click",A),c.removeEventListener("click",O),d.addEventListener("click",b),h.addEventListener("click",A),c.addEventListener("click",O)})):null;case"ExcludedFileType":return a.classList.add("visible"),this._dialogTitle=u.fileTypeTitle,this._dialogContent=u.fileType.replace("{FileName}",e.name)+" ("+l.excludedFileExtensions.join(",")+")",this._dialogCheckBox=u.checkAgain,this._dialogPrimaryButton=u.buttonOk,this._dialogSecondaryButton=u.buttonCancel,yield o.requestStateUpdate.call(this,!0),new Promise(r=>{const n=this.renderRoot.querySelector(".file-upload-dialog-ok"),c=this.renderRoot.querySelector(".file-upload-dialog-cancel"),d=this.renderRoot.querySelector(".file-upload-dialog-close"),h=this.renderRoot.querySelector("#file-upload-dialog-check");h.checked=!1,h.classList.remove("hide");const x=()=>{a.classList.remove("visible"),r([h.checked?1:0])},b=()=>{a.classList.remove("visible"),r([0])};n.removeEventListener("click",x),c.removeEventListener("click",b),d.removeEventListener("click",b),n.addEventListener("click",x),c.addEventListener("click",b),d.addEventListener("click",b)});case"MaxFileSize":return a.classList.add("visible"),this._dialogTitle=u.maximumFileSizeTitle,this._dialogContent=u.maximumFileSize.replace("{FileSize}",N(l.maxFileSize*1024)).replace("{FileName}",e.name)+N(e.size)+".",this._dialogCheckBox=u.checkAgain,this._dialogPrimaryButton=u.buttonOk,this._dialogSecondaryButton=u.buttonCancel,yield o.requestStateUpdate.call(this,!0),new Promise(r=>{const n=this.renderRoot.querySelector(".file-upload-dialog-ok"),c=this.renderRoot.querySelector(".file-upload-dialog-cancel"),d=this.renderRoot.querySelector(".file-upload-dialog-close"),h=this.renderRoot.querySelector("#file-upload-dialog-check");h.checked=!1,h.classList.remove("hide");const x=()=>{a.classList.remove("visible"),r([h.checked?1:0])},b=()=>{a.classList.remove("visible"),r([0])};n.removeEventListener("click",x),c.removeEventListener("click",b),d.removeEventListener("click",b),n.addEventListener("click",x),c.addEventListener("click",b),d.addEventListener("click",b)})}})}getGrapQuery(e){let t="";return this.fileUploadList.itemPath&&this.fileUploadList.itemPath.length>0&&(t=this.fileUploadList.itemPath.startsWith("/")?this.fileUploadList.itemPath:"/"+this.fileUploadList.itemPath),this.fileUploadList.userId&&this.fileUploadList.itemId?`/users/${this.fileUploadList.userId}/drive/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.userId&&this.fileUploadList.itemPath?`/users/${this.fileUploadList.userId}/drive/root:${t}${e}`:this.fileUploadList.groupId&&this.fileUploadList.itemId?`/groups/${this.fileUploadList.groupId}/drive/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.groupId&&this.fileUploadList.itemPath?`/groups/${this.fileUploadList.groupId}/drive/root:${t}${e}`:this.fileUploadList.driveId&&this.fileUploadList.itemId?`/drives/${this.fileUploadList.driveId}/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.driveId&&this.fileUploadList.itemPath?`/drives/${this.fileUploadList.driveId}/root:${t}${e}`:this.fileUploadList.siteId&&this.fileUploadList.itemId?`/sites/${this.fileUploadList.siteId}/drive/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.siteId&&this.fileUploadList.itemPath?`/sites/${this.fileUploadList.siteId}/drive/root:${t}${e}`:this.fileUploadList.itemId?`/me/drive/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.itemPath?`/me/drive/root:${t}${e}`:`/me/drive/root:${e}`}sendFileItemGraph(e){return F(this,void 0,void 0,function*(){const t=this.fileUploadList.graph;let i="";if(e.file.size<this._maxChunkSize)try{e.completed||((e.conflictBehavior===null||e.conflictBehavior==="replace")&&(i=`${this.getGrapQuery(e.fullPath)}:/content`),e.conflictBehavior==="rename"&&(i=`${this.getGrapQuery(e.fullPath)}:/content?@microsoft.graph.conflictBehavior=rename`),e.driveItem=yield $e(t,i,e.file),e.driveItem!==null?this.setUploadSuccess(e):(e.driveItem={name:e.file.name},this.setUploadFail(e,u.failUploadFile)))}catch{this.setUploadFail(e,u.failUploadFile)}else if(!e.completed&&e.uploadUrl===void 0){const l=yield Ce(t,`${this.getGrapQuery(e.fullPath)}:/createUploadSession`,e.conflictBehavior);try{if(l!==null){e.uploadUrl=l.uploadUrl;const o=yield this.sendSessionUrlGraph(t,e);o!==null?(e.driveItem=o,this.setUploadSuccess(e)):this.setUploadFail(e,u.failUploadFile)}else this.setUploadFail(e,u.failUploadFile)}catch{}}})}sendSessionUrlGraph(e,t){const i=Object.create(null,{requestStateUpdate:{get:()=>super.requestStateUpdate}});return F(this,void 0,void 0,function*(){for(;t.file.size>t.minSize;){t.mimeStreamString===void 0&&(t.mimeStreamString=yield this.readFileContent(t.file));const l=new Blob([t.mimeStreamString.slice(t.minSize,t.maxSize)]);if(t.percent=Math.round(t.maxSize/t.file.size*100),yield i.requestStateUpdate.call(this,!0),t.uploadUrl!==void 0){const o=yield _e(e,t.uploadUrl,`${t.maxSize-t.minSize}`,`bytes ${t.minSize}-${t.maxSize-1}/${t.file.size}`,l);if(o===null)return null;if(Le(o))t.minSize=parseInt(o.nextExpectedRanges[0].split("-")[0],10),t.maxSize=t.minSize+this._maxChunkSize,t.maxSize>t.file.size&&(t.maxSize=t.file.size);else if(o.id!==void 0)return o}else return null}})}setUploadSuccess(e){e.percent=100,super.requestStateUpdate(!0),setTimeout(()=>{e.iconStatus=S(k.Success),e.view="twolines",e.fieldUploadResponse="lastModifiedDateTime",e.completed=!0,super.requestStateUpdate(!0),J(),this.fireCustomEvent("__uploadsuccess")},500)}setUploadFail(e,t){setTimeout(()=>{e.iconStatus=S(k.Fail),e.view="twolines",e.driveItem.description=t,e.fieldUploadResponse="description",e.completed=!0,super.requestStateUpdate(!0),this.fireCustomEvent("__uploadfailed")},500)}readFileContent(e){return new Promise((t,i)=>{const l=new FileReader;l.onloadend=()=>{t(l.result)},l.onerror=o=>{i(o)},l.readAsArrayBuffer(e)})}getFilesFromUploadArea(e){return F(this,void 0,void 0,function*(){const t=[];let i;const l=[];for(const o of e)if(mt(o))if(gt(o))if(i=o.getAsEntry(),z(i))t.push(i);else{const a=o.getAsFile();a&&(this.writeFilePath(a,""),l.push(a))}else if(o.webkitGetAsEntry)if(i=o.webkitGetAsEntry(),z(i))t.push(i);else{const a=o.getAsFile();a&&(this.writeFilePath(a,""),l.push(a))}else{const a=o.getAsFile();a&&(this.writeFilePath(a,""),l.push(a))}else this.writeFilePath(o,""),l.push(o);if(t.length>0){const o=yield this.getFolderFiles(t);l.push(...o)}return l})}getFolderFiles(e){return new Promise(t=>{let i=0;const l=[];e.forEach(r=>{o(r,"")});const o=(r,n)=>{z(r)?a(r.createReader()):ft(r)&&(i++,r.file(c=>{i--,this.writeFilePath(c,n),l.push(c),i===0&&t(l)}))},a=r=>{i++,r.readEntries(n=>{i--;for(const c of n)o(c,c.fullPath);i===0&&t(l)})}})}writeFilePath(e,t){e.fullPath=t}}oe([f({type:Object}),le("design:type",Array)],R.prototype,"filesToUpload",void 0);oe([f({type:Object}),le("design:type",Object)],R.prototype,"fileUploadList",void 0);const yt=[W`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size, 14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{font-size:var(--default-font-size)}:host .title{font-size:14px;font-weight:600;padding:20px 0 12px;line-height:19px}:host .file-list-wrapper{background-color:var(--file-list-background-color,var(--neutral-layer-floating));border:var(--file-list-border,none);position:relative;display:flex;flex-direction:column;border-radius:8px}:host .file-list-wrapper .title{font-size:14px;font-weight:600;margin:0 20px -15px}:host .file-list-wrapper .file-list{display:flex;padding:var(--file-list-padding,0);margin:var(--file-list-margin,0);flex-direction:column;list-style:none}:host .file-list-wrapper .file-list .file-item{cursor:pointer;border-radius:var(--file-border-radius)}:host .file-list-wrapper .file-list .file-item:focus,:host .file-list-wrapper .file-list .file-item:focus-within{--file-background-color:var(--file-background-color-focus, var(--neutral-layer-2))}:host .file-list-wrapper .file-list .file-item.selected{--file-background-color:var(--file-background-color-active, var(--neutral-layer-3))}:host .file-list-wrapper .file-list .file-item .mgt-file-item{--file-padding:10px 20px 10px 20px;--file-padding-inline-start:24px;--file-border-radius:2px;--file-background-color:var(--file-item-background-color, var(--neutral-layer-1))}:host .file-list-wrapper .progress-ring{margin:4px auto;width:var(--progress-ring-size,24px);height:var(--progress-ring-size,24px)}:host .file-list-wrapper .show-more{text-align:center;font-size:var(--show-more-button-font-size, 12px);padding:var(--show-more-button-padding,0);border-radius:0 0 var(--show-more-button-border-bottom-right-radius,var(--file-list-border-radius,8px)) var(--show-more-button-border-bottom-left-radius,var(--file-list-border-radius,8px));background-color:var(--show-more-button-background-color,var(--neutral-stroke-divider-rest))}:host .file-list-wrapper .show-more:hover{background-color:var(--show-more-button-background-color-hover,var(--neutral-fill-input-alt-active))}:host .file-list-wrapper .show-more-text{font-size:var(--show-more-button-font-size, 12px)}
`],xt={showMoreSubtitle:"Show more items",filesSectionTitle:"Files",sharedTextSubtitle:"Shared"};var g=function(s,e,t,i){var l=arguments.length,o=l<3?e:i===null?i=Object.getOwnPropertyDescriptor(e,t):i,a;if(typeof Reflect=="object"&&typeof Reflect.decorate=="function")o=Reflect.decorate(s,e,t,i);else for(var r=s.length-1;r>=0;r--)(a=s[r])&&(o=(l<3?a(o):l>3?a(e,t,o):a(e,t))||o);return l>3&&o&&Object.defineProperty(e,t,o),o},m=function(s,e){if(typeof Reflect=="object"&&typeof Reflect.metadata=="function")return Reflect.metadata(s,e)},V=function(s,e,t,i){function l(o){return o instanceof t?o:new t(function(a){a(o)})}return new(t||(t=Promise))(function(o,a){function r(d){try{c(i.next(d))}catch(h){a(h)}}function n(d){try{c(i.throw(d))}catch(h){a(h)}}function c(d){d.done?o(d.value):l(d.value).then(r,n)}c((i=i.apply(s,e||[])).next())})};const Ct=()=>{X(Fe),ee(),vt(),Z("file-list",p)};class p extends xe{static get styles(){return yt}get strings(){return xt}get displayName(){return this.strings.filesSectionTitle}get cardTitle(){return this.strings.filesSectionTitle}renderIcon(){return S(k.Files)}static get requiredScopes(){return[...new Set([...Ge.requiredScopes])]}constructor(){super(),this._isCompact=!1,this.fileQueries=null,this.files=null,this.itemView="threelines",this.fileExtensions=[],this.pageSize=10,this.disableOpenOnClick=!1,this.hideMoreFilesButton=!1,this.enableFileUpload=!1,this.maxUploadFile=10,this.excludedFileExtensions=[],this._preloadedFiles=[],this._focusedItemIndex=-1,this.renderLoading=()=>this.files?this.renderContent():this.renderTemplate("loading",null)||v``,this.renderContent=()=>!this.files||this.files.length===0?this.renderNoData():this._isCompact?this.renderCompactView():this.renderFullView(),this.onFocusFirstItem=()=>this._focusedItemIndex=0,this.onFileListKeyDown=e=>{const t=this.renderRoot.querySelector(".file-list");let i;if(t!=null&&t.children.length){if((e.code==="ArrowUp"||e.code==="ArrowDown")&&(e.code==="ArrowUp"&&(this._focusedItemIndex===-1&&(this._focusedItemIndex=t.children.length),this._focusedItemIndex=(this._focusedItemIndex-1+t.children.length)%t.children.length),e.code==="ArrowDown"&&(this._focusedItemIndex=(this._focusedItemIndex+1)%t.children.length),i=t.children[this._focusedItemIndex],this.updateItemBackgroundColor(t,i,"focused")),e.code==="Enter"||e.code==="Space"){i=t.children[this._focusedItemIndex];const l=i.children[0];e.preventDefault(),this.fireCustomEvent("itemClick",l.fileDetails),this.handleFileClick(l.fileDetails),this.updateItemBackgroundColor(t,i,"selected")}e.code==="Tab"&&(i=t.children[this._focusedItemIndex])}}}clearState(){super.clearState(),this.files=null}asCompactView(){return this._isCompact=!0,this}asFullView(){return this._isCompact=!1,this}args(){return[this.providerState,this.fileListQuery,this.fileQueries,this.siteId,this.driveId,this.groupId,this.itemId,this.itemPath,this.userId,this.insightType,this.fileExtensions,this.pageSize,this.maxFileSize]}renderCompactView(){const e=this.files.slice(0,3);return this.renderFiles(e)}renderFullView(){return this.renderTemplate("default",{files:this.files})||this.renderFiles(this.files)}renderNoData(){return this.renderTemplate("no-data",null)||(this.enableFileUpload===!0&&D.globalProvider!==void 0?v`
            <div class="file-list-wrapper" dir=${this.direction}>
              ${this.renderFileUpload()}
            </div>`:v``)}renderFiles(e){return v`
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
          ${ke(e.slice(1),t=>t.id,t=>v`
              <li
                class="file-item"
                @keydown="${this.onFileListKeyDown}"
                @click=${i=>this.handleItemSelect(t,i)}>
                ${this.renderFile(t)}
              </li>
            `)}
        </ul>
        ${!this.hideMoreFilesButton&&this.pageIterator&&(this.pageIterator.hasNext||this._preloadedFiles.length)&&!this._isCompact?this.renderMoreFileButton():null}
      </div>
    `}renderFile(e){const t=this.itemView;return this.renderTemplate("file",{file:e},e.id)||L`
        <mgt-file class="mgt-file-item" .fileDetails=${e} .view=${t}></mgt-file>
      `}renderMoreFileButton(){return this._isLoadingMore?v`
        <fluent-progress-ring role="progressbar" viewBox="0 0 8 8" class="progress-ring"></fluent-progress-ring>
      `:v`
        <fluent-button
          appearance="stealth"
          id="show-more"
          class="show-more"
          @click=${()=>this.renderNextPage()}
        >
          <span class="show-more-text">${this.strings.showMoreSubtitle}</span>
        </fluent-button>`}renderFileUpload(){const e={graph:D.globalProvider.graph.forComponent(this),driveId:this.driveId,excludedFileExtensions:this.excludedFileExtensions,groupId:this.groupId,itemId:this.itemId,itemPath:this.itemPath,userId:this.userId,siteId:this.siteId,maxFileSize:this.maxFileSize,maxUploadFile:this.maxUploadFile};return L`
        <mgt-file-upload .fileUploadList=${e} ></mgt-file-upload>
      `}loadState(){var e,t;return V(this,void 0,void 0,function*(){const i=D.globalProvider;if(!i||i.state===Q.Loading)return;if(i.state===Q.SignedOut){this.files=null;return}const l=i.graph.forComponent(this);let o,a;const r=!this.driveId&&!this.siteId&&!this.groupId&&!this.userId;if((this.driveId&&!this.itemId&&!this.itemPath||this.groupId&&!this.itemId&&!this.itemPath||this.siteId&&!this.itemId&&!this.itemPath||this.userId&&!this.insightType&&!this.itemId&&!this.itemPath)&&(this.files=null),!this.files){this.fileListQuery?a=yield Te(l,this.fileListQuery,this.pageSize):this.fileQueries?o=yield Ie(l,this.fileQueries):r?this.itemId?a=yield Ee(l,this.itemId,this.pageSize):this.itemPath?a=yield De(l,this.itemPath,this.pageSize):this.insightType?o=yield ze(l,this.insightType):a=yield Pe(l,this.pageSize):this.driveId?this.itemId?a=yield Be(l,this.driveId,this.itemId,this.pageSize):this.itemPath&&(a=yield Re(l,this.driveId,this.itemPath,this.pageSize)):this.groupId?this.itemId?a=yield Ae(l,this.groupId,this.itemId,this.pageSize):this.itemPath&&(a=yield Oe(l,this.groupId,this.itemPath,this.pageSize)):this.siteId?this.itemId?a=yield Me(l,this.siteId,this.itemId,this.pageSize):this.itemPath&&(a=yield qe(l,this.siteId,this.itemPath,this.pageSize)):this.userId&&(this.itemId?a=yield Ne(l,this.userId,this.itemId,this.pageSize):this.itemPath?a=yield je(l,this.userId,this.itemPath,this.pageSize):this.insightType&&(o=yield Qe(l,this.userId,this.insightType))),a&&(this.pageIterator=a,this._preloadedFiles=[...this.pageIterator.value],this._preloadedFiles.length>=this.pageSize?o=this._preloadedFiles.splice(0,this.pageSize):o=this._preloadedFiles.splice(0,this._preloadedFiles.length));let n;if(((e=this.fileExtensions)===null||e===void 0?void 0:e.length)>0){if(!((t=this.pageIterator)===null||t===void 0)&&t.value){for(;this.pageIterator.hasNext;)yield j(this.pageIterator);o=this.pageIterator.value,this._preloadedFiles=[]}n=o.filter(c=>{for(const d of this.fileExtensions)if(d===this.getFileExtension(c.name))return c})}(n==null?void 0:n.length)>=0?(this.files=n,this.pageSize&&(o=this.files.splice(0,this.pageSize),this.files=o)):this.files=o}})}handleItemSelect(e,t){if(this.handleFileClick(e),this.fireCustomEvent("itemClick",e),t){const i=this.renderRoot.querySelector(".file-list"),l=Array.from(i.children),o=t.target.closest("li"),a=l.indexOf(o);this._focusedItemIndex=a;const r=i.children[this._focusedItemIndex];this.updateItemBackgroundColor(i,r,"selected")}}renderNextPage(){return V(this,void 0,void 0,function*(){if(this._preloadedFiles.length>0)this.files=[...this.files,...this._preloadedFiles.splice(0,Math.min(this.pageSize,this._preloadedFiles.length))];else if(this.pageIterator.hasNext){this._isLoadingMore=!0;const e=this.renderRoot.querySelector("#file-list-wrapper");e!=null&&e.animate&&e.animate([{height:"auto",transformOrigin:"top left"},{height:"auto",transformOrigin:"top left"}],{duration:1e3,easing:"ease-in-out",fill:"both"}),yield j(this.pageIterator),this._isLoadingMore=!1,this.files=this.pageIterator.value}this.requestUpdate()})}handleFileClick(e){e!=null&&e.webUrl&&!this.disableOpenOnClick&&window.open(e.webUrl,"_blank","noreferrer")}getFileExtension(e){return/(?:\.([^.]+))?$/.exec(e)[1]||""}updateItemBackgroundColor(e,t,i){for(const l of e.children)l.classList.remove(i),l.removeAttribute("tabindex");t&&(t.classList.add(i),t.scrollIntoView({behavior:"smooth",block:"nearest",inline:"start"}),t.setAttribute("tabindex","0"),t.focus());for(const l of e.children)l.classList.remove("selected")}reload(e=!1){e&&J(),this._task.run()}}g([Y(),m("design:type",Object)],p.prototype,"_isCompact",void 0);g([f({attribute:"file-list-query"}),m("design:type",String)],p.prototype,"fileListQuery",void 0);g([f({attribute:"file-queries",converter:(s,e)=>s?s.split(",").map(t=>t.trim()):null}),m("design:type",Array)],p.prototype,"fileQueries",void 0);g([f({type:Object}),m("design:type",Array)],p.prototype,"files",void 0);g([f({attribute:"site-id"}),m("design:type",String)],p.prototype,"siteId",void 0);g([f({attribute:"drive-id"}),m("design:type",String)],p.prototype,"driveId",void 0);g([f({attribute:"group-id"}),m("design:type",String)],p.prototype,"groupId",void 0);g([f({attribute:"item-id"}),m("design:type",String)],p.prototype,"itemId",void 0);g([f({attribute:"item-path"}),m("design:type",String)],p.prototype,"itemPath",void 0);g([f({attribute:"user-id"}),m("design:type",String)],p.prototype,"userId",void 0);g([f({attribute:"insight-type"}),m("design:type",String)],p.prototype,"insightType",void 0);g([f({attribute:"item-view",converter:s=>Se(s,"threelines")}),m("design:type",String)],p.prototype,"itemView",void 0);g([f({attribute:"file-extensions",converter:(s,e)=>s.split(",").map(t=>t.trim())}),m("design:type",Array)],p.prototype,"fileExtensions",void 0);g([f({attribute:"page-size",type:Number}),m("design:type",Object)],p.prototype,"pageSize",void 0);g([f({attribute:"disable-open-on-click",type:Boolean}),m("design:type",Object)],p.prototype,"disableOpenOnClick",void 0);g([f({attribute:"hide-more-files-button",type:Boolean}),m("design:type",Object)],p.prototype,"hideMoreFilesButton",void 0);g([f({attribute:"max-file-size",type:Number}),m("design:type",Number)],p.prototype,"maxFileSize",void 0);g([f({attribute:"enable-file-upload",type:Boolean}),m("design:type",Object)],p.prototype,"enableFileUpload",void 0);g([f({attribute:"max-upload-file",type:Number}),m("design:type",Object)],p.prototype,"maxUploadFile",void 0);g([f({attribute:"excluded-file-extensions",converter:(s,e)=>s.split(",").map(t=>t.trim())}),m("design:type",Array)],p.prototype,"excludedFileExtensions",void 0);g([Y(),m("design:type",Boolean)],p.prototype,"_isLoadingMore",void 0);export{p as M,Ct as r};
