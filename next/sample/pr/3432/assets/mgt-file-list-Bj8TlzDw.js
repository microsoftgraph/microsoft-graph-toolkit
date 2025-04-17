import{m as C,r as Y,aF as ne,_,b as k,F as de,cB as ce,p as ue,D as pe,O as he,y as R,cC as fe,cd as ge,N as me,J as $,z as ve,cD as be,E as U,M as E,L as q,Q as ye,R as N,cE as xe,X as J,a0 as ee,a1 as Fe,a3 as te,cF as Se,$ as v,a4 as S,a5 as w,a8 as T,a6 as z,cG as we,cH as j,cI as _e,cJ as ke,cK as Ue,cL as $e,cM as Ce,cN as ie,Y as m,a2 as A,Z as Le,aQ as Te,cO as Ie,cP as Ee,cQ as ze,cR as De,cS as Pe,cT as Q,cU as G,cV as Re,cW as Be,cX as Ae,cY as Oe,cZ as Me,c_ as qe,c$ as Ne,d0 as je,d1 as V,at as Qe,d2 as Ge,d3 as Ve}from"./App-B6S7UShP.js";import{c as Ke}from"./repeat-B9aMobyH.js";import{r as le,M as He,g as Xe}from"./mgt-file-C6LDyJdu.js";import{a as We}from"./index-D41GWJEh.js";import{c as K,b as D,a as H}from"./index-CUIyG1wb.js";const Ze=(a,e)=>C`
    <div class="positioning-region" part="positioning-region">
        ${Y(t=>t.modal,C`
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
            ${ne("dialog")}
        >
            <slot></slot>
        </div>
    </div>
`;/*!
* tabbable 5.3.3
* @license MIT, https://github.com/focus-trap/tabbable/blob/master/LICENSE
*/var Ye=["input","select","textarea","a[href]","button","[tabindex]:not(slot)","audio[controls]","video[controls]",'[contenteditable]:not([contenteditable="false"])',"details>summary:first-of-type","details"],Je=Ye.join(","),oe=typeof Element>"u",I=oe?function(){}:Element.prototype.matches||Element.prototype.msMatchesSelector||Element.prototype.webkitMatchesSelector,B=!oe&&Element.prototype.getRootNode?function(a){return a.getRootNode()}:function(a){return a.ownerDocument},et=function(e,t){return e.tabIndex<0&&(/^(AUDIO|VIDEO|DETAILS)$/.test(e.tagName)||e.isContentEditable)&&isNaN(parseInt(e.getAttribute("tabindex"),10))?0:e.tabIndex},se=function(e){return e.tagName==="INPUT"},tt=function(e){return se(e)&&e.type==="hidden"},it=function(e){var t=e.tagName==="DETAILS"&&Array.prototype.slice.apply(e.children).some(function(i){return i.tagName==="SUMMARY"});return t},lt=function(e,t){for(var i=0;i<e.length;i++)if(e[i].checked&&e[i].form===t)return e[i]},ot=function(e){if(!e.name)return!0;var t=e.form||B(e),i=function(r){return t.querySelectorAll('input[type="radio"][name="'+r+'"]')},o;if(typeof window<"u"&&typeof window.CSS<"u"&&typeof window.CSS.escape=="function")o=i(window.CSS.escape(e.name));else try{o=i(e.name)}catch(s){return console.error("Looks like you have a radio button with a name attribute containing invalid CSS selector characters and need the CSS.escape polyfill: %s",s.message),!1}var l=lt(o,e.form);return!l||l===e},st=function(e){return se(e)&&e.type==="radio"},at=function(e){return st(e)&&!ot(e)},X=function(e){var t=e.getBoundingClientRect(),i=t.width,o=t.height;return i===0&&o===0},rt=function(e,t){var i=t.displayCheck,o=t.getShadowRoot;if(getComputedStyle(e).visibility==="hidden")return!0;var l=I.call(e,"details>summary:first-of-type"),s=l?e.parentElement:e;if(I.call(s,"details:not([open]) *"))return!0;var r=B(e).host,n=(r==null?void 0:r.ownerDocument.contains(r))||e.ownerDocument.contains(e);if(!i||i==="full"){if(typeof o=="function"){for(var u=e;e;){var d=e.parentElement,c=B(e);if(d&&!d.shadowRoot&&o(d)===!0)return X(e);e.assignedSlot?e=e.assignedSlot:!d&&c!==e.ownerDocument?e=c.host:e=d}e=u}if(n)return!e.getClientRects().length}else if(i==="non-zero-area")return X(e);return!1},nt=function(e){if(/^(INPUT|BUTTON|SELECT|TEXTAREA)$/.test(e.tagName))for(var t=e.parentElement;t;){if(t.tagName==="FIELDSET"&&t.disabled){for(var i=0;i<t.children.length;i++){var o=t.children.item(i);if(o.tagName==="LEGEND")return I.call(t,"fieldset[disabled] *")?!0:!o.contains(e)}return!0}t=t.parentElement}return!1},dt=function(e,t){return!(t.disabled||tt(t)||rt(t,e)||it(t)||nt(t))},ct=function(e,t){return!(at(t)||et(t)<0||!dt(e,t))},W=function(e,t){if(t=t||{},!e)throw new Error("No node provided");return I.call(e,Je)===!1?!1:ct(t,e)};class x extends de{constructor(){super(...arguments),this.modal=!0,this.hidden=!1,this.trapFocus=!0,this.trapFocusChanged=()=>{this.$fastController.isConnected&&this.updateTrapFocus()},this.isTrappingFocus=!1,this.handleDocumentKeydown=e=>{if(!e.defaultPrevented&&!this.hidden)switch(e.key){case ue:this.dismiss(),e.preventDefault();break;case ce:this.handleTabKeyDown(e);break}},this.handleDocumentFocus=e=>{!e.defaultPrevented&&this.shouldForceFocus(e.target)&&(this.focusFirstElement(),e.preventDefault())},this.handleTabKeyDown=e=>{if(!this.trapFocus||this.hidden)return;const t=this.getTabQueueBounds();if(t.length!==0){if(t.length===1){t[0].focus(),e.preventDefault();return}e.shiftKey&&e.target===t[0]?(t[t.length-1].focus(),e.preventDefault()):!e.shiftKey&&e.target===t[t.length-1]&&(t[0].focus(),e.preventDefault())}},this.getTabQueueBounds=()=>{const e=[];return x.reduceTabbableItems(e,this)},this.focusFirstElement=()=>{const e=this.getTabQueueBounds();e.length>0?e[0].focus():this.dialog instanceof HTMLElement&&this.dialog.focus()},this.shouldForceFocus=e=>this.isTrappingFocus&&!this.contains(e),this.shouldTrapFocus=()=>this.trapFocus&&!this.hidden,this.updateTrapFocus=e=>{const t=e===void 0?this.shouldTrapFocus():e;t&&!this.isTrappingFocus?(this.isTrappingFocus=!0,document.addEventListener("focusin",this.handleDocumentFocus),pe.queueUpdate(()=>{this.shouldForceFocus(document.activeElement)&&this.focusFirstElement()})):!t&&this.isTrappingFocus&&(this.isTrappingFocus=!1,document.removeEventListener("focusin",this.handleDocumentFocus))}}dismiss(){this.$emit("dismiss"),this.$emit("cancel")}show(){this.hidden=!1}hide(){this.hidden=!0,this.$emit("close")}connectedCallback(){super.connectedCallback(),document.addEventListener("keydown",this.handleDocumentKeydown),this.notifier=he.getNotifier(this),this.notifier.subscribe(this,"hidden"),this.updateTrapFocus()}disconnectedCallback(){super.disconnectedCallback(),document.removeEventListener("keydown",this.handleDocumentKeydown),this.updateTrapFocus(!1),this.notifier.unsubscribe(this,"hidden")}handleChange(e,t){switch(t){case"hidden":this.updateTrapFocus();break}}static reduceTabbableItems(e,t){return t.getAttribute("tabindex")==="-1"?e:W(t)||x.isFocusableFastElement(t)&&x.hasTabbableShadow(t)?(e.push(t),e):t.childElementCount?e.concat(Array.from(t.children).reduce(x.reduceTabbableItems,[])):e}static isFocusableFastElement(e){var t,i;return!!(!((i=(t=e.$fastController)===null||t===void 0?void 0:t.definition.shadowOptions)===null||i===void 0)&&i.delegatesFocus)}static hasTabbableShadow(e){var t,i;return Array.from((i=(t=e.shadowRoot)===null||t===void 0?void 0:t.querySelectorAll("*"))!==null&&i!==void 0?i:[]).some(o=>W(o))}}_([k({mode:"boolean"})],x.prototype,"modal",void 0);_([k({mode:"boolean"})],x.prototype,"hidden",void 0);_([k({attribute:"trap-focus",mode:"boolean"})],x.prototype,"trapFocus",void 0);_([k({attribute:"aria-describedby"})],x.prototype,"ariaDescribedby",void 0);_([k({attribute:"aria-labelledby"})],x.prototype,"ariaLabelledby",void 0);_([k({attribute:"aria-label"})],x.prototype,"ariaLabel",void 0);const ut=(a,e)=>C`
    <template
        role="progressbar"
        aria-valuenow="${t=>t.value}"
        aria-valuemin="${t=>t.min}"
        aria-valuemax="${t=>t.max}"
        class="${t=>t.paused?"paused":""}"
    >
        ${Y(t=>typeof t.value=="number",C`
                <div class="progress" part="progress" slot="determinate">
                    <div
                        class="determinate"
                        part="determinate"
                        style="width: ${t=>t.percentComplete}%"
                    ></div>
                </div>
            `,C`
                <div class="progress" part="progress" slot="indeterminate">
                    <slot class="indeterminate" name="indeterminate">
                        ${e.indeterminateIndicator1||""}
                        ${e.indeterminateIndicator2||""}
                    </slot>
                </div>
            `)}
    </template>
`,pt=(a,e)=>R`
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
    box-shadow: ${fe};
    margin-top: auto;
    margin-bottom: auto;
    border-radius: calc(${ge} * 1px);
    width: var(--dialog-width);
    height: var(--dialog-height);
    background: ${me};
    z-index: 1;
    border: calc(${$} * 1px) solid transparent;
  }
`,ht=x.compose({baseName:"dialog",template:Ze,styles:pt}),ft=(a,e)=>R`
    ${ve("flex")} :host {
      align-items: center;
      height: calc((${$} * 3) * 1px);
    }

    .progress {
      background-color: ${be};
      border-radius: calc(${U} * 1px);
      width: 100%;
      height: calc(${$} * 1px);
      display: flex;
      align-items: center;
      position: relative;
    }

    .determinate {
      background-color: ${E};
      border-radius: calc(${U} * 1px);
      height: calc((${$} * 3) * 1px);
      transition: all 0.2s ease-in-out;
      display: flex;
    }

    .indeterminate {
      height: calc((${$} * 3) * 1px);
      border-radius: calc(${U} * 1px);
      display: flex;
      width: 100%;
      position: relative;
      overflow: hidden;
    }

    .indeterminate-indicator-1 {
      position: absolute;
      opacity: 0;
      height: 100%;
      background-color: ${E};
      border-radius: calc(${U} * 1px);
      animation-timing-function: cubic-bezier(0.4, 0, 0.6, 1);
      width: 40%;
      animation: indeterminate-1 2s infinite;
    }

    .indeterminate-indicator-2 {
      position: absolute;
      opacity: 0;
      height: 100%;
      background-color: ${E};
      border-radius: calc(${U} * 1px);
      animation-timing-function: cubic-bezier(0.4, 0, 0.6, 1);
      width: 60%;
      animation: indeterminate-2 2s infinite;
    }

    :host(.paused) .indeterminate-indicator-1,
    :host(.paused) .indeterminate-indicator-2 {
      animation: none;
      background-color: ${q};
      width: 100%;
      opacity: 1;
    }

    :host(.paused) .determinate {
      background-color: ${q};
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
  `.withBehaviors(ye(R`
        .indeterminate-indicator-1,
        .indeterminate-indicator-2,
        .determinate,
        .progress {
          background-color: ${N.ButtonText};
        }
        :host(.paused) .indeterminate-indicator-1,
        :host(.paused) .indeterminate-indicator-2,
        :host(.paused) .determinate {
          background-color: ${N.GrayText};
        }
      `));class gt extends xe{}const mt=gt.compose({baseName:"progress",template:ut,styles:ft,indeterminateIndicator1:`
    <span class="indeterminate-indicator-1" part="indeterminate-indicator-1"></span>
  `,indeterminateIndicator2:`
    <span class="indeterminate-indicator-2" part="indeterminate-indicator-2"></span>
  `}),vt=[J`
:host .file-upload-area-button{width:auto;display:flex;align-items:end;justify-content:end;margin-inline-end:36px;margin-top:30px}:host .focus,:host :focus{outline:0}:host fluent-button .upload-icon path{fill:var(--file-upload-button-text-color,var(--foreground-on-accent-rest))}:host fluent-button.file-upload-button::part(control){border:var(--file-upload-button-border,none);background:var(--file-upload-button-background-color,var(--accent-fill-rest))}:host fluent-button.file-upload-button::part(control):hover{background:var(--file-upload-button-background-color-hover,var(--accent-fill-hover))}:host fluent-button.file-upload-button .upload-text{color:var(--file-upload-button-text-color,var(--foreground-on-accent-rest));font-weight:400;line-height:20px}:host input{display:none}:host fluent-progress.file-upload-bar{width:180px;margin-top:10px}:host fluent-dialog::part(overlay){opacity:.5}:host fluent-dialog::part(control){--dialog-width:$file-upload-dialog-width;--dialog-height:$file-upload-dialog-height;padding:var(--file-upload-dialog-padding,24px);border:var(--file-upload-dialog-border,1px solid var(--neutral-fill-rest))}:host fluent-dialog .file-upload-dialog-ok{background:var(--file-upload-dialog-keep-both-button-background-color,var(--accent-fill-rest));border:var(--file-upload-dialog-keep-both-button-border,none);color:var(--file-upload-dialog-keep-both-button-text-color,var(--foreground-on-accent-rest))}:host fluent-dialog .file-upload-dialog-ok:hover{background:var(--file-upload-dialog-keep-both-button-background-color-hover,var(--accent-fill-hover))}:host fluent-dialog .file-upload-dialog-cancel{background:var(--file-upload-dialog-replace-button-background-color,var(--accent-fill-rest));border:var(--file-upload-dialog-replace-button-border,1px solid var(--neutral-foreground-rest));color:var(--file-upload-dialog-replace-button-text-color,var(--neutral-foreground-rest))}:host fluent-dialog .file-upload-dialog-cancel:hover{background:var(--file-upload-dialog-replace-button-background-color-hover,var(--accent-fill-hover))}:host fluent-checkbox{margin-top:12px}:host fluent-checkbox .file-upload-dialog-check{color:var(--file-upload-dialog-text-color,--foreground-on-accent-rest)}:host .file-upload-table{display:flex}:host .file-upload-table.upload{display:flex}:host .file-upload-table .file-upload-cell{padding:1px 0 1px 1px;display:table-cell;vertical-align:middle;position:relative}:host .file-upload-table .file-upload-cell.percent-indicator{padding-inline-start:10px}:host .file-upload-table .file-upload-cell .description{opacity:.5;position:relative}:host .file-upload-table .file-upload-cell .file-upload-filename{max-width:250px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}:host .file-upload-table .file-upload-cell .file-upload-status{position:absolute;left:28px}:host .file-upload-table .file-upload-cell .file-upload-cancel{cursor:pointer;margin-inline-start:20px}:host .file-upload-table .file-upload-cell .file-upload-name{width:auto}:host .file-upload-table .file-upload-cell .cancel-icon{fill:var(--file-upload-dialog-text-color,var(--neutral-foreground-rest))}:host .mgt-file-item{--file-background-color:transparent;--file-padding:0 12px;--file-padding-inline-start:24px}:host .file-upload-template{clear:both}:host .file-upload-template .file-upload-folder-tab{padding-inline-start:20px}:host .file-upload-dialog{display:none}:host .file-upload-dialog .file-upload-dialog-content{background-color:var(--file-upload-dialog-background-color,var(--accent-fill-rest));color:var(--file-upload-dialog-text-color,var(--neutral-foreground-rest));z-index:1;position:relative}:host .file-upload-dialog .file-upload-dialog-content-text{margin-bottom:36px}:host .file-upload-dialog .file-upload-dialog-title{margin-top:0}:host .file-upload-dialog .file-upload-dialog-editor{display:flex;align-items:end;justify-content:end;gap:5px}:host .file-upload-dialog .file-upload-dialog-close{float:right;cursor:pointer}:host .file-upload-dialog .file-upload-dialog-close svg{fill:var(--file-upload-dialog-text-color,var(--neutral-foreground-rest));padding-right:5px}:host .file-upload-dialog.visible{display:block}:host fluent-checkbox.file-upload-dialog-check.hide{display:none}:host .file-upload-dialog-success{cursor:pointer;opacity:.5}:host #file-upload-border{display:none}:host #file-upload-border.visible{border:var(--file-upload-border-drag,1px dashed #0078d4);background-color:var(--file-upload-background-color-drag,rgba(0,120,212,.1));position:absolute;inset:0;z-index:1;display:inline-block}[dir=rtl] :host .file-upload-status{left:0;right:28px}@media (forced-colors:active){:host fluent-button .upload-icon path{fill:highlighttext}:host fluent-button.file-upload-button::part(control){border-width:1px;border-style:solid;border-color:buttontext;background:highlight}:host fluent-button.file-upload-button::part(control):hover{background:highlighttext;border-color:highlight}:host fluent-button.file-upload-button .upload-text{color:highlighttext}:host fluent-button:hover .upload-icon path{fill:highlight}:host fluent-button:hover.file-upload-button::part(control){border-color:highlight;background:highlighttext}:host fluent-button:hover.file-upload-button .upload-text{color:highlight}}
`],p={failUploadFile:"File upload failed",successUploadFile:"File upload succeeded",cancelUploadFile:"File cancel.",buttonUploadFile:"Upload Files",maximumFilesTitle:"Maximum files",maximumFiles:"Sorry, the maximum number of files you can upload at once is {MaxNumber}. Do you want to upload the first {MaxNumber} files or reselect?",maximumFileSizeTitle:"Maximum files size",maximumFileSize:'Sorry, the maximum file size to upload is {FileSize}. The file "{FileName}" has ',fileTypeTitle:"File type",fileType:'Sorry, the format of following file "{FileName}" cannot be uploaded.',checkAgain:"Don't show again",checkApplyAll:"Apply to all",buttonOk:"OK",buttonCancel:"Cancel",buttonUpload:"Upload",buttonKeep:"Keep both",buttonReplace:"Replace",buttonReselect:"Reselect",fileReplaceTitle:"Replace file",fileReplace:'Do you want to replace the file "{FileName}" or keep both files?',uploadButtonLabel:"File upload button"};var ae=function(a,e,t,i){var o=arguments.length,l=o<3?e:i===null?i=Object.getOwnPropertyDescriptor(e,t):i,s;if(typeof Reflect=="object"&&typeof Reflect.decorate=="function")l=Reflect.decorate(a,e,t,i);else for(var r=a.length-1;r>=0;r--)(s=a[r])&&(l=(o<3?s(l):o>3?s(e,t,l):s(e,t))||l);return o>3&&l&&Object.defineProperty(e,t,l),l},re=function(a,e){if(typeof Reflect=="object"&&typeof Reflect.metadata=="function")return Reflect.metadata(a,e)},F=function(a,e,t,i){function o(l){return l instanceof t?l:new t(function(s){s(l)})}return new(t||(t=Promise))(function(l,s){function r(d){try{u(i.next(d))}catch(c){s(c)}}function n(d){try{u(i.throw(d))}catch(c){s(c)}}function u(d){d.done?l(d.value):o(d.value).then(r,n)}u((i=i.apply(a,e||[])).next())})};const P=a=>a.isDirectory,bt=a=>a.isFile,yt=a=>"getAsEntry"in a&&typeof a.getAsEntry=="function",xt=a=>"getAsFile"in a&&typeof a.getAsFile=="function"||"webkitGetAsEntry"in a&&typeof a.webkitGetAsEntry=="function",Ft=()=>{ee(mt,Fe,We,ht),le(),te("file-upload",O)},St=a=>(a==null?void 0:a.length)>1?a[1]==="replace"?"replace":"rename":null;class O extends Se{static get styles(){return vt}get strings(){return p}static get requiredScopes(){return[...new Set(["files.readwrite","files.readwrite.all","sites.readwrite.all"])]}get _dropEffect(){return"copy"}constructor(){super(),this._dragCounter=0,this._maxChunkSize=4*1024*1024,this._dialogTitle="",this._dialogContent="",this._dialogPrimaryButton="",this._dialogSecondaryButton="",this._dialogCheckBox="",this._applyAll=!1,this._applyAllConflictBehavior=null,this._maximumFileSize=!1,this._excludedFileType=!1,this.onFileUploadChange=e=>{const t=e.target;!e||t.files.length<1||this.readUploadedFiles(t.files,()=>t.value=null)},this.onFileUploadClick=()=>{this.renderRoot.querySelector("#file-upload-input").click()},this.handleonDragOver=e=>{e.preventDefault(),e.stopPropagation(),e.dataTransfer.items&&e.dataTransfer.items.length>0&&(e.dataTransfer.dropEffect=e.dataTransfer.dropEffect=this._dropEffect)},this.handleonDragEnter=e=>{e.preventDefault(),e.stopPropagation(),this._dragCounter++,e.dataTransfer.items&&e.dataTransfer.items.length>0&&(e.dataTransfer.dropEffect=this._dropEffect,this.renderRoot.querySelector("#file-upload-border").classList.add("visible"))},this.handleonDragLeave=e=>{e.preventDefault(),e.stopPropagation(),this._dragCounter--,this._dragCounter===0&&this.renderRoot.querySelector("#file-upload-border").classList.remove("visible")},this.handleonDrop=e=>{var t;e.preventDefault(),e.stopPropagation();const i=()=>{e.dataTransfer.clearData()};this.renderRoot.querySelector("#file-upload-border").classList.remove("visible"),!((t=e.dataTransfer)===null||t===void 0)&&t.items&&this.readUploadedFiles(e.dataTransfer.items,i),this._dragCounter=0},this.filesToUpload=[],this.addEventListener("__uploadfailed",this.focusOnUpload),this.addEventListener("__uploadsuccess",this.focusOnUpload)}focusOnUpload(){const e=this.renderRoot.querySelector('mgt-file[part="upload"]');e&&(e.setAttribute("tabindex","0"),e.classList.add("upload"),e.focus())}render(){if(this.parentElement!==null){const e=this.parentElement;e.addEventListener("dragenter",this.handleonDragEnter),e.addEventListener("dragleave",this.handleonDragLeave),e.addEventListener("dragover",this.handleonDragOver),e.addEventListener("drop",this.handleonDrop)}return v`
        <div id="file-upload-dialog" class="file-upload-dialog">
          <!-- Modal content -->
          <fluent-dialog modal="true" class="file-upload-dialog-content">
            <span
              class="file-upload-dialog-close"
              id="file-upload-dialog-close">
                ${S(w.Cancel)}
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
              <span slot="start">${S(w.Upload)}</span>
              <span class="upload-text">${this.strings.buttonUploadFile}</span>
          </fluent-button>
        </div>
        <div class="file-upload-template">
          ${this.renderFolderTemplate(this.filesToUpload)}
        </div>
       `}renderFolderTemplate(e){const t=[];if(e.length>0){const i=e.map(o=>t.indexOf(o.fullPath.substring(0,o.fullPath.lastIndexOf("/")))===-1?o.fullPath.endsWith("/")?v`${this.renderFileTemplate(o,"")}`:(t.push(o.fullPath.substring(0,o.fullPath.lastIndexOf("/"))),T`
            <div class='file-upload-table'>
              <div class='file-upload-cell'>
                <mgt-file
                  .fileDetails=${{name:o.fullPath.substring(1,o.fullPath.lastIndexOf("/")),folder:"Folder"}}
                  view="oneline"
                  class="mgt-file-item">
                </mgt-file>
              </div>
            </div>
            ${this.renderFileTemplate(o,"file-upload-folder-tab")}`):v`${this.renderFileTemplate(o,"file-upload-folder-tab")}`);return v`${i}`}return v``}renderFileTemplate(e,t){const i=z({"file-upload-table":!0,upload:e.completed}),o=t+(e.fieldUploadResponse==="lastModifiedDateTime"?" file-upload-dialog-success":""),l=e.fieldUploadResponse==="description",s=z({description:l}),r=e.completed?v``:this.renderFileUploadTemplate(e),n=l?p.failUploadFile:p.successUploadFile;return T`
        <div class="${i}">
          <div class="${o}">
            <div class='file-upload-cell'>
              <div class="${s}">
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
        </div>`}renderFileUploadTemplate(e){const t=z({"file-upload-table":!0,upload:e.completed});return v`
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
                ${S(w.Cancel)}
              </span>
            </div>
          </div>
        </div>
      </div>
    </div>
    `}deleteFileUploadSession(e){return F(this,void 0,void 0,function*(){try{e.uploadUrl!==void 0?(yield we(this.fileUploadList.graph,e.uploadUrl),e.uploadUrl=void 0,e.completed=!0,this.setUploadFail(e,p.cancelUploadFile)):(e.uploadUrl=void 0,e.completed=!0,this.setUploadFail(e,p.cancelUploadFile))}catch{e.uploadUrl=void 0,e.completed=!0,this.setUploadFail(e,p.cancelUploadFile)}})}readUploadedFiles(e,t){return F(this,void 0,void 0,function*(){const i=yield this.getFilesFromUploadArea(e);yield this.getSelectedFiles(i),t()})}getSelectedFiles(e){return F(this,void 0,void 0,function*(){let t=[];const i=[];this._applyAll=!1,this._applyAllConflictBehavior=null,this._maximumFileSize=!1,this._excludedFileType=!1,this.filesToUpload.forEach(l=>{l.completed?i.push(l):t.push(l)});for(const l of e){const s=l.fullPath===""?"/"+l.name:l.fullPath;if(t.filter(r=>r.fullPath===s).length===0){let r=!0;if(this.fileUploadList.maxFileSize!==void 0&&r&&l.size>this.fileUploadList.maxFileSize*1024&&(r=!1,this._maximumFileSize===!1)){const n=yield this.getFileUploadStatus(l,s,"MaxFileSize",this.fileUploadList);n!==null&&n[0]===1&&(this._maximumFileSize=!0)}if(this.fileUploadList.excludedFileExtensions!==void 0&&this.fileUploadList.excludedFileExtensions.length>0&&r&&this.fileUploadList.excludedFileExtensions.filter(n=>l.name.toLowerCase().indexOf(n.toLowerCase())>-1).length>0&&(r=!1,this._excludedFileType===!1)){const n=yield this.getFileUploadStatus(l,s,"ExcludedFileType",this.fileUploadList);n!==null&&n[0]===1&&(this._excludedFileType=!0)}if(r){const n=yield this.getFileUploadStatus(l,s,"Upload",this.fileUploadList);let u=!1;n!==null&&(n[0]===-1?u=!0:(this._applyAll=!!n[0],this._applyAllConflictBehavior=n[1]?1:0)),t.push({file:l,driveItem:{name:l.name},fullPath:s,conflictBehavior:St(n),iconStatus:null,percent:1,view:"image",completed:u,maxSize:this._maxChunkSize,minSize:0})}}}t=t.sort((l,s)=>l.fullPath.substring(0,l.fullPath.lastIndexOf("/")).localeCompare(s.fullPath.substring(0,s.fullPath.lastIndexOf("/")))),t.forEach(l=>{if(i.filter(s=>s.fullPath===l.fullPath).length!==0){const s=i.findIndex(r=>r.fullPath===l.fullPath);i.splice(s,1)}}),t.push(...i),this.filesToUpload=t;const o=this.filesToUpload.map(l=>this.sendFileItemGraph(l));yield Promise.all(o)})}getFileUploadStatus(e,t,i,o){const l=Object.create(null,{requestStateUpdate:{get:()=>super.requestStateUpdate}});return F(this,void 0,void 0,function*(){const s=this.renderRoot.querySelector("#file-upload-dialog");switch(i){case"Upload":return(yield _e(this.fileUploadList.graph,`${this.getGrapQuery(t)}?$select=id`))!==null?this._applyAll===!0?[this._applyAll,this._applyAllConflictBehavior]:(s.classList.add("visible"),this._dialogTitle=p.fileReplaceTitle,this._dialogContent=p.fileReplace.replace("{FileName}",e.name),this._dialogCheckBox=p.checkApplyAll,this._dialogPrimaryButton=p.buttonReplace,this._dialogSecondaryButton=p.buttonKeep,yield l.requestStateUpdate.call(this,!0),new Promise(n=>{const u=this.renderRoot.querySelector(".file-upload-dialog-close"),d=this.renderRoot.querySelector(".file-upload-dialog-ok"),c=this.renderRoot.querySelector(".file-upload-dialog-cancel"),y=this.renderRoot.querySelector("#file-upload-dialog-check");y.checked=!1,y.classList.remove("hide");const b=()=>{s.classList.remove("visible"),n([y.checked?1:0,"replace"])},L=()=>{s.classList.remove("visible"),n([y.checked?1:0,"rename"])},M=()=>{s.classList.remove("visible"),n([-1])};d.removeEventListener("click",b),c.removeEventListener("click",L),u.removeEventListener("click",M),d.addEventListener("click",b),c.addEventListener("click",L),u.addEventListener("click",M)})):null;case"ExcludedFileType":return s.classList.add("visible"),this._dialogTitle=p.fileTypeTitle,this._dialogContent=p.fileType.replace("{FileName}",e.name)+" ("+o.excludedFileExtensions.join(",")+")",this._dialogCheckBox=p.checkAgain,this._dialogPrimaryButton=p.buttonOk,this._dialogSecondaryButton=p.buttonCancel,yield l.requestStateUpdate.call(this,!0),new Promise(r=>{const n=this.renderRoot.querySelector(".file-upload-dialog-ok"),u=this.renderRoot.querySelector(".file-upload-dialog-cancel"),d=this.renderRoot.querySelector(".file-upload-dialog-close"),c=this.renderRoot.querySelector("#file-upload-dialog-check");c.checked=!1,c.classList.remove("hide");const y=()=>{s.classList.remove("visible"),r([c.checked?1:0])},b=()=>{s.classList.remove("visible"),r([0])};n.removeEventListener("click",y),u.removeEventListener("click",b),d.removeEventListener("click",b),n.addEventListener("click",y),u.addEventListener("click",b),d.addEventListener("click",b)});case"MaxFileSize":return s.classList.add("visible"),this._dialogTitle=p.maximumFileSizeTitle,this._dialogContent=p.maximumFileSize.replace("{FileSize}",j(o.maxFileSize*1024)).replace("{FileName}",e.name)+j(e.size)+".",this._dialogCheckBox=p.checkAgain,this._dialogPrimaryButton=p.buttonOk,this._dialogSecondaryButton=p.buttonCancel,yield l.requestStateUpdate.call(this,!0),new Promise(r=>{const n=this.renderRoot.querySelector(".file-upload-dialog-ok"),u=this.renderRoot.querySelector(".file-upload-dialog-cancel"),d=this.renderRoot.querySelector(".file-upload-dialog-close"),c=this.renderRoot.querySelector("#file-upload-dialog-check");c.checked=!1,c.classList.remove("hide");const y=()=>{s.classList.remove("visible"),r([c.checked?1:0])},b=()=>{s.classList.remove("visible"),r([0])};n.removeEventListener("click",y),u.removeEventListener("click",b),d.removeEventListener("click",b),n.addEventListener("click",y),u.addEventListener("click",b),d.addEventListener("click",b)})}})}getGrapQuery(e){let t="";return this.fileUploadList.itemPath&&this.fileUploadList.itemPath.length>0&&(t=this.fileUploadList.itemPath.startsWith("/")?this.fileUploadList.itemPath:"/"+this.fileUploadList.itemPath),this.fileUploadList.userId&&this.fileUploadList.itemId?`/users/${this.fileUploadList.userId}/drive/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.userId&&this.fileUploadList.itemPath?`/users/${this.fileUploadList.userId}/drive/root:${t}${e}`:this.fileUploadList.groupId&&this.fileUploadList.itemId?`/groups/${this.fileUploadList.groupId}/drive/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.groupId&&this.fileUploadList.itemPath?`/groups/${this.fileUploadList.groupId}/drive/root:${t}${e}`:this.fileUploadList.driveId&&this.fileUploadList.itemId?`/drives/${this.fileUploadList.driveId}/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.driveId&&this.fileUploadList.itemPath?`/drives/${this.fileUploadList.driveId}/root:${t}${e}`:this.fileUploadList.siteId&&this.fileUploadList.itemId?`/sites/${this.fileUploadList.siteId}/drive/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.siteId&&this.fileUploadList.itemPath?`/sites/${this.fileUploadList.siteId}/drive/root:${t}${e}`:this.fileUploadList.itemId?`/me/drive/items/${this.fileUploadList.itemId}:${e}`:this.fileUploadList.itemPath?`/me/drive/root:${t}${e}`:`/me/drive/root:${e}`}sendFileItemGraph(e){return F(this,void 0,void 0,function*(){const t=this.fileUploadList.graph;let i="";if(e.file.size<this._maxChunkSize)try{e.completed||((e.conflictBehavior===null||e.conflictBehavior==="replace")&&(i=`${this.getGrapQuery(e.fullPath)}:/content`),e.conflictBehavior==="rename"&&(i=`${this.getGrapQuery(e.fullPath)}:/content?@microsoft.graph.conflictBehavior=rename`),e.driveItem=yield ke(t,i,e.file),e.driveItem!==null?this.setUploadSuccess(e):(e.driveItem={name:e.file.name},this.setUploadFail(e,p.failUploadFile)))}catch{this.setUploadFail(e,p.failUploadFile)}else if(!e.completed&&e.uploadUrl===void 0){const o=yield Ue(t,`${this.getGrapQuery(e.fullPath)}:/createUploadSession`,e.conflictBehavior);try{if(o!==null){e.uploadUrl=o.uploadUrl;const l=yield this.sendSessionUrlGraph(t,e);l!==null?(e.driveItem=l,this.setUploadSuccess(e)):this.setUploadFail(e,p.failUploadFile)}else this.setUploadFail(e,p.failUploadFile)}catch{}}})}sendSessionUrlGraph(e,t){const i=Object.create(null,{requestStateUpdate:{get:()=>super.requestStateUpdate}});return F(this,void 0,void 0,function*(){for(;t.file.size>t.minSize;){t.mimeStreamString===void 0&&(t.mimeStreamString=yield this.readFileContent(t.file));const o=new Blob([t.mimeStreamString.slice(t.minSize,t.maxSize)]);if(t.percent=Math.round(t.maxSize/t.file.size*100),yield i.requestStateUpdate.call(this,!0),t.uploadUrl!==void 0){const l=yield $e(e,t.uploadUrl,`${t.maxSize-t.minSize}`,`bytes ${t.minSize}-${t.maxSize-1}/${t.file.size}`,o);if(l===null)return null;if(Ce(l))t.minSize=parseInt(l.nextExpectedRanges[0].split("-")[0],10),t.maxSize=t.minSize+this._maxChunkSize,t.maxSize>t.file.size&&(t.maxSize=t.file.size);else if(l.id!==void 0)return l}else return null}})}setUploadSuccess(e){e.percent=100,super.requestStateUpdate(!0),setTimeout(()=>{e.iconStatus=S(w.Success),e.view="twolines",e.fieldUploadResponse="lastModifiedDateTime",e.completed=!0,super.requestStateUpdate(!0),ie(),this.fireCustomEvent("__uploadsuccess")},500)}setUploadFail(e,t){setTimeout(()=>{e.iconStatus=S(w.Fail),e.view="twolines",e.driveItem.description=t,e.fieldUploadResponse="description",e.completed=!0,super.requestStateUpdate(!0),this.fireCustomEvent("__uploadfailed")},500)}readFileContent(e){return new Promise((t,i)=>{const o=new FileReader;o.onloadend=()=>{t(o.result)},o.onerror=l=>{i(l)},o.readAsArrayBuffer(e)})}getFilesFromUploadArea(e){return F(this,void 0,void 0,function*(){const t=[];let i;const o=[];for(const l of e)if(xt(l))if(yt(l))if(i=l.getAsEntry(),P(i))t.push(i);else{const s=l.getAsFile();s&&(this.writeFilePath(s,""),o.push(s))}else if(l.webkitGetAsEntry)if(i=l.webkitGetAsEntry(),P(i))t.push(i);else{const s=l.getAsFile();s&&(this.writeFilePath(s,""),o.push(s))}else{const s=l.getAsFile();s&&(this.writeFilePath(s,""),o.push(s))}else this.writeFilePath(l,""),o.push(l);if(t.length>0){const l=yield this.getFolderFiles(t);o.push(...l)}return o})}getFolderFiles(e){return new Promise(t=>{let i=0;const o=[];e.forEach(r=>{l(r,"")});const l=(r,n)=>{P(r)?s(r.createReader()):bt(r)&&(i++,r.file(u=>{i--,this.writeFilePath(u,n),o.push(u),i===0&&t(o)}))},s=r=>{i++,r.readEntries(n=>{i--;for(const u of n)l(u,u.fullPath);i===0&&t(o)})}})}writeFilePath(e,t){e.fullPath=t}}ae([m({type:Object}),re("design:type",Array)],O.prototype,"filesToUpload",void 0);ae([m({type:Object}),re("design:type",Object)],O.prototype,"fileUploadList",void 0);const wt=[J`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size, 14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host{font-size:var(--default-font-size)}:host .title{font-size:14px;font-weight:600;padding:20px 0 12px;line-height:19px}:host .file-list-wrapper{background-color:var(--file-list-background-color,var(--neutral-layer-1));border:var(--file-list-border,none);position:relative;display:flex;flex-direction:column;border-radius:8px}:host .file-list-wrapper .title{font-size:14px;font-weight:600;margin:0 20px -15px}:host .file-list-wrapper .file-list{display:flex;padding:var(--file-list-padding,0);margin:var(--file-list-margin,0);flex-direction:column;list-style:none}:host .file-list-wrapper .file-list .file-list-children-show{display:block;margin-inline-start:calc(16px * 2)}:host .file-list-wrapper .file-list .file-list-children-show ul{list-style:none}:host .file-list-wrapper .file-list .file-list-children-show ul li{padding:8px 0}:host .file-list-wrapper .file-list .file-list-children-hide{display:none}:host .file-list-wrapper .file-list .file-item{cursor:pointer;border-radius:var(--file-border-radius)}:host .file-list-wrapper .file-list .file-item:focus,:host .file-list-wrapper .file-list .file-item:focus-within{--file-background-color:var(--file-background-color-focus, var(--neutral-layer-2))}:host .file-list-wrapper .file-list .file-item.selected{--file-background-color:var(--file-background-color-active, var(--neutral-layer-3))}:host .file-list-wrapper .file-list .file-item .mgt-file-item{--file-padding:10px 20px 10px 20px;--file-padding-inline-start:24px;--file-border-radius:2px;--file-background-color:var(--file-item-background-color, var(--neutral-layer-1))}:host .file-list-wrapper .file-list .file-item .file-list-children{background-color:beige}:host .file-list-wrapper .progress-ring{margin:4px auto;width:var(--progress-ring-size,24px);height:var(--progress-ring-size,24px)}:host .file-list-wrapper .show-more{text-align:center;font-size:var(--show-more-button-font-size, 12px);padding:var(--show-more-button-padding,0);border-radius:0 0 var(--show-more-button-border-bottom-right-radius,var(--file-list-border-radius,8px)) var(--show-more-button-border-bottom-left-radius,var(--file-list-border-radius,8px));background-color:var(--show-more-button-background-color,var(--neutral-stroke-divider-rest))}:host .file-list-wrapper .show-more:hover{background-color:var(--show-more-button-background-color-hover,var(--neutral-fill-input-alt-active))}:host .file-list-wrapper .show-more-text{font-size:var(--show-more-button-font-size, 12px)}:host .file-list-wrapper .shared_insight_file{display:flex;align-items:center;padding:11px 20px}:host .file-list-wrapper .shared_insight_file:hover{background-color:var(--file-item-background-color,var(--neutral-layer-1));cursor:pointer}:host .file-list-wrapper .shared_insight_file:last-child{margin-bottom:unset}:host .file-list-wrapper .shared_insight_file .shared_insight_file__icon{width:28px;min-width:28px;height:28px;margin-inline-end:12px;display:flex;align-items:center;justify-content:center}:host .file-list-wrapper .shared_insight_file .shared_insight_file__icon img{height:28px;width:28px}:host .file-list-wrapper .shared_insight_file .shared_insight_file__details{min-width:0;margin-inline-end:40px}:host .file-list-wrapper .shared_insight_file .shared_insight_file__details .shared_insight_file__name{font-size:var(--file-line1-font-size, var(--size-font-size, 12px));color:var(--file-line1-color,var(--neutral-foreground-rest))}:host .file-list-wrapper .shared_insight_file .shared_insight_file__details .shared_insight_file__last-modified{font-size:var(--file-line3-font-size, var(--size-font-size, 12px));color:var(--file-line3-color,var(--secondary-text-color))}:host .file-list-wrapper .shared_insight_file .shared_insight_file__details>div{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
`],_t={showMoreSubtitle:"Show more items",filesSectionTitle:"Files",sharedTextSubtitle:"Shared"};var f=function(a,e,t,i){var o=arguments.length,l=o<3?e:i===null?i=Object.getOwnPropertyDescriptor(e,t):i,s;if(typeof Reflect=="object"&&typeof Reflect.decorate=="function")l=Reflect.decorate(a,e,t,i);else for(var r=a.length-1;r>=0;r--)(s=a[r])&&(l=(o<3?s(l):o>3?s(e,t,l):s(e,t))||l);return o>3&&l&&Object.defineProperty(e,t,l),l},g=function(a,e){if(typeof Reflect=="object"&&typeof Reflect.metadata=="function")return Reflect.metadata(a,e)},Z=function(a,e,t,i){function o(l){return l instanceof t?l:new t(function(s){s(l)})}return new(t||(t=Promise))(function(l,s){function r(d){try{u(i.next(d))}catch(c){s(c)}}function n(d){try{u(i.throw(d))}catch(c){s(c)}}function u(d){d.done?l(d.value):o(d.value).then(r,n)}u((i=i.apply(a,e||[])).next())})};const It=()=>{ee(Ve),le(),Ft(),te("file-list",h)},kt=a=>"lastShared"in a;class h extends Le{static get styles(){return wt}get strings(){return _t}get displayName(){return this.strings.filesSectionTitle}get cardTitle(){return this.strings.filesSectionTitle}renderIcon(){return S(w.Files)}static get requiredScopes(){return[...new Set([...He.requiredScopes])]}constructor(e){super(),this._isCompact=!1,this.fileQueries=null,this.files=null,this.itemView="threelines",this.fileExtensions=[],this.pageSize=10,this.disableOpenOnClick=!1,this.hideMoreFilesButton=!1,this.enableFileUpload=!1,this.maxUploadFile=10,this.excludedFileExtensions=[],this._preloadedFiles=[],this._focusedItemIndex=-1,this.renderLoading=()=>this.files?this.renderContent():this.renderTemplate("loading",null)||v``,this.renderContent=()=>!this.files||this.files.length===0?this.renderNoData():(this._personCardFiles&&(this.files=this._personCardFiles),this._isCompact?this.renderCompactView():this.renderFullView()),this.onFocusFirstItem=()=>this._focusedItemIndex=0,this.onFileListKeyDown=t=>{const i=t.target;let o;i.classList?o=this.renderRoot.querySelector(".file-list"):o=this.renderRoot.querySelector(".file-list-children");let l;if(o!=null&&o.children.length){if((t.code==="ArrowUp"||t.code==="ArrowDown")&&(t.code==="ArrowUp"&&(this._focusedItemIndex===-1&&(this._focusedItemIndex=o.children.length),this._focusedItemIndex=(this._focusedItemIndex-1+o.children.length)%o.children.length),t.code==="ArrowDown"&&(this._focusedItemIndex=(this._focusedItemIndex+1)%o.children.length),l=o.children[this._focusedItemIndex],this.updateItemBackgroundColor(o,l,"focused")),t.code==="Enter"||t.code==="Space"){l=o.children[this._focusedItemIndex];const s=l.children[0];t.preventDefault(),this.fireCustomEvent("itemClick",s.fileDetails),this.handleFileClick(s.fileDetails),this.updateItemBackgroundColor(o,l,"selected")}t.code==="Tab"&&(l=o.children[this._focusedItemIndex])}},this.handleSharedInsightClick=(t,i)=>{var o;!((o=t.resourceReference)===null||o===void 0)&&o.webUrl&&!this.disableOpenOnClick&&(i.preventDefault(),window.open(t.resourceReference.webUrl,"_blank","noreferrer"))},this.handleFileClick=t=>{var i;if(((i=t==null?void 0:t.folder)===null||i===void 0?void 0:i.childCount)>0&&(t==null?void 0:t.children)){this.showChildren(t.id);return}t!=null&&t.webUrl&&!this.disableOpenOnClick&&window.open(t.webUrl,"_blank","noreferrer")},this.showChildren=t=>{const i=this.renderRoot.querySelector(`#file-list-item-${t}`);this.renderChildren(t,i)},this.renderChildren=(t,i)=>{const o=K.isDisambiguated?`${K.prefix}-file-list`:"mgt-file-list",l=this.renderRoot.querySelector(`#file-list-children-${t}`);if(l)l.classList.contains("file-list-children-hide")?l.setAttribute("class","file-list-children-show"):l.setAttribute("class","file-list-children-hide");else{const s=document.createElement(o);s.setAttribute("item-id",t),s.setAttribute("id",`file-list-children-${t}`),s.setAttribute("class","file-list-children-show"),i.after(s)}},this._personCardFiles=e}clearState(){super.clearState(),this.files=null,this._personCardFiles=null}asCompactView(){return this._isCompact=!0,this}asFullView(){return this._isCompact=!1,this}args(){return[this.providerState,this.fileListQuery,this.fileQueries,this.siteId,this.driveId,this.groupId,this.itemId,this.itemPath,this.userId,this.insightType,this.fileExtensions,this.pageSize,this.maxFileSize]}renderCompactView(){const e=this.files.slice(0,3);return this.renderFiles(e)}renderFullView(){return this.renderTemplate("default",{files:this.files})||this.renderFiles(this.files)}renderNoData(){return this.renderTemplate("no-data",null)||(this.enableFileUpload===!0&&D.globalProvider!==void 0?v`
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
            id="file-list-item-${this.files[0].id}"
            tabindex="0"
            class="file-item"
            @keydown="${this.onFileListKeyDown}"
            @focus="${this.onFocusFirstItem}"
            @click=${t=>this.handleItemSelect(e[0],t)}>
            ${this.renderFile(e[0])}
          </li>
          ${Ke(e.slice(1),t=>t.id,t=>v`
              <li
                id="file-list-item-${t.id}"
                class="file-item"
                @keydown="${this.onFileListKeyDown}"
                @click=${i=>this.handleItemSelect(t,i)}>
                ${this.renderFile(t)}
              </li>
            `)}
        </ul>
        ${!this.hideMoreFilesButton&&this.pageIterator&&(this.pageIterator.hasNext||this._preloadedFiles.length)&&!this._isCompact?this.renderMoreFileButton():null}
      </div>
    `}renderFile(e){const t=this.itemView;return kt(e)?this.renderSharedInsightFile(e):this.renderTemplate("file",{file:e},e.id)||T`
        <mgt-file class="mgt-file-item" .fileDetails=${e} .view=${t}></mgt-file>
      `}renderSharedInsightFile(e){const t=e.lastShared?v`
          <div class="shared_insight_file__last-modified">
            ${this.strings.sharedTextSubtitle} ${Te(new Date(e.lastShared.sharedDateTime))}
          </div>
        `:null;return v`
      <div class="shared_insight_file" @click=${i=>this.handleSharedInsightClick(e,i)} tabindex="0">
        <div class="shared_insight_file__icon">
          <img alt="${e.resourceVisualization.title}" src=${Xe(e.resourceVisualization.type,48,"svg")} />
        </div>
        <div class="shared_insight_file__details">
          <div class="shared_insight_file__name">${e.resourceVisualization.title}</div>
          ${t}
        </div>
      </div>
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
        </fluent-button>`}renderFileUpload(){const e={graph:D.globalProvider.graph.forComponent(this),driveId:this.driveId,excludedFileExtensions:this.excludedFileExtensions,groupId:this.groupId,itemId:this.itemId,itemPath:this.itemPath,userId:this.userId,siteId:this.siteId,maxFileSize:this.maxFileSize,maxUploadFile:this.maxUploadFile};return T`
        <mgt-file-upload .fileUploadList=${e} ></mgt-file-upload>
      `}loadState(){var e,t,i,o;return Z(this,void 0,void 0,function*(){const l=D.globalProvider;if(!l||l.state===H.Loading)return;if(l.state===H.SignedOut){this.files=null;return}const s=l.graph.forComponent(this);let r,n;const u=!this.driveId&&!this.siteId&&!this.groupId&&!this.userId;if((this.driveId&&!this.itemId&&!this.itemPath||this.groupId&&!this.itemId&&!this.itemPath||this.siteId&&!this.itemId&&!this.itemPath||this.userId&&!this.insightType&&!this.itemId&&!this.itemPath)&&(this.files=null),!this.files){this.fileListQuery?n=yield Ie(s,this.fileListQuery,this.pageSize):this.fileQueries?r=yield Ee(s,this.fileQueries):u?this.itemId?n=yield ze(s,this.itemId,this.pageSize):this.itemPath?n=yield De(s,this.itemPath,this.pageSize):this.insightType?r=yield Pe(s,this.insightType):n=yield Q(s,this.pageSize):this.driveId?this.itemId?n=yield G(s,this.driveId,this.itemId,this.pageSize):this.itemPath&&(n=yield Re(s,this.driveId,this.itemPath,this.pageSize)):this.groupId?this.itemId?n=yield Be(s,this.groupId,this.itemId,this.pageSize):this.itemPath&&(n=yield Ae(s,this.groupId,this.itemPath,this.pageSize)):this.siteId?this.itemId?n=yield Oe(s,this.siteId,this.itemId,this.pageSize):this.itemPath&&(n=yield Me(s,this.siteId,this.itemPath,this.pageSize)):this.userId&&(this.itemId?n=yield qe(s,this.userId,this.itemId,this.pageSize):this.itemPath?n=yield Ne(s,this.userId,this.itemPath,this.pageSize):this.insightType&&(r=yield je(s,this.userId,this.insightType))),n&&(this.pageIterator=n,this._preloadedFiles=[...this.pageIterator.value],this._preloadedFiles.length>=this.pageSize?r=this._preloadedFiles.splice(0,this.pageSize):r=this._preloadedFiles.splice(0,this._preloadedFiles.length));let d;if(((e=this.fileExtensions)===null||e===void 0?void 0:e.length)>0){if(!((t=this.pageIterator)===null||t===void 0)&&t.value){const c=yield Q(s,1e3);for(;c.hasNext;)yield V(c);this.pageIterator=Qe.createFromValue(s,c.value),r=this.pageIterator.value,this._preloadedFiles=[]}d=r.filter(c=>{for(const y of this.fileExtensions)if(y===this.getFileExtension(c.name))return c}),this._preloadedFiles=[...d]}(d==null?void 0:d.length)>=0?(this.files=d,this.pageSize&&(r=this.files.splice(0,this.pageSize),this.files=r)):this.files=r}for(const d of this.files)if(((i=d==null?void 0:d.folder)===null||i===void 0?void 0:i.childCount)>0){const c=(o=d==null?void 0:d.parentReference)===null||o===void 0?void 0:o.driveId,y=d==null?void 0:d.id,b=yield G(s,c,y,5);if(b){const L=[...b.value];d.children=L}}})}handleItemSelect(e,t){if(this.handleFileClick(e),this.fireCustomEvent("itemClick",e),t){const i=this.renderRoot.querySelector(".file-list"),o=Array.from(i.children),l=t.target.closest("li"),s=o.indexOf(l);this._focusedItemIndex=s;const r=i.children[this._focusedItemIndex];this.updateItemBackgroundColor(i,r,"selected")}}renderNextPage(){return Z(this,void 0,void 0,function*(){if(this._preloadedFiles.length>0)this.files=[...this.files,...this._preloadedFiles.splice(0,Math.min(this.pageSize,this._preloadedFiles.length))];else if(this.pageIterator.hasNext){this._isLoadingMore=!0;const e=this.renderRoot.querySelector("#file-list-wrapper");e!=null&&e.animate&&e.animate([{height:"auto",transformOrigin:"top left"},{height:"auto",transformOrigin:"top left"}],{duration:1e3,easing:"ease-in-out",fill:"both"}),yield V(this.pageIterator),this._isLoadingMore=!1,this.files=this.pageIterator.value}this.requestUpdate()})}getFileExtension(e){return/(?:\.([^.]+))?$/.exec(e)[1]||""}updateItemBackgroundColor(e,t,i){for(const o of e.children)o.classList.remove(i),o.removeAttribute("tabindex");t&&(t.classList.add(i),t.scrollIntoView({behavior:"smooth",block:"nearest",inline:"start"}),t.setAttribute("tabindex","0"),t.focus());for(const o of e.children)o.classList.remove("selected")}reload(e=!1){e&&ie(),this._task.run()}}f([A(),g("design:type",Object)],h.prototype,"_isCompact",void 0);f([A(),g("design:type",Array)],h.prototype,"_personCardFiles",void 0);f([m({attribute:"file-list-query"}),g("design:type",String)],h.prototype,"fileListQuery",void 0);f([m({attribute:"file-queries",converter:(a,e)=>a?a.split(",").map(t=>t.trim()):null}),g("design:type",Array)],h.prototype,"fileQueries",void 0);f([m({type:Object}),g("design:type",Array)],h.prototype,"files",void 0);f([m({attribute:"site-id"}),g("design:type",String)],h.prototype,"siteId",void 0);f([m({attribute:"drive-id"}),g("design:type",String)],h.prototype,"driveId",void 0);f([m({attribute:"group-id"}),g("design:type",String)],h.prototype,"groupId",void 0);f([m({attribute:"item-id"}),g("design:type",String)],h.prototype,"itemId",void 0);f([m({attribute:"item-path"}),g("design:type",String)],h.prototype,"itemPath",void 0);f([m({attribute:"user-id"}),g("design:type",String)],h.prototype,"userId",void 0);f([m({attribute:"insight-type"}),g("design:type",String)],h.prototype,"insightType",void 0);f([m({attribute:"item-view",converter:a=>Ge(a,"threelines")}),g("design:type",String)],h.prototype,"itemView",void 0);f([m({attribute:"file-extensions",converter:(a,e)=>a.split(",").map(t=>t.trim())}),g("design:type",Array)],h.prototype,"fileExtensions",void 0);f([m({attribute:"page-size",type:Number}),g("design:type",Object)],h.prototype,"pageSize",void 0);f([m({attribute:"disable-open-on-click",type:Boolean}),g("design:type",Object)],h.prototype,"disableOpenOnClick",void 0);f([m({attribute:"hide-more-files-button",type:Boolean}),g("design:type",Object)],h.prototype,"hideMoreFilesButton",void 0);f([m({attribute:"max-file-size",type:Number}),g("design:type",Number)],h.prototype,"maxFileSize",void 0);f([m({attribute:"enable-file-upload",type:Boolean}),g("design:type",Object)],h.prototype,"enableFileUpload",void 0);f([m({attribute:"max-upload-file",type:Number}),g("design:type",Object)],h.prototype,"maxUploadFile",void 0);f([m({attribute:"excluded-file-extensions",converter:(a,e)=>a.split(",").map(t=>t.trim())}),g("design:type",Array)],h.prototype,"excludedFileExtensions",void 0);f([A(),g("design:type",Boolean)],h.prototype,"_isLoadingMore",void 0);export{h as M,It as r};
