import{p as fe,b as P,a as Z,r as S,j as d}from"./index-CUIyG1wb.js";import{P as me}from"./PageHeader-CaxvQlgU.js";import{bx as ge,by as ve,m as b,aF as z,bz as H,u as A,bA as j,_ as a,b as g,o as x,bB as B,bC as be,bD as q,F as V,r as I,v as ie,bE as se,k as xe,l as ye,j as we,f as $e,g as ke,c as Ce,d as Se,D as Ie,bF as Fe,y as k,z as N,G as K,I as M,bG as y,K as T,bH as _e,bI as Te,bJ as L,bK as Le,Q as re,R as h,bL as ae,bM as Ee,B as Ae,E as w,C as Be,bN as le,bO as de,bP as Ve,bQ as Ne,bR as Re,J as De,bS as Pe,bT as Me,bU as ze,bV as He,bW as je,bX as qe,M as Ke,X as Ue,bY as X,aK as We,aL as J,bZ as Ge,b_ as Oe,a0 as Ze,a3 as Xe,Z as Je,$ as f,a6 as Qe,a8 as Ye,b$ as et,c0 as Q,a4 as Y,a5 as ee,a2 as ce,au as tt,w as U,an as W,aq as ot,ar as nt,as as _}from"./App-B6S7UShP.js";import{c as te}from"./repeat-B9aMobyH.js";import{r as it}from"./mgt-spinner-DVktg0Ba.js";import{D as st}from"./index-CvJyBokB.js";import{f as rt}from"./index-D41GWJEh.js";import{r as at}from"./mgt-file-list-Bj8TlzDw.js";import{r as lt}from"./mgt-picker-C8WI9eOq.js";import"./mgt-file-C6LDyJdu.js";import"./mgt-get-BSVtnumE.js";class dt extends ve{constructor(t,e){super(t,e),this.observer=null,e.childList=!0}observe(){this.observer===null&&(this.observer=new MutationObserver(this.handleEvent.bind(this))),this.observer.observe(this.target,this.options)}disconnect(){this.observer.disconnect()}getNodes(){return"subtree"in this.options?Array.from(this.target.querySelectorAll(this.options.selector)):Array.from(this.target.childNodes)}}function ct(n){return typeof n=="string"&&(n={property:n}),new ge("fast-children",dt,n)}const ht=(n,t)=>b`
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
        ${z("control")}
    >
        ${H(n,t)}
        <span class="content" part="content">
            <slot ${A("defaultSlottedContent")}></slot>
        </span>
        ${j(n,t)}
    </a>
`;let v=class extends V{constructor(){super(...arguments),this.handleUnsupportedDelegatesFocus=()=>{var t;window.ShadowRoot&&!window.ShadowRoot.prototype.hasOwnProperty("delegatesFocus")&&(!((t=this.$fastController.definition.shadowOptions)===null||t===void 0)&&t.delegatesFocus)&&(this.focus=()=>{var e;(e=this.control)===null||e===void 0||e.focus()})}}connectedCallback(){super.connectedCallback(),this.handleUnsupportedDelegatesFocus()}};a([g],v.prototype,"download",void 0);a([g],v.prototype,"href",void 0);a([g],v.prototype,"hreflang",void 0);a([g],v.prototype,"ping",void 0);a([g],v.prototype,"referrerpolicy",void 0);a([g],v.prototype,"rel",void 0);a([g],v.prototype,"target",void 0);a([g],v.prototype,"type",void 0);a([x],v.prototype,"defaultSlottedContent",void 0);class R{}a([g({attribute:"aria-expanded"})],R.prototype,"ariaExpanded",void 0);B(R,be);B(v,q,R);const ut=(n,t)=>b`
    <div role="listitem" class="listitem" part="listitem">
        ${I(e=>e.href&&e.href.length>0,b`
                ${ht(n,t)}
            `)}
        ${I(e=>!e.href,b`
                ${H(n,t)}
                <slot></slot>
                ${j(n,t)}
            `)}
        ${I(e=>e.separator,b`
                <span class="separator" part="separator" aria-hidden="true">
                    <slot name="separator">${t.separator||""}</slot>
                </span>
            `)}
    </div>
`;class F extends v{constructor(){super(...arguments),this.separator=!0}}a([x],F.prototype,"separator",void 0);B(F,q,R);const pt=(n,t)=>b`
    <template role="navigation">
        <div role="list" class="list" part="list">
            <slot
                ${A({property:"slottedBreadcrumbItems",filter:ie()})}
            ></slot>
        </div>
    </template>
`;class he extends V{slottedBreadcrumbItemsChanged(){if(this.$fastController.isConnected){if(this.slottedBreadcrumbItems===void 0||this.slottedBreadcrumbItems.length===0)return;const t=this.slottedBreadcrumbItems[this.slottedBreadcrumbItems.length-1];this.slottedBreadcrumbItems.forEach(e=>{const o=e===t;this.setItemSeparator(e,o),this.setAriaCurrent(e,o)})}}setItemSeparator(t,e){t instanceof F&&(t.separator=!e)}findChildWithHref(t){var e,o;return t.childElementCount>0?t.querySelector("a[href]"):!((e=t.shadowRoot)===null||e===void 0)&&e.childElementCount?(o=t.shadowRoot)===null||o===void 0?void 0:o.querySelector("a[href]"):null}setAriaCurrent(t,e){const o=this.findChildWithHref(t);o===null&&t.hasAttribute("href")&&t instanceof F?e?t.setAttribute("aria-current","page"):t.removeAttribute("aria-current"):o!==null&&(e?o.setAttribute("aria-current","page"):o.removeAttribute("aria-current"))}}a([x],he.prototype,"slottedBreadcrumbItems",void 0);const ft=(n,t)=>b`
    <template
        role="treeitem"
        slot="${e=>e.isNestedItem()?"item":void 0}"
        tabindex="-1"
        class="${e=>e.expanded?"expanded":""} ${e=>e.selected?"selected":""} ${e=>e.nested?"nested":""}
            ${e=>e.disabled?"disabled":""}"
        aria-expanded="${e=>e.childItems&&e.childItemLength()>0?e.expanded:void 0}"
        aria-selected="${e=>e.selected}"
        aria-disabled="${e=>e.disabled}"
        @focusin="${(e,o)=>e.handleFocus(o.event)}"
        @focusout="${(e,o)=>e.handleBlur(o.event)}"
        ${ct({property:"childItems",filter:ie()})}
    >
        <div class="positioning-region" part="positioning-region">
            <div class="content-region" part="content-region">
                ${I(e=>e.childItems&&e.childItemLength()>0,b`
                        <div
                            aria-hidden="true"
                            class="expand-collapse-button"
                            part="expand-collapse-button"
                            @click="${(e,o)=>e.handleExpandCollapseButtonClick(o.event)}"
                            ${z("expandCollapseButton")}
                        >
                            <slot name="expand-collapse-glyph">
                                ${t.expandCollapseGlyph||""}
                            </slot>
                        </div>
                    `)}
                ${H(n,t)}
                <slot></slot>
                ${j(n,t)}
            </div>
        </div>
        ${I(e=>e.childItems&&e.childItemLength()>0&&(e.expanded||e.renderCollapsedChildren),b`
                <div role="group" class="items" part="items">
                    <slot name="item" ${A("items")}></slot>
                </div>
            `)}
    </template>
`;function $(n){return se(n)&&n.getAttribute("role")==="treeitem"}class c extends V{constructor(){super(...arguments),this.expanded=!1,this.focusable=!1,this.isNestedItem=()=>$(this.parentElement),this.handleExpandCollapseButtonClick=t=>{!this.disabled&&!t.defaultPrevented&&(this.expanded=!this.expanded)},this.handleFocus=t=>{this.setAttribute("tabindex","0")},this.handleBlur=t=>{this.setAttribute("tabindex","-1")}}expandedChanged(){this.$fastController.isConnected&&this.$emit("expanded-change",this)}selectedChanged(){this.$fastController.isConnected&&this.$emit("selected-change",this)}itemsChanged(t,e){this.$fastController.isConnected&&this.items.forEach(o=>{$(o)&&(o.nested=!0)})}static focusItem(t){t.focusable=!0,t.focus()}childItemLength(){const t=this.childItems.filter(e=>$(e));return t?t.length:0}}a([g({mode:"boolean"})],c.prototype,"expanded",void 0);a([g({mode:"boolean"})],c.prototype,"selected",void 0);a([g({mode:"boolean"})],c.prototype,"disabled",void 0);a([x],c.prototype,"focusable",void 0);a([x],c.prototype,"childItems",void 0);a([x],c.prototype,"items",void 0);a([x],c.prototype,"nested",void 0);a([x],c.prototype,"renderCollapsedChildren",void 0);B(c,q);const mt=(n,t)=>b`
    <template
        role="tree"
        ${z("treeView")}
        @keydown="${(e,o)=>e.handleKeyDown(o.event)}"
        @focusin="${(e,o)=>e.handleFocus(o.event)}"
        @focusout="${(e,o)=>e.handleBlur(o.event)}"
        @click="${(e,o)=>e.handleClick(o.event)}"
        @selected-change="${(e,o)=>e.handleSelectedChange(o.event)}"
    >
        <slot ${A("slottedTreeItems")}></slot>
    </template>
`;class D extends V{constructor(){super(...arguments),this.currentFocused=null,this.handleFocus=t=>{if(!(this.slottedTreeItems.length<1)){if(t.target===this){this.currentFocused===null&&(this.currentFocused=this.getValidFocusableItem()),this.currentFocused!==null&&c.focusItem(this.currentFocused);return}this.contains(t.target)&&(this.setAttribute("tabindex","-1"),this.currentFocused=t.target)}},this.handleBlur=t=>{t.target instanceof HTMLElement&&(t.relatedTarget===null||!this.contains(t.relatedTarget))&&this.setAttribute("tabindex","0")},this.handleKeyDown=t=>{if(t.defaultPrevented)return;if(this.slottedTreeItems.length<1)return!0;const e=this.getVisibleNodes();switch(t.key){case Se:e.length&&c.focusItem(e[0]);return;case Ce:e.length&&c.focusItem(e[e.length-1]);return;case ke:if(t.target&&this.isFocusableElement(t.target)){const o=t.target;o instanceof c&&o.childItemLength()>0&&o.expanded?o.expanded=!1:o instanceof c&&o.parentElement instanceof c&&c.focusItem(o.parentElement)}return!1;case $e:if(t.target&&this.isFocusableElement(t.target)){const o=t.target;o instanceof c&&o.childItemLength()>0&&!o.expanded?o.expanded=!0:o instanceof c&&o.childItemLength()>0&&this.focusNextNode(1,t.target)}return;case we:t.target&&this.isFocusableElement(t.target)&&this.focusNextNode(1,t.target);return;case ye:t.target&&this.isFocusableElement(t.target)&&this.focusNextNode(-1,t.target);return;case xe:this.handleClick(t);return}return!0},this.handleSelectedChange=t=>{if(t.defaultPrevented)return;if(!(t.target instanceof Element)||!$(t.target))return!0;const e=t.target;e.selected?(this.currentSelected&&this.currentSelected!==e&&(this.currentSelected.selected=!1),this.currentSelected=e):!e.selected&&this.currentSelected===e&&(this.currentSelected=null)},this.setItems=()=>{const t=this.treeView.querySelector("[aria-selected='true']");this.currentSelected=t,(this.currentFocused===null||!this.contains(this.currentFocused))&&(this.currentFocused=this.getValidFocusableItem()),this.nested=this.checkForNestedItems(),this.getVisibleNodes().forEach(o=>{$(o)&&(o.nested=this.nested)})},this.isFocusableElement=t=>$(t),this.isSelectedElement=t=>t.selected}slottedTreeItemsChanged(){this.$fastController.isConnected&&this.setItems()}connectedCallback(){super.connectedCallback(),this.setAttribute("tabindex","0"),Ie.queueUpdate(()=>{this.setItems()})}handleClick(t){if(t.defaultPrevented)return;if(!(t.target instanceof Element)||!$(t.target))return!0;const e=t.target;e.disabled||(e.selected=!e.selected)}focusNextNode(t,e){const o=this.getVisibleNodes();if(!o)return;const s=o[o.indexOf(e)+t];se(s)&&c.focusItem(s)}getValidFocusableItem(){const t=this.getVisibleNodes();let e=t.findIndex(this.isSelectedElement);return e===-1&&(e=t.findIndex(this.isFocusableElement)),e!==-1?t[e]:null}checkForNestedItems(){return this.slottedTreeItems.some(t=>$(t)&&t.querySelector("[role='treeitem']"))}getVisibleNodes(){return Fe(this,"[role='treeitem']")||[]}}a([g({attribute:"render-collapsed-nodes"})],D.prototype,"renderCollapsedNodes",void 0);a([x],D.prototype,"currentSelected",void 0);a([x],D.prototype,"slottedTreeItems",void 0);const gt=(n,t)=>k`
  ${N("inline-block")} :host {
    box-sizing: border-box;
    ${K};
  }

  .list {
    display: flex;
  }
`,vt=he.compose({baseName:"breadcrumb",template:pt,styles:gt}),bt=(n,t)=>k`
    ${N("inline-flex")} :host {
      background: transparent;
      color: ${M};
      fill: currentcolor;
      box-sizing: border-box;
      ${K};
      min-width: calc(${y} * 1px);
      border-radius: calc(${T} * 1px);
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
      color: ${_e};
    }

    .control:active {
      color: ${Te};
    }

    .control:${L} {
      ${Le}
    }

    :host(:not([href])),
    :host([aria-current]) .control {
      color: ${M};
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
  `.withBehaviors(re(k`
        :host(:not([href])),
        .start,
        .end,
        .separator {
          background: ${h.ButtonFace};
          color: ${h.ButtonText};
          fill: currentcolor;
        }
        .separator {
          fill: ${h.ButtonText};
        }
        :host([href]) {
          forced-color-adjust: none;
          background: ${h.ButtonFace};
          color: ${h.LinkText};
        }
        :host([href]) .control:hover {
          background: ${h.LinkText};
          color: ${h.HighlightText};
          fill: currentcolor;
        }
        .control:${L} {
          outline-color: ${h.LinkText};
        }
      `)),xt=F.compose({baseName:"breadcrumb-item",template:ut,styles:bt,shadowOptions:{delegatesFocus:!0},separator:`
    <svg width="12" height="12" xmlns="http://www.w3.org/2000/svg">
      <path d="M4.65 2.15a.5.5 0 000 .7L7.79 6 4.65 9.15a.5.5 0 10.7.7l3.5-3.5a.5.5 0 000-.7l-3.5-3.5a.5.5 0 00-.7 0z"/>
    </svg>
  `}),yt=(n,t)=>k`
  :host([hidden]) {
    display: none;
  }

  ${N("flex")} :host {
    flex-direction: column;
    align-items: stretch;
    min-width: fit-content;
    font-size: 0;
  }
`,wt=D.compose({baseName:"tree-view",template:mt,styles:yt}),$t=k`
  .expand-collapse-button svg {
    transform: rotate(0deg);
  }
  :host(.nested) .expand-collapse-button {
    left: var(--expand-collapse-button-nested-width, calc(${y} * -1px));
  }
  :host([selected])::after {
    left: calc(${ae} * 1px);
  }
  :host([expanded]) > .positioning-region .expand-collapse-button svg {
    transform: rotate(90deg);
  }
`,kt=k`
  .expand-collapse-button svg {
    transform: rotate(180deg);
  }
  :host(.nested) .expand-collapse-button {
    right: var(--expand-collapse-button-nested-width, calc(${y} * -1px));
  }
  :host([selected])::after {
    right: calc(${ae} * 1px);
  }
  :host([expanded]) > .positioning-region .expand-collapse-button svg {
    transform: rotate(90deg);
  }
`,oe=Ee`((${Ae} / 2) * ${w}) + ((${w} * ${Be}) / 2)`,Ct=le.create("tree-item-expand-collapse-hover").withDefault(n=>{const t=de.getValueFor(n);return t.evaluate(n,t.evaluate(n).hover).hover}),St=le.create("tree-item-expand-collapse-selected-hover").withDefault(n=>{const t=Ve.getValueFor(n);return de.getValueFor(n).evaluate(n,t.evaluate(n).rest).hover}),It=(n,t)=>k`
    ${N("block")} :host {
      contain: content;
      position: relative;
      outline: none;
      color: ${M};
      fill: currentcolor;
      cursor: pointer;
      font-family: ${Ne};
      --expand-collapse-button-size: calc(${y} * 1px);
      --tree-item-nested-width: 0;
    }

    .positioning-region {
      display: flex;
      position: relative;
      box-sizing: border-box;
      background: ${Re};
      border: calc(${De} * 1px) solid transparent;
      border-radius: calc(${T} * 1px);
      height: calc((${y} + 1) * 1px);
    }

    :host(:${L}) .positioning-region {
      ${Pe}
    }

    .positioning-region::before {
      content: '';
      display: block;
      width: var(--tree-item-nested-width);
      flex-shrink: 0;
    }

    :host(:not([disabled])) .positioning-region:hover {
      background: ${Me};
    }

    :host(:not([disabled])) .positioning-region:active {
      background: ${ze};
    }

    .content-region {
      display: inline-flex;
      align-items: center;
      white-space: nowrap;
      width: 100%;
      height: calc(${y} * 1px);
      margin-inline-start: calc(${w} * 2px + 8px);
      ${K}
    }

    .items {
      display: none;
      ${""} font-size: calc(1em + (${w} + 16) * 1px);
    }

    .expand-collapse-button {
      background: none;
      border: none;
      border-radius: calc(${T} * 1px);
      ${""} width: calc((${oe} + (${w} * 2)) * 1px);
      height: calc((${oe} + (${w} * 2)) * 1px);
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
      ${""} margin-inline-end: calc(${w} * 2px + 2px);
    }

    .end {
      ${""} margin-inline-start: calc(${w} * 2px + 2px);
    }

    :host(.expanded) > .items {
      display: block;
    }

    :host([disabled]) {
      opacity: ${He};
      cursor: ${je};
    }

    :host(.nested) .content-region {
      position: relative;
      margin-inline-start: var(--expand-collapse-button-size);
    }

    :host(.nested) .expand-collapse-button {
      position: absolute;
    }

    :host(.nested) .expand-collapse-button:hover {
      background: ${Ct};
    }

    :host(:not([disabled])[selected]) .positioning-region {
      background: ${qe};
    }

    :host(:not([disabled])[selected]) .expand-collapse-button:hover {
      background: ${St};
    }

    :host([selected])::after {
      content: '';
      display: block;
      position: absolute;
      top: calc((${y} / 4) * 1px);
      width: 3px;
      height: calc((${y} / 2) * 1px);
      ${""} background: ${Ke};
      border-radius: calc(${T} * 1px);
    }

    ::slotted(fluent-tree-item) {
      --tree-item-nested-width: 1em;
      --expand-collapse-button-nested-width: calc(${y} * -1px);
    }
  `.withBehaviors(new st($t,kt),re(k`
        :host {
          color: ${h.ButtonText};
        }
        .positioning-region {
          border-color: ${h.ButtonFace};
          background: ${h.ButtonFace};
        }
        :host(:not([disabled])) .positioning-region:hover,
        :host(:not([disabled])) .positioning-region:active,
        :host(:not([disabled])[selected]) .positioning-region {
          background: ${h.Highlight};
        }
        :host .positioning-region:hover .content-region,
        :host([selected]) .positioning-region .content-region {
          forced-color-adjust: none;
          color: ${h.HighlightText};
        }
        :host([disabled][selected]) .positioning-region .content-region {
          color: ${h.GrayText};
        }
        :host([selected])::after {
          background: ${h.HighlightText};
        }
        :host(:${L}) .positioning-region {
          forced-color-adjust: none;
          outline-color: ${h.ButtonFace};
        }
        :host([disabled]),
        :host([disabled]) .content-region,
        :host([disabled]) .positioning-region:hover .content-region {
          opacity: 1;
          color: ${h.GrayText};
        }
        :host(.nested) .expand-collapse-button:hover,
        :host(:not([disabled])[selected]) .expand-collapse-button:hover {
          background: ${h.ButtonFace};
          fill: ${h.ButtonText};
        }
      `)),Ft=c.compose({baseName:"tree-item",template:ft,styles:It,expandCollapseGlyph:`
    <svg width="12" height="12" xmlns="http://www.w3.org/2000/svg">
      <path d="M4.65 2.15a.5.5 0 000 .7L7.79 6 4.65 9.15a.5.5 0 10.7.7l3.5-3.5a.5.5 0 000-.7l-3.5-3.5a.5.5 0 00-.7 0z"/>
    </svg>
  `}),_t=[Ue`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size, 14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}:host .container{display:flex;flex-direction:column;position:relative}:host .dropdown{display:none;position:absolute;z-index:1000;top:34px}:host .dropdown.visible{display:flex}:host .dropdown .team-photo{width:24px;position:inherit;border-radius:50%;margin:0 6px}:host .dropdown .team-start-slot{width:max-content}:host .dropdown .team-parent-name{width:auto}:host .loading-text,:host .search-error-text{font-style:normal;font-weight:400;font-size:14px;line-height:20px}:host .message-parent{display:flex;flex-direction:row;gap:5px;padding:5px}:host .message-parent .loading-text{margin:auto}:host fluent-card{background:var(--channel-picker-dropdown-background-color,var(--fill-color));padding:2px;--card-height:auto;--width:var(--card-width)}:host fluent-text-field{width:100%}:host fluent-text-field::part(root){background:padding-box linear-gradient(var(--channel-picker-input-background-color,var(--neutral-fill-input-rest)),var(--channel-picker-input-background-color,var(--neutral-fill-input-rest))),border-box var(--channel-picker-input-border-color,var(--neutral-stroke-input-rest))}:host fluent-text-field::part(root):hover{background:padding-box linear-gradient(var(--channel-picker-input-background-color-hover,var(--neutral-fill-input-hover)),var(--channel-picker-input-background-color-hover,var(--neutral-fill-input-hover))),border-box var(--channel-picker-input-hover-border-color,var(--neutral-stroke-input-hover))}:host fluent-text-field::part(root):focus,:host fluent-text-field::part(root):focus-within{background:padding-box linear-gradient(var(--channel-picker-input-background-color-focus,var(--neutral-fill-input-focus)),var(--channel-picker-input-background-color-focus,var(--neutral-fill-input-focus))),border-box var(--channel-picker-input-focus-border-color,var(--neutral-stroke-input-focus))}:host fluent-text-field::part(control){word-spacing:inherit;text-indent:inherit;letter-spacing:inherit;padding:0}:host fluent-text-field::part(control)::placeholder{color:var(--channel-picker-input-placeholder-text-color,var(--input-placeholder-rest))}:host fluent-text-field::part(control):hover::placeholder{color:var(--channel-picker-input-placeholder-text-color-hover,var(--input-placeholder-hover))}:host fluent-text-field::part(control):focus-within::placeholder,:host fluent-text-field::part(control):focus::placeholder{color:var(--channel-picker-input-placeholder-text-color-focus,var(--input-placeholder-filled))}:host fluent-text-field .search-icon svg path{fill:var(--channel-picker-search-icon-color,currentColor)}:host fluent-text-field .search-icon::part(control){border-color:transparent}:host fluent-text-field .down-chevron{height:auto;min-width:auto}:host fluent-text-field .down-chevron svg path{fill:var(--channel-picker-down-chevron-color,currentColor)}:host fluent-text-field .up-chevron{height:auto;min-width:auto}:host fluent-text-field .up-chevron svg path{fill:var(--channel-picker-up-chevron-color,currentColor)}:host fluent-text-field .close-icon{height:auto;min-width:auto}:host fluent-text-field .close-icon svg path{fill:var(--channel-picker-close-icon-color,currentColor)}:host fluent-tree-view{min-width:100%;--tree-item-nested-width:2em}:host fluent-tree-item{width:100%;--tree-item-nested-width:2em}:host fluent-tree-item:focus-visible{outline:0}:host fluent-tree-item::part(expand-collapse-button){background:0 0}:host fluent-tree-item::part(content-region),:host fluent-tree-item::part(positioning-region){color:var(--channel-picker-dropdown-item-text-color,currentColor);background:var(--channel-picker-dropdown-background-color,transparent);border:calc(var(--stroke-width) * 2px) solid transparent;height:auto}:host fluent-tree-item::part(content-region):hover,:host fluent-tree-item::part(positioning-region):hover{background:var(--channel-picker-dropdown-item-background-color-hover,var(--neutral-fill-stealth-hover))}:host fluent-tree-item::part(content-region):hover::part(expand-collapse-button),:host fluent-tree-item::part(positioning-region):hover::part(expand-collapse-button){background:var(--channel-picker-dropdown-item-background-color-hover,var(--neutral-fill-stealth-hover))}:host fluent-tree-item fluent-tree-item::part(content-region){height:auto}:host fluent-breadcrumb-item{color:var(--channel-picker-dropdown-item-text-color-selected,var(--neutral-foreground-rest))}:host fluent-breadcrumb-item .team-parent-name{font-weight:700}:host fluent-breadcrumb-item .team-photo{width:19px;position:inherit;border-radius:50%}:host fluent-breadcrumb-item .arrow{margin-left:8px;margin-right:8px}:host fluent-breadcrumb-item .arrow svg{stroke:var(--channel-picker-arrow-fill,var(--neutral-foreground-rest))}[dir=rtl] :host{--direction:rtl}[dir=rtl] .dropdown{text-align:right}[dir=rtl] .arrow{transform:scaleX(-1);filter:fliph;filter:"FlipH";margin-right:0;margin-left:5px}[dir=rtl] .selected-team{padding-left:10px}[dir=rtl] .message-parent .loading-text{right:auto;left:10px;padding-right:8px;text-align:right}@media (forced-colors:active) and (prefers-color-scheme:dark){:host fluent-text-field svg{stroke:rgb(255,255,255)!important}}@media (forced-colors:active) and (prefers-color-scheme:light){:host fluent-text-field svg{stroke:rgb(0,0,0)!important}}
`];var G=function(n,t,e,o){function s(i){return i instanceof e?i:new e(function(r){r(i)})}return new(e||(e=Promise))(function(i,r){function l(p){try{m(o.next(p))}catch(C){r(C)}}function u(p){try{m(o.throw(p))}catch(C){r(C)}}function m(p){p.done?i(p.value):s(p.value).then(l,u)}m((o=o.apply(n,t||[])).next())})};const Tt=["Team.ReadBasic.All","TeamSettings.Read.All","TeamSettings.ReadWrite.All","User.Read.All","User.ReadWrite.All"],Lt=["Channel.ReadBasic.All","ChannelSettings.Read.All","ChannelSettings.ReadWrite.All"],Et=["Team.ReadBasic.All","TeamSettings.Read.All","TeamSettings.ReadWrite.All"],At=n=>G(void 0,void 0,void 0,function*(){const t=Tt,e=yield n.api("/me/joinedTeams").select(["displayName","id","isArchived"]).middlewareOptions(fe(t)).get();return e==null?void 0:e.value}),Bt=(n,t)=>G(void 0,void 0,void 0,function*(){let e,o={};if(X()){e=We.getCache(J.photos,J.photos.stores.teams);for(const s of t)try{const i=yield e.getValue(s);i&&Ge()>Date.now()-i.timeCached&&(o[s]=i)}catch{}if(Object.keys(o).length)return o}o={};for(const s of t)try{const i=yield Oe(n,`/teams/${s}`,Et);X()&&i&&(yield e.putValue(s,i)),o[s]=i}catch{}return o}),Vt=(n,t)=>G(void 0,void 0,void 0,function*(){var e,o;const s=n.createBatch();for(const l of t)s.get(l.id,`teams/${l.id}/channels`,Lt);const i=yield s.executeAll(),r=[];for(const l of t){const u=i.get(l.id);!((o=(e=u==null?void 0:u.content)===null||e===void 0?void 0:e.value)===null||o===void 0)&&o.length&&r.push({item:l,channels:u.content.value.map(m=>({item:m}))})}return r}),Nt={inputPlaceholderText:"Select a channel",noResultsFound:"We didn't find any matches.",loadingMessage:"Loading...",photoFor:"Teams photo for",teamsChannels:"Teams and channels results",closeButtonAriaLabel:"remove the selected channel",downChevronButtonAriaLabel:"Teams show results",upChevronButtonAriaLabel:"Teams hide results",searchButtonAriaLabel:"Search icon"};var ue=function(n,t,e,o){var s=arguments.length,i=s<3?t:o===null?o=Object.getOwnPropertyDescriptor(t,e):o,r;if(typeof Reflect=="object"&&typeof Reflect.decorate=="function")i=Reflect.decorate(n,t,e,o);else for(var l=n.length-1;l>=0;l--)(r=n[l])&&(i=(s<3?r(i):s>3?r(t,e,i):r(t,e))||i);return s>3&&i&&Object.defineProperty(t,e,i),i},pe=function(n,t){if(typeof Reflect=="object"&&typeof Reflect.metadata=="function")return Reflect.metadata(n,t)},ne=function(n,t,e,o){function s(i){return i instanceof e?i:new e(function(r){r(i)})}return new(e||(e=Promise))(function(i,r){function l(p){try{m(o.next(p))}catch(C){r(C)}}function u(p){try{m(o.throw(p))}catch(C){r(C)}}function m(p){p.done?i(p.value):s(p.value).then(l,u)}m((o=o.apply(n,t||[])).next())})};const Rt=()=>{Ze(vt,xt,tt,wt,Ft,rt),it(),Xe("teams-channel-picker",O)};class O extends Je{static get styles(){return _t}get strings(){return Nt}get selectedItem(){return this._selectedItemState?{channel:this._selectedItemState.item,team:this._selectedItemState.parent.item}:null}static get requiredScopes(){return["team.readbasic.all","channel.readbasic.all"]}set items(t){this._items!==t&&(this._items=t,this._treeViewState=t?this.generateTreeViewState(t):[],this.resetFocusState())}get items(){return this._items}get _inputWrapper(){return this.renderRoot.querySelector("fluent-text-field")}get _input(){const t=this._inputWrapper;return t==null?void 0:t.shadowRoot.querySelector("input")}constructor(){super(),this.teamsPhotos={},this._inputValue="",this._treeViewState=[],this._focusList=[],this.renderLoading=()=>this.renderContent(),this.renderContent=()=>{var t;const e={dropdown:!0,visible:this._isDropdownVisible};return this.renderTemplate("default",{teams:(t=this.items)!==null&&t!==void 0?t:[]})||f`
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
            <div slot="start" style="width: max-content;" @keydown=${this.handleStartSlotKeydown}>${this.renderSelected()}</div>
            <div slot="end" @keydown=${this.handleChevronKeydown}>${this.renderChevrons()}${this.renderCloseButton()}</div>
          </fluent-text-field>
          <fluent-card
            class=${Qe(e)}
          >
            ${this.renderDropdown()}
          </fluent-card>
        </div>`},this.handleInputClick=t=>{t.stopPropagation(),this.gainedFocus()},this.handleInputKeydown=t=>{const e=t.key;["ArrowDown","Enter"].includes(e)?this._isDropdownVisible?this.renderRoot.querySelector("fluent-tree-item").focus():this.gainedFocus():e==="Escape"?this.lostFocus():e==="Tab"&&this.blurPicker()},this.onClickCloseButton=()=>{this.removeSelectedChannel(null)},this.onKeydownCloseButton=t=>{t.key==="Enter"&&this.removeSelectedChannel(null)},this.renderError=()=>this.renderTemplate("error",null,"error")||f`
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
      `,this.renderLoadingIndicator=()=>this.renderTemplate("loading",null,"loading")||Ye`
        <div class="message-parent">
          <mgt-spinner></mgt-spinner>
          <div label="loading-text" aria-label="loading" class="loading-text">
            ${this.strings.loadingMessage}
          </div>
        </div>
      `,this.onKeydownTreeView=t=>{t.key==="Escape"&&this.lostFocus()},this.handleTeamTreeItemClick=t=>{t.preventDefault(),t.stopImmediatePropagation();const e=t.target;e&&(e.hasAttribute("expanded")?e.removeAttribute("expanded"):e.setAttribute("expanded","true"),e.hasAttribute("id")?e.setAttribute("selected","true"):e.removeAttribute("selected"))},this.handleInputChanged=t=>{const e=t.target;if(this._inputValue!==(e==null?void 0:e.value))this._inputValue=e==null?void 0:e.value;else return;t.key!=="Tab"&&t.key!=="Enter"&&t.key!=="Escape"&&this.gainedFocus(),this.debouncedSearch||(this.debouncedSearch=et(()=>{this.filterList()},400)),this.debouncedSearch()},this.loadTeamsIfNotLoaded=()=>{!this.items&&this._task.status!==Q.PENDING&&this._task.run()},this.handleWindowClick=t=>{t.target!==this&&this.lostFocus()},this.gainedFocus=()=>{const t=this._input;t&&t.focus(),this._isDropdownVisible=!0,this.toggleChevron(),this.resetFocusState(),this.requestUpdate()},this.lostFocus=()=>{this._inputValue="",this._input&&(this._input.value=this._inputValue,this._input.textContent="");const t=this._inputWrapper;t&&(t.value="",t.blur()),this._isDropdownVisible=!1,this.filterList(),this.toggleChevron(),this.requestUpdate(),this._selectedItemState!==void 0&&this.showCloseIcon()},this.handleFocus=()=>{this.gainedFocus()},this.handleUpChevronClick=t=>{t.stopPropagation(),this.lostFocus()},this.handleChevronKeydown=t=>{if(t.key==="Tab")this.blurPicker();else if(t.key==="Enter"){t.preventDefault();const e=t.target;e.classList.contains("down-chevron")?this.gainedFocus():e.classList.contains("up-chevron")&&this.lostFocus()}},this.handleStartSlotKeydown=t=>{t.key==="Tab"&&t.shiftKey&&this.blurPicker()},this.blurPicker=()=>{const t=this._inputWrapper,e=this._input;t==null||t.blur(),e==null||e.blur()},this._inputValue="",this._treeViewState=[],this._focusList=[],this._isDropdownVisible=!1}connectedCallback(){super.connectedCallback(),window.addEventListener("click",this.handleWindowClick),this.addEventListener("focus",this.loadTeamsIfNotLoaded),this.addEventListener("mouseover",this.loadTeamsIfNotLoaded),this.addEventListener("blur",this.lostFocus);const t=this.renderRoot.ownerDocument;t&&t.documentElement.setAttribute("dir",this.direction)}disconnectedCallback(){window.removeEventListener("click",this.handleWindowClick),this.removeEventListener("focus",this.loadTeamsIfNotLoaded),this.removeEventListener("mouseover",this.loadTeamsIfNotLoaded),this.removeEventListener("blur",this.lostFocus),super.disconnectedCallback()}args(){return[]}selectChannelById(t){return ne(this,void 0,void 0,function*(){const e=P.globalProvider;if(e&&e.state===Z.SignedIn){this.items||(yield this._task.run());for(const o of this._treeViewState)for(const s of o.channels)if(s.item.id===t)return o.isExpanded=!0,this.selectChannel(s),this.markSelectedChannelInDropdown(t),!0}return!1})}markSelectedChannelInDropdown(t){const e=this.renderRoot.querySelector(`[id='${t}']`);e&&(e.setAttribute("selected","true"),e.parentElement&&e.parentElement.setAttribute("expanded","true"))}renderSelected(){var t,e,o,s,i,r;if(!this._selectedItemState)return this.renderSearchIcon();let l;if(this._selectedItemState.parent.channels){const p=(t=this.teamsPhotos[this._selectedItemState.parent.item.id])===null||t===void 0?void 0:t.photo;l=f`<img
        class="team-photo"
        alt="${this._selectedItemState.parent.item.displayName}"
        role="img"
        src=${p} />`}const u=(s=(o=(e=this._selectedItemState)===null||e===void 0?void 0:e.parent)===null||o===void 0?void 0:o.item)===null||s===void 0?void 0:s.displayName.trim(),m=(r=(i=this._selectedItemState)===null||i===void 0?void 0:i.item)===null||r===void 0?void 0:r.displayName.trim();return f`
      <fluent-breadcrumb title=${this._selectedItemState.item.displayName}>
        <fluent-breadcrumb-item>
          <span slot="start">${l}</span>
          <span class="team-parent-name">${u}</span>
          <span slot="separator" class="arrow">${Y(ee.TeamSeparator,"#000000")}</span>
        </fluent-breadcrumb-item>
        <fluent-breadcrumb-item>${m}</fluent-breadcrumb-item>
      </fluent-breadcrumb>`}clearState(){this._inputValue="",this._treeViewState=[],this._focusList=[],this._isDropdownVisible=!1}renderSearchIcon(){return f`
      <fluent-button 
        appearance="outline" 
        class="search-icon" 
        aria-label=${this.strings.searchButtonAriaLabel} 
        @click=${this.handleStartSlotKeydown}>
        ${Y(ee.Search)}
      </fluent-button>
    `}renderCloseButton(){return f`
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
    `}showCloseIcon(){const t=this.renderRoot.querySelector(".down-chevron"),e=this.renderRoot.querySelector(".up-chevron"),o=this.renderRoot.querySelector(".close-icon");t&&(t.style.display="none"),e&&(e.style.display="none"),o&&(o.style.display=null)}renderDownChevron(){return f`
      <fluent-button
        aria-label=${this.strings.downChevronButtonAriaLabel}
        appearance="stealth"
        class="down-chevron"
        @click=${this.gainedFocus}
        @keydown=${this.handleChevronKeydown}>
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M2.21967 4.46967C2.51256 4.17678 2.98744 4.17678 3.28033 4.46967L6 7.18934L8.71967 4.46967C9.01256 4.17678 9.48744 4.17678 9.78033 4.46967C10.0732 4.76256 10.0732 5.23744 9.78033 5.53033L6.53033 8.78033C6.23744 9.07322 5.76256 9.07322 5.46967 8.78033L2.21967 5.53033C1.92678 5.23744 1.92678 4.76256 2.21967 4.46967Z" fill="#212121" />
        </svg>
      </fluent-button>`}renderUpChevron(){return f`
      <fluent-button
        aria-label=${this.strings.upChevronButtonAriaLabel}
        appearance="stealth"
        style="display:none"
        class="up-chevron"
        @click=${this.handleUpChevronClick}
        @keydown=${this.handleChevronKeydown}>
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M2.21967 7.53033C2.51256 7.82322 2.98744 7.82322 3.28033 7.53033L6 4.81066L8.71967 7.53033C9.01256 7.82322 9.48744 7.82322 9.78033 7.53033C10.0732 7.23744 10.0732 6.76256 9.78033 6.46967L6.53033 3.21967C6.23744 2.92678 5.76256 2.92678 5.46967 3.21967L2.21967 6.46967C1.92678 6.76256 1.92678 7.23744 2.21967 7.53033Z" fill="#212121" />
        </svg>
      </fluent-button>`}renderChevrons(){return f`${this.renderUpChevron()}${this.renderDownChevron()}`}renderDropdown(){return this._task.status===Q.PENDING||!this._treeViewState?this.renderLoadingIndicator():this._treeViewState?this._treeViewState.length===0&&this._inputValue.length>0?this.renderError():this.renderDropdownList(this._treeViewState):f``}renderDropdownList(t){if(t&&t.length>0){let e=null;return f`
        <fluent-tree-view
          class="tree-view"
          dir=${this.direction}
          aria-live="polite"
          aria-relevant="all"
          aria-atomic="true"
          aria-label=${this.strings.teamsChannels}
          aria-orientation="horizontal"
          @keydown=${this.onKeydownTreeView}>
          ${te(t,o=>o==null?void 0:o.item,o=>{var s;return o.channels&&(e=f`<img
                  class="team-photo"
                  alt="${this.strings.photoFor} ${o.item.displayName}"
                  src=${(s=this.teamsPhotos[o.item.id])===null||s===void 0?void 0:s.photo} />`),f`
                <fluent-tree-item
                  ?expanded=${o==null?void 0:o.isExpanded}
                  @click=${this.handleTeamTreeItemClick}>
                    <span slot="start">${e}</span>${o.item.displayName}
                    ${te(o==null?void 0:o.channels,i=>i.item,i=>this.renderItem(i))}
                </fluent-tree-item>`})}
        </fluent-tree-view>`}return null}renderItem(t){var e;return f`
      <fluent-tree-item
        id=${(e=t==null?void 0:t.item)===null||e===void 0?void 0:e.id}
        @keydown=${o=>this.onUserKeyDown(o,t)}
        @click=${()=>this.handleItemClick(t)}>
          ${t==null?void 0:t.item.displayName}
      </fluent-tree-item>`}loadState(){return ne(this,void 0,void 0,function*(){const t=P.globalProvider;let e;if(t&&t.state===Z.SignedIn){const o=t.graph.forComponent(this);e=yield At(o),e=e.filter(i=>!i.isArchived);const s=e.map(i=>i.id);this.teamsPhotos=yield Bt(o,s),this._items=yield Vt(o,e)}this.filterList(),this.resetFocusState()})}clearSelectedItem(){this.removeSelectedChannel()}removeSelectedChannel(t){this.selectChannel(t);const e=this.renderRoot.querySelectorAll("fluent-tree-item");e&&e.forEach(o=>{o.removeAttribute("expanded"),o.removeAttribute("selected")})}handleItemClick(t){t.channels?t.isExpanded=!t.isExpanded:(this.selectChannel(t),this.lostFocus())}onUserKeyDown(t,e){switch(t.code){case"Enter":this.selectChannel(e),this.resetFocusState(),this.lostFocus(),t.preventDefault();break;case"Backspace":this._inputValue.length===0&&this._selectedItemState&&(this.selectChannel(null),this.resetFocusState(),t.preventDefault());break}}filterList(){this.items&&(this._treeViewState=this.generateTreeViewState(this.items,this._inputValue),this.resetFocusState())}generateTreeViewState(t,e="",o=null){const s=[];if(e=e.toLowerCase(),t)for(const i of t){let r;if(e.length===0||i.item.displayName.toLowerCase().includes(e))r={item:i.item,parent:o},i.channels&&(r.channels=this.generateTreeViewState(i.channels,"",r),r.isExpanded=e.length>0);else if(i.channels){const l={item:i.item,parent:o},u=this.generateTreeViewState(i.channels,e,l);u.length>0&&(r=l,r.channels=u,r.isExpanded=!0)}r&&s.push(r)}return s}generateFocusList(t){if(!t||t.length===0)return[];let e=[];for(const o of t)e.push(o),o.channels&&o.isExpanded&&(e=[...e,...this.generateFocusList(o.channels)]);return e}resetFocusState(){this._focusList=this.generateFocusList(this._treeViewState),this.requestUpdate()}selectChannel(t){var e,o;t&&this._selectedItemState!==t?(e=this._input)===null||e===void 0||e.setAttribute("disabled","true"):(o=this._input)===null||o===void 0||o.removeAttribute("disabled"),this._selectedItemState=t,this.lostFocus(),this.fireCustomEvent("selectionChanged",this.selectedItem)}hideCloseIcon(){const t=this.renderRoot.querySelector(".close-icon");t&&(t.style.display="none")}toggleChevron(){const t=this.renderRoot.querySelector(".down-chevron"),e=this.renderRoot.querySelector(".up-chevron");this._isDropdownVisible?(t&&(t.style.display="none"),e&&(e.style.display=null)):(t&&(t.style.display=null,this.hideCloseIcon()),e&&(e.style.display="none")),this.hideCloseIcon()}}ue([ce(),pe("design:type",Object)],O.prototype,"_selectedItemState",void 0);ue([ce(),pe("design:type",Boolean)],O.prototype,"_isDropdownVisible",void 0);const Dt=U("teams-channel-picker",Rt),E=U("file-list",at),Pt=U("picker",lt),Mt=W({fileGrid:{paddingBottom:"10px"}}),zt=()=>{const[n,t]=S.useState(null),e=Mt();return d.jsxs("div",{children:[d.jsx(Dt,{selectionChanged:o=>t(o.detail),className:e.fileGrid}),n&&d.jsx(E,{groupId:n.team.id,itemPath:n.channel.displayName,pageSize:100})]})},Ht=W({picker:{paddingBottom:"10px",display:"block"}}),jt=()=>{const[n,t]=S.useState(null),[e,o]=S.useState(""),[s,i]=S.useState(""),r=Ht(),l=async u=>{if(u.detail.list.template==="documentLibrary"){const m=await P.globalProvider.graph.client.api(`/sites/root/lists/${u.detail.id}/drive`).get();t(u.detail),o(m.id),i("")}else t(null),o(""),i("Please select a document library")};return d.jsxs("div",{children:[d.jsx(Pt,{resource:"/sites/root/lists",placeholder:"Select a list",keyName:"displayName",selectionChanged:l,className:r.picker}),n&&e&&d.jsx(E,{itemPath:"/",driveId:e,pageSize:100}),s&&d.jsx("div",{children:s})]})},qt=W({panels:{...ot.padding("10px")}}),oo=()=>{const n=qt(),[t,e]=S.useState("my"),o=(s,i)=>{e(i.value)};return d.jsxs(d.Fragment,{children:[d.jsx(me,{title:"Files",description:"View your files from accross your OneDrive, channels you are a member of and your SharePoint sites"}),d.jsxs("div",{children:[d.jsxs(nt,{selectedValue:t,onTabSelect:o,children:[d.jsx(_,{value:"my",children:"My Files"}),d.jsx(_,{value:"recent",children:"Recent Files"}),d.jsx(_,{value:"site",children:"Site Files"}),d.jsx(_,{value:"channel",children:"Channel Files"})]}),d.jsxs("div",{className:n.panels,children:[t==="my"&&d.jsx(E,{pageSize:100}),t==="recent"&&d.jsx(E,{insightType:"used",enableFileUpload:!1,pageSize:100}),t==="site"&&d.jsx(jt,{}),t==="channel"&&d.jsx(zt,{})]})]})]})};export{oo as default};
