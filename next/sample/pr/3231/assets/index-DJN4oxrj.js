import{m as g,u as x,d0 as S,F as k,d1 as H,_ as i,b as s,o as p,c4 as I,D as w,n as u,bv as v,bw as A,bx as O,bt as L,aF as E,bu as M,d2 as _,y as h,z as f,bA as b,E as r,K as z,J as N,cz as R,d3 as D,G as B,I as m,P as q,d4 as G,d5 as P,d6 as j,d7 as K,bD as y,d8 as Q,bE as U,M as J,d9 as W,da as V,bQ as X,bP as Y,Q as F,R as o,db as Z,cc as ee,cd as $,ce as te,cf as ae,cg as oe}from"./App-U48ayLiK.js";import{_ as ie}from"./index-DwD95uW3.js";const le=(t,a)=>g`
    <template
        role="checkbox"
        aria-checked="${e=>e.checked}"
        aria-required="${e=>e.required}"
        aria-disabled="${e=>e.disabled}"
        aria-readonly="${e=>e.readOnly}"
        tabindex="${e=>e.disabled?null:0}"
        @keypress="${(e,c)=>e.keypressHandler(c.event)}"
        @click="${(e,c)=>e.clickHandler(c.event)}"
        class="${e=>e.readOnly?"readonly":""} ${e=>e.checked?"checked":""} ${e=>e.indeterminate?"indeterminate":""}"
    >
        <div part="control" class="control">
            <slot name="checked-indicator">
                ${a.checkedIndicator||""}
            </slot>
            <slot name="indeterminate-indicator">
                ${a.indeterminateIndicator||""}
            </slot>
        </div>
        <label
            part="label"
            class="${e=>e.defaultSlottedNodes&&e.defaultSlottedNodes.length?"label":"label label__hidden"}"
        >
            <slot ${x("defaultSlottedNodes")}></slot>
        </label>
    </template>
`;class se extends k{}class re extends S(se){constructor(){super(...arguments),this.proxy=document.createElement("input")}}class d extends re{constructor(){super(),this.initialValue="on",this.indeterminate=!1,this.keypressHandler=a=>{if(!this.readOnly)switch(a.key){case H:this.indeterminate&&(this.indeterminate=!1),this.checked=!this.checked;break}},this.clickHandler=a=>{!this.disabled&&!this.readOnly&&(this.indeterminate&&(this.indeterminate=!1),this.checked=!this.checked)},this.proxy.setAttribute("type","checkbox")}readOnlyChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.readOnly=this.readOnly)}}i([s({attribute:"readonly",mode:"boolean"})],d.prototype,"readOnly",void 0);i([p],d.prototype,"defaultSlottedNodes",void 0);i([p],d.prototype,"indeterminate",void 0);class ne extends k{}class de extends I(ne){constructor(){super(...arguments),this.proxy=document.createElement("input")}}const ce={email:"email",password:"password",tel:"tel",text:"text",url:"url"};let l=class extends de{constructor(){super(...arguments),this.type=ce.text}readOnlyChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.readOnly=this.readOnly,this.validate())}autofocusChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.autofocus=this.autofocus,this.validate())}placeholderChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.placeholder=this.placeholder)}typeChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.type=this.type,this.validate())}listChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.setAttribute("list",this.list),this.validate())}maxlengthChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.maxLength=this.maxlength,this.validate())}minlengthChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.minLength=this.minlength,this.validate())}patternChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.pattern=this.pattern,this.validate())}sizeChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.size=this.size)}spellcheckChanged(){this.proxy instanceof HTMLInputElement&&(this.proxy.spellcheck=this.spellcheck)}connectedCallback(){super.connectedCallback(),this.proxy.setAttribute("type",this.type),this.validate(),this.autofocus&&w.queueUpdate(()=>{this.focus()})}select(){this.control.select(),this.$emit("select")}handleTextInput(){this.value=this.control.value}handleChange(){this.$emit("change")}validate(){super.validate(this.control)}};i([s({attribute:"readonly",mode:"boolean"})],l.prototype,"readOnly",void 0);i([s({mode:"boolean"})],l.prototype,"autofocus",void 0);i([s],l.prototype,"placeholder",void 0);i([s],l.prototype,"type",void 0);i([s],l.prototype,"list",void 0);i([s({converter:u})],l.prototype,"maxlength",void 0);i([s({converter:u})],l.prototype,"minlength",void 0);i([s],l.prototype,"pattern",void 0);i([s({converter:u})],l.prototype,"size",void 0);i([s({mode:"boolean"})],l.prototype,"spellcheck",void 0);i([p],l.prototype,"defaultSlottedNodes",void 0);class T{}v(T,A);v(l,O,T);const he=(t,a)=>g`
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
                ${x({property:"defaultSlottedNodes",filter:_})}
            ></slot>
        </label>
        <div class="root" part="root">
            ${L(t,a)}
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
                ${E("control")}
            />
            ${M(t,a)}
        </div>
    </template>
`,pe=(t,a)=>h`
    ${f("inline-flex")} :host {
      align-items: center;
      outline: none;
      ${""} user-select: none;
    }

    .control {
      position: relative;
      width: calc((${b} / 2 + ${r}) * 1px);
      height: calc((${b} / 2 + ${r}) * 1px);
      box-sizing: border-box;
      border-radius: calc(${z} * 1px);
      border: calc(${N} * 1px) solid ${R};
      background: ${D};
      cursor: pointer;
    }

    .label__hidden {
      display: none;
      visibility: hidden;
    }

    .label {
      ${B}
      color: ${m};
      ${""} padding-inline-start: calc(${r} * 2px + 2px);
      margin-inline-end: calc(${r} * 2px + 2px);
      cursor: pointer;
    }

    slot[name='checked-indicator'],
    slot[name='indeterminate-indicator'] {
      display: flex;
      align-items: center;
      justify-content: center;
      width: 100%;
      height: 100%;
      fill: ${m};
      opacity: 0;
      pointer-events: none;
    }

    slot[name='indeterminate-indicator'] {
      position: absolute;
      top: 0;
    }

    :host(.checked) slot[name='checked-indicator'],
    :host(.checked) slot[name='indeterminate-indicator'] {
      fill: ${q};
    }

    :host(:not(.disabled):hover) .control {
      background: ${G};
      border-color: ${P};
    }

    :host(:not(.disabled):active) .control {
      background: ${j};
      border-color: ${K};
    }

    :host(:${y}) .control {
      background: ${Q};
      ${U}
    }

    :host(.checked) .control {
      background: ${J};
      border-color: transparent;
    }

    :host(.checked:not(.disabled):hover) .control {
      background: ${W};
      border-color: transparent;
    }

    :host(.checked:not(.disabled):active) .control {
      background: ${V};
      border-color: transparent;
    }

    :host(.disabled) .label,
    :host(.readonly) .label,
    :host(.readonly) .control,
    :host(.disabled) .control {
      cursor: ${X};
    }

    :host(.checked:not(.indeterminate)) slot[name='checked-indicator'],
    :host(.indeterminate) slot[name='indeterminate-indicator'] {
      opacity: 1;
    }

    :host(.disabled) {
      opacity: ${Y};
    }
  `.withBehaviors(F(h`
        .control {
          border-color: ${o.FieldText};
          background: ${o.Field};
        }
        :host(:not(.disabled):hover) .control,
        :host(:not(.disabled):active) .control {
          border-color: ${o.Highlight};
          background: ${o.Field};
        }
        slot[name='checked-indicator'],
        slot[name='indeterminate-indicator'] {
          fill: ${o.FieldText};
        }
        :host(:${y}) .control {
          forced-color-adjust: none;
          outline-color: ${o.FieldText};
          background: ${o.Field};
          border-color: ${o.Highlight};
        }
        :host(.checked) .control {
          background: ${o.Highlight};
          border-color: ${o.Highlight};
        }
        :host(.checked:not(.disabled):hover) .control,
        :host(.checked:not(.disabled):active) .control {
          background: ${o.HighlightText};
          border-color: ${o.Highlight};
        }
        :host(.checked) slot[name='checked-indicator'],
        :host(.checked) slot[name='indeterminate-indicator'] {
          fill: ${o.HighlightText};
        }
        :host(.checked:hover ) .control slot[name='checked-indicator'],
        :host(.checked:hover ) .control slot[name='indeterminate-indicator'] {
          fill: ${o.Highlight};
        }
        :host(.disabled) {
          opacity: 1;
        }
        :host(.disabled) .control {
          border-color: ${o.GrayText};
          background: ${o.Field};
        }
        :host(.disabled) slot[name='checked-indicator'],
        :host(.checked.disabled:hover) .control slot[name='checked-indicator'],
        :host(.disabled) slot[name='indeterminate-indicator'],
        :host(.checked.disabled:hover) .control slot[name='indeterminate-indicator'] {
          fill: ${o.GrayText};
        }
      `)),$e=d.compose({baseName:"checkbox",template:le,styles:pe,checkedIndicator:`
    <svg width="16" height="16" xmlns="http://www.w3.org/2000/svg">
      <path d="M13.86 3.66a.5.5 0 01-.02.7l-7.93 7.48a.6.6 0 01-.84-.02L2.4 9.1a.5.5 0 01.72-.7l2.4 2.44 7.65-7.2a.5.5 0 01.7.02z"/>
    </svg>
  `,indeterminateIndicator:`
    <svg width="16" height="16" xmlns="http://www.w3.org/2000/svg">
      <path d="M3 8c0-.28.22-.5.5-.5h9a.5.5 0 010 1h-9A.5.5 0 013 8z"/>
    </svg>
  `}),n=".root",ue=(t,a)=>h`
    ${f("inline-block")}

    ${Z(t,a,n)}

    ${ee()}

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
      padding: 0 calc(${r} * 2px + 1px);
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
  `.withBehaviors($("outline",te(t,a,n)),$("filled",ae(t,a,n)),F(oe(t,a,n)));class C extends l{appearanceChanged(a,e){a!==e&&(this.classList.add(e),this.classList.remove(a))}connectedCallback(){super.connectedCallback(),this.appearance||(this.appearance="outline")}}ie([s],C.prototype,"appearance",void 0);const ge=C.compose({baseName:"text-field",baseClass:l,template:he,styles:ue,shadowOptions:{delegatesFocus:!0}});export{$e as a,ge as f};
