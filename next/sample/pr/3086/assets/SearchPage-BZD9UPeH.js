import{G as $e,b as Ce,a as Se,p as Ee,j as n,r as H}from"./index-4aN9tDwu.js";import{P as Le}from"./PageHeader-DPoFmWE1.js";import{aD as Ie,F as B,v as J,D as me,_ as l,b as u,o as R,m as F,w as ve,aE as Me,p as Fe,x as O,y as K,I as M,aF as ae,C as Oe,aG as le,aH as je,J as be,P as ye,Q as Ae,aI as W,M as Z,G as _e,E as Ue,aJ as qe,aK as Ve,W as He,Z as Ne,at as Be,a1 as We,X as Xe,Y as f,aL as he,aM as _,aN as X,aO as Ye,aP as Ge,a5 as E,a2 as T,a3 as z,a6 as U,aQ as L,aR as ce,aS as I,aT as Y,a0 as k,a8 as Qe,$ as Je,a9 as we,an as Re,aU as Ke,aq as xe,aV as de,ap as Ze,aW as et,ar as tt,as as q}from"./App-ZYGtWVu9.js";import{r as Pe}from"./mgt-file-BwfpROew.js";import{g as ee,l as ue,m as pe,O as it,D as st,d as rt,f as ge,h as ot,i as nt,j as at,k as G,T as Q,S as V}from"./DataGridHeaderCell-DiyCUpcb.js";import"./graph.files-GhcJkAzL.js";import"./useFocusWithin-EZiYdyQg.js";const lt="beta";class N extends $e{static fromGraph(e){const t=new N(e.client);return t.setComponent(e.componentName),t}constructor(e,t=lt){super(e,t)}forComponent(e){const t=new N(this.client);return this.setComponent(e),t}}class ht{constructor(){this.intersectionDetector=null,this.observedElements=new Map,this.requestPosition=(e,t)=>{var i;if(this.intersectionDetector!==null){if(this.observedElements.has(e)){(i=this.observedElements.get(e))===null||i===void 0||i.push(t);return}this.observedElements.set(e,[t]),this.intersectionDetector.observe(e)}},this.cancelRequestPosition=(e,t)=>{const i=this.observedElements.get(e);if(i!==void 0){const r=i.indexOf(t);r!==-1&&i.splice(r,1)}},this.initializeIntersectionDetector=()=>{Ie.IntersectionObserver&&(this.intersectionDetector=new IntersectionObserver(this.handleIntersection,{root:null,rootMargin:"0px",threshold:[0,1]}))},this.handleIntersection=e=>{if(this.intersectionDetector===null)return;const t=[],i=[];e.forEach(r=>{var o;(o=this.intersectionDetector)===null||o===void 0||o.unobserve(r.target);const h=this.observedElements.get(r.target);h!==void 0&&(h.forEach(c=>{let d=t.indexOf(c);d===-1&&(d=t.length,t.push(c),i.push([])),i[d].push(r)}),this.observedElements.delete(r.target))}),t.forEach((r,o)=>{r(i[o])})},this.initializeIntersectionDetector()}}class a extends B{constructor(){super(...arguments),this.anchor="",this.viewport="",this.horizontalPositioningMode="uncontrolled",this.horizontalDefaultPosition="unset",this.horizontalViewportLock=!1,this.horizontalInset=!1,this.horizontalScaling="content",this.verticalPositioningMode="uncontrolled",this.verticalDefaultPosition="unset",this.verticalViewportLock=!1,this.verticalInset=!1,this.verticalScaling="content",this.fixedPlacement=!1,this.autoUpdateMode="anchor",this.anchorElement=null,this.viewportElement=null,this.initialLayoutComplete=!1,this.resizeDetector=null,this.baseHorizontalOffset=0,this.baseVerticalOffset=0,this.pendingPositioningUpdate=!1,this.pendingReset=!1,this.currentDirection=J.ltr,this.regionVisible=!1,this.forceUpdate=!1,this.updateThreshold=.5,this.update=()=>{this.pendingPositioningUpdate||this.requestPositionUpdates()},this.startObservers=()=>{this.stopObservers(),this.anchorElement!==null&&(this.requestPositionUpdates(),this.resizeDetector!==null&&(this.resizeDetector.observe(this.anchorElement),this.resizeDetector.observe(this)))},this.requestPositionUpdates=()=>{this.anchorElement===null||this.pendingPositioningUpdate||(a.intersectionService.requestPosition(this,this.handleIntersection),a.intersectionService.requestPosition(this.anchorElement,this.handleIntersection),this.viewportElement!==null&&a.intersectionService.requestPosition(this.viewportElement,this.handleIntersection),this.pendingPositioningUpdate=!0)},this.stopObservers=()=>{this.pendingPositioningUpdate&&(this.pendingPositioningUpdate=!1,a.intersectionService.cancelRequestPosition(this,this.handleIntersection),this.anchorElement!==null&&a.intersectionService.cancelRequestPosition(this.anchorElement,this.handleIntersection),this.viewportElement!==null&&a.intersectionService.cancelRequestPosition(this.viewportElement,this.handleIntersection)),this.resizeDetector!==null&&this.resizeDetector.disconnect()},this.getViewport=()=>typeof this.viewport!="string"||this.viewport===""?document.documentElement:document.getElementById(this.viewport),this.getAnchor=()=>document.getElementById(this.anchor),this.handleIntersection=e=>{this.pendingPositioningUpdate&&(this.pendingPositioningUpdate=!1,this.applyIntersectionEntries(e)&&this.updateLayout())},this.applyIntersectionEntries=e=>{const t=e.find(o=>o.target===this),i=e.find(o=>o.target===this.anchorElement),r=e.find(o=>o.target===this.viewportElement);return t===void 0||r===void 0||i===void 0?!1:!this.regionVisible||this.forceUpdate||this.regionRect===void 0||this.anchorRect===void 0||this.viewportRect===void 0||this.isRectDifferent(this.anchorRect,i.boundingClientRect)||this.isRectDifferent(this.viewportRect,r.boundingClientRect)||this.isRectDifferent(this.regionRect,t.boundingClientRect)?(this.regionRect=t.boundingClientRect,this.anchorRect=i.boundingClientRect,this.viewportElement===document.documentElement?this.viewportRect=new DOMRectReadOnly(r.boundingClientRect.x+document.documentElement.scrollLeft,r.boundingClientRect.y+document.documentElement.scrollTop,r.boundingClientRect.width,r.boundingClientRect.height):this.viewportRect=r.boundingClientRect,this.updateRegionOffset(),this.forceUpdate=!1,!0):!1},this.updateRegionOffset=()=>{this.anchorRect&&this.regionRect&&(this.baseHorizontalOffset=this.baseHorizontalOffset+(this.anchorRect.left-this.regionRect.left)+(this.translateX-this.baseHorizontalOffset),this.baseVerticalOffset=this.baseVerticalOffset+(this.anchorRect.top-this.regionRect.top)+(this.translateY-this.baseVerticalOffset))},this.isRectDifferent=(e,t)=>Math.abs(e.top-t.top)>this.updateThreshold||Math.abs(e.right-t.right)>this.updateThreshold||Math.abs(e.bottom-t.bottom)>this.updateThreshold||Math.abs(e.left-t.left)>this.updateThreshold,this.handleResize=e=>{this.update()},this.reset=()=>{this.pendingReset&&(this.pendingReset=!1,this.anchorElement===null&&(this.anchorElement=this.getAnchor()),this.viewportElement===null&&(this.viewportElement=this.getViewport()),this.currentDirection=ee(this),this.startObservers())},this.updateLayout=()=>{let e,t;if(this.horizontalPositioningMode!=="uncontrolled"){const o=this.getPositioningOptions(this.horizontalInset);if(this.horizontalDefaultPosition==="center")t="center";else if(this.horizontalDefaultPosition!=="unset"){let b=this.horizontalDefaultPosition;if(b==="start"||b==="end"){const S=ee(this);if(S!==this.currentDirection){this.currentDirection=S,this.initialize();return}this.currentDirection===J.ltr?b=b==="start"?"left":"right":b=b==="start"?"right":"left"}switch(b){case"left":t=this.horizontalInset?"insetStart":"start";break;case"right":t=this.horizontalInset?"insetEnd":"end";break}}const h=this.horizontalThreshold!==void 0?this.horizontalThreshold:this.regionRect!==void 0?this.regionRect.width:0,c=this.anchorRect!==void 0?this.anchorRect.left:0,d=this.anchorRect!==void 0?this.anchorRect.right:0,g=this.anchorRect!==void 0?this.anchorRect.width:0,m=this.viewportRect!==void 0?this.viewportRect.left:0,p=this.viewportRect!==void 0?this.viewportRect.right:0;(t===void 0||this.horizontalPositioningMode!=="locktodefault"&&this.getAvailableSpace(t,c,d,g,m,p)<h)&&(t=this.getAvailableSpace(o[0],c,d,g,m,p)>this.getAvailableSpace(o[1],c,d,g,m,p)?o[0]:o[1])}if(this.verticalPositioningMode!=="uncontrolled"){const o=this.getPositioningOptions(this.verticalInset);if(this.verticalDefaultPosition==="center")e="center";else if(this.verticalDefaultPosition!=="unset")switch(this.verticalDefaultPosition){case"top":e=this.verticalInset?"insetStart":"start";break;case"bottom":e=this.verticalInset?"insetEnd":"end";break}const h=this.verticalThreshold!==void 0?this.verticalThreshold:this.regionRect!==void 0?this.regionRect.height:0,c=this.anchorRect!==void 0?this.anchorRect.top:0,d=this.anchorRect!==void 0?this.anchorRect.bottom:0,g=this.anchorRect!==void 0?this.anchorRect.height:0,m=this.viewportRect!==void 0?this.viewportRect.top:0,p=this.viewportRect!==void 0?this.viewportRect.bottom:0;(e===void 0||this.verticalPositioningMode!=="locktodefault"&&this.getAvailableSpace(e,c,d,g,m,p)<h)&&(e=this.getAvailableSpace(o[0],c,d,g,m,p)>this.getAvailableSpace(o[1],c,d,g,m,p)?o[0]:o[1])}const i=this.getNextRegionDimension(t,e),r=this.horizontalPosition!==t||this.verticalPosition!==e;if(this.setHorizontalPosition(t,i),this.setVerticalPosition(e,i),this.updateRegionStyle(),!this.initialLayoutComplete){this.initialLayoutComplete=!0,this.requestPositionUpdates();return}this.regionVisible||(this.regionVisible=!0,this.style.removeProperty("pointer-events"),this.style.removeProperty("opacity"),this.classList.toggle("loaded",!0),this.$emit("loaded",this,{bubbles:!1})),this.updatePositionClasses(),r&&this.$emit("positionchange",this,{bubbles:!1})},this.updateRegionStyle=()=>{this.style.width=this.regionWidth,this.style.height=this.regionHeight,this.style.transform=`translate(${this.translateX}px, ${this.translateY}px)`},this.updatePositionClasses=()=>{this.classList.toggle("top",this.verticalPosition==="start"),this.classList.toggle("bottom",this.verticalPosition==="end"),this.classList.toggle("inset-top",this.verticalPosition==="insetStart"),this.classList.toggle("inset-bottom",this.verticalPosition==="insetEnd"),this.classList.toggle("vertical-center",this.verticalPosition==="center"),this.classList.toggle("left",this.horizontalPosition==="start"),this.classList.toggle("right",this.horizontalPosition==="end"),this.classList.toggle("inset-left",this.horizontalPosition==="insetStart"),this.classList.toggle("inset-right",this.horizontalPosition==="insetEnd"),this.classList.toggle("horizontal-center",this.horizontalPosition==="center")},this.setHorizontalPosition=(e,t)=>{if(e===void 0||this.regionRect===void 0||this.anchorRect===void 0||this.viewportRect===void 0)return;let i=0;switch(this.horizontalScaling){case"anchor":case"fill":i=this.horizontalViewportLock?this.viewportRect.width:t.width,this.regionWidth=`${i}px`;break;case"content":i=this.regionRect.width,this.regionWidth="unset";break}let r=0;switch(e){case"start":this.translateX=this.baseHorizontalOffset-i,this.horizontalViewportLock&&this.anchorRect.left>this.viewportRect.right&&(this.translateX=this.translateX-(this.anchorRect.left-this.viewportRect.right));break;case"insetStart":this.translateX=this.baseHorizontalOffset-i+this.anchorRect.width,this.horizontalViewportLock&&this.anchorRect.right>this.viewportRect.right&&(this.translateX=this.translateX-(this.anchorRect.right-this.viewportRect.right));break;case"insetEnd":this.translateX=this.baseHorizontalOffset,this.horizontalViewportLock&&this.anchorRect.left<this.viewportRect.left&&(this.translateX=this.translateX-(this.anchorRect.left-this.viewportRect.left));break;case"end":this.translateX=this.baseHorizontalOffset+this.anchorRect.width,this.horizontalViewportLock&&this.anchorRect.right<this.viewportRect.left&&(this.translateX=this.translateX-(this.anchorRect.right-this.viewportRect.left));break;case"center":if(r=(this.anchorRect.width-i)/2,this.translateX=this.baseHorizontalOffset+r,this.horizontalViewportLock){const o=this.anchorRect.left+r,h=this.anchorRect.right-r;o<this.viewportRect.left&&!(h>this.viewportRect.right)?this.translateX=this.translateX-(o-this.viewportRect.left):h>this.viewportRect.right&&!(o<this.viewportRect.left)&&(this.translateX=this.translateX-(h-this.viewportRect.right))}break}this.horizontalPosition=e},this.setVerticalPosition=(e,t)=>{if(e===void 0||this.regionRect===void 0||this.anchorRect===void 0||this.viewportRect===void 0)return;let i=0;switch(this.verticalScaling){case"anchor":case"fill":i=this.verticalViewportLock?this.viewportRect.height:t.height,this.regionHeight=`${i}px`;break;case"content":i=this.regionRect.height,this.regionHeight="unset";break}let r=0;switch(e){case"start":this.translateY=this.baseVerticalOffset-i,this.verticalViewportLock&&this.anchorRect.top>this.viewportRect.bottom&&(this.translateY=this.translateY-(this.anchorRect.top-this.viewportRect.bottom));break;case"insetStart":this.translateY=this.baseVerticalOffset-i+this.anchorRect.height,this.verticalViewportLock&&this.anchorRect.bottom>this.viewportRect.bottom&&(this.translateY=this.translateY-(this.anchorRect.bottom-this.viewportRect.bottom));break;case"insetEnd":this.translateY=this.baseVerticalOffset,this.verticalViewportLock&&this.anchorRect.top<this.viewportRect.top&&(this.translateY=this.translateY-(this.anchorRect.top-this.viewportRect.top));break;case"end":this.translateY=this.baseVerticalOffset+this.anchorRect.height,this.verticalViewportLock&&this.anchorRect.bottom<this.viewportRect.top&&(this.translateY=this.translateY-(this.anchorRect.bottom-this.viewportRect.top));break;case"center":if(r=(this.anchorRect.height-i)/2,this.translateY=this.baseVerticalOffset+r,this.verticalViewportLock){const o=this.anchorRect.top+r,h=this.anchorRect.bottom-r;o<this.viewportRect.top&&!(h>this.viewportRect.bottom)?this.translateY=this.translateY-(o-this.viewportRect.top):h>this.viewportRect.bottom&&!(o<this.viewportRect.top)&&(this.translateY=this.translateY-(h-this.viewportRect.bottom))}}this.verticalPosition=e},this.getPositioningOptions=e=>e?["insetStart","insetEnd"]:["start","end"],this.getAvailableSpace=(e,t,i,r,o,h)=>{const c=t-o,d=h-(t+r);switch(e){case"start":return c;case"insetStart":return c+r;case"insetEnd":return d+r;case"end":return d;case"center":return Math.min(c,d)*2+r}},this.getNextRegionDimension=(e,t)=>{const i={height:this.regionRect!==void 0?this.regionRect.height:0,width:this.regionRect!==void 0?this.regionRect.width:0};return e!==void 0&&this.horizontalScaling==="fill"?i.width=this.getAvailableSpace(e,this.anchorRect!==void 0?this.anchorRect.left:0,this.anchorRect!==void 0?this.anchorRect.right:0,this.anchorRect!==void 0?this.anchorRect.width:0,this.viewportRect!==void 0?this.viewportRect.left:0,this.viewportRect!==void 0?this.viewportRect.right:0):this.horizontalScaling==="anchor"&&(i.width=this.anchorRect!==void 0?this.anchorRect.width:0),t!==void 0&&this.verticalScaling==="fill"?i.height=this.getAvailableSpace(t,this.anchorRect!==void 0?this.anchorRect.top:0,this.anchorRect!==void 0?this.anchorRect.bottom:0,this.anchorRect!==void 0?this.anchorRect.height:0,this.viewportRect!==void 0?this.viewportRect.top:0,this.viewportRect!==void 0?this.viewportRect.bottom:0):this.verticalScaling==="anchor"&&(i.height=this.anchorRect!==void 0?this.anchorRect.height:0),i},this.startAutoUpdateEventListeners=()=>{window.addEventListener(ue,this.update,{passive:!0}),window.addEventListener(pe,this.update,{passive:!0,capture:!0}),this.resizeDetector!==null&&this.viewportElement!==null&&this.resizeDetector.observe(this.viewportElement)},this.stopAutoUpdateEventListeners=()=>{window.removeEventListener(ue,this.update),window.removeEventListener(pe,this.update),this.resizeDetector!==null&&this.viewportElement!==null&&this.resizeDetector.unobserve(this.viewportElement)}}anchorChanged(){this.initialLayoutComplete&&(this.anchorElement=this.getAnchor())}viewportChanged(){this.initialLayoutComplete&&(this.viewportElement=this.getViewport())}horizontalPositioningModeChanged(){this.requestReset()}horizontalDefaultPositionChanged(){this.updateForAttributeChange()}horizontalViewportLockChanged(){this.updateForAttributeChange()}horizontalInsetChanged(){this.updateForAttributeChange()}horizontalThresholdChanged(){this.updateForAttributeChange()}horizontalScalingChanged(){this.updateForAttributeChange()}verticalPositioningModeChanged(){this.requestReset()}verticalDefaultPositionChanged(){this.updateForAttributeChange()}verticalViewportLockChanged(){this.updateForAttributeChange()}verticalInsetChanged(){this.updateForAttributeChange()}verticalThresholdChanged(){this.updateForAttributeChange()}verticalScalingChanged(){this.updateForAttributeChange()}fixedPlacementChanged(){this.$fastController.isConnected&&this.initialLayoutComplete&&this.initialize()}autoUpdateModeChanged(e,t){this.$fastController.isConnected&&this.initialLayoutComplete&&(e==="auto"&&this.stopAutoUpdateEventListeners(),t==="auto"&&this.startAutoUpdateEventListeners())}anchorElementChanged(){this.requestReset()}viewportElementChanged(){this.$fastController.isConnected&&this.initialLayoutComplete&&this.initialize()}connectedCallback(){super.connectedCallback(),this.autoUpdateMode==="auto"&&this.startAutoUpdateEventListeners(),this.initialize()}disconnectedCallback(){super.disconnectedCallback(),this.autoUpdateMode==="auto"&&this.stopAutoUpdateEventListeners(),this.stopObservers(),this.disconnectResizeDetector()}adoptedCallback(){this.initialize()}disconnectResizeDetector(){this.resizeDetector!==null&&(this.resizeDetector.disconnect(),this.resizeDetector=null)}initializeResizeDetector(){this.disconnectResizeDetector(),this.resizeDetector=new window.ResizeObserver(this.handleResize)}updateForAttributeChange(){this.$fastController.isConnected&&this.initialLayoutComplete&&(this.forceUpdate=!0,this.update())}initialize(){this.initializeResizeDetector(),this.anchorElement===null&&(this.anchorElement=this.getAnchor()),this.requestReset()}requestReset(){this.$fastController.isConnected&&this.pendingReset===!1&&(this.setInitialState(),me.queueUpdate(()=>this.reset()),this.pendingReset=!0)}setInitialState(){this.initialLayoutComplete=!1,this.regionVisible=!1,this.translateX=0,this.translateY=0,this.baseHorizontalOffset=0,this.baseVerticalOffset=0,this.viewportRect=void 0,this.regionRect=void 0,this.anchorRect=void 0,this.verticalPosition=void 0,this.horizontalPosition=void 0,this.style.opacity="0",this.style.pointerEvents="none",this.forceUpdate=!1,this.style.position=this.fixedPlacement?"fixed":"absolute",this.updatePositionClasses(),this.updateRegionStyle()}}a.intersectionService=new ht;l([u],a.prototype,"anchor",void 0);l([u],a.prototype,"viewport",void 0);l([u({attribute:"horizontal-positioning-mode"})],a.prototype,"horizontalPositioningMode",void 0);l([u({attribute:"horizontal-default-position"})],a.prototype,"horizontalDefaultPosition",void 0);l([u({attribute:"horizontal-viewport-lock",mode:"boolean"})],a.prototype,"horizontalViewportLock",void 0);l([u({attribute:"horizontal-inset",mode:"boolean"})],a.prototype,"horizontalInset",void 0);l([u({attribute:"horizontal-threshold"})],a.prototype,"horizontalThreshold",void 0);l([u({attribute:"horizontal-scaling"})],a.prototype,"horizontalScaling",void 0);l([u({attribute:"vertical-positioning-mode"})],a.prototype,"verticalPositioningMode",void 0);l([u({attribute:"vertical-default-position"})],a.prototype,"verticalDefaultPosition",void 0);l([u({attribute:"vertical-viewport-lock",mode:"boolean"})],a.prototype,"verticalViewportLock",void 0);l([u({attribute:"vertical-inset",mode:"boolean"})],a.prototype,"verticalInset",void 0);l([u({attribute:"vertical-threshold"})],a.prototype,"verticalThreshold",void 0);l([u({attribute:"vertical-scaling"})],a.prototype,"verticalScaling",void 0);l([u({attribute:"fixed-placement",mode:"boolean"})],a.prototype,"fixedPlacement",void 0);l([u({attribute:"auto-update-mode"})],a.prototype,"autoUpdateMode",void 0);l([R],a.prototype,"anchorElement",void 0);l([R],a.prototype,"viewportElement",void 0);l([R],a.prototype,"initialLayoutComplete",void 0);const ct=(s,e)=>F`
    <template role="${t=>t.role}" aria-orientation="${t=>t.orientation}"></template>
`,dt={separator:"separator",presentation:"presentation"};class te extends B{constructor(){super(...arguments),this.role=dt.separator,this.orientation=it.horizontal}}l([u],te.prototype,"role",void 0);l([u],te.prototype,"orientation",void 0);const ut=(s,e)=>F`
    <template
        class="${t=>t.shape==="circle"?"circle":"rect"}"
        pattern="${t=>t.pattern}"
        ?shimmer="${t=>t.shimmer}"
    >
        ${ve(t=>t.shimmer===!0,F`
                <span class="shimmer"></span>
            `)}
        <object type="image/svg+xml" data="${t=>t.pattern}" role="presentation">
            <img class="pattern" src="${t=>t.pattern}" />
        </object>
        <slot></slot>
    </template>
`;class A extends B{constructor(){super(...arguments),this.shape="rect"}}l([u],A.prototype,"fill",void 0);l([u],A.prototype,"shape",void 0);l([u],A.prototype,"pattern",void 0);l([u({mode:"boolean"})],A.prototype,"shimmer",void 0);const pt=(s,e)=>F`
        ${ve(t=>t.tooltipVisible,F`
            <${s.tagFor(a)}
                fixed-placement="true"
                auto-update-mode="${t=>t.autoUpdateMode}"
                vertical-positioning-mode="${t=>t.verticalPositioningMode}"
                vertical-default-position="${t=>t.verticalDefaultPosition}"
                vertical-inset="${t=>t.verticalInset}"
                vertical-scaling="${t=>t.verticalScaling}"
                horizontal-positioning-mode="${t=>t.horizontalPositioningMode}"
                horizontal-default-position="${t=>t.horizontalDefaultPosition}"
                horizontal-scaling="${t=>t.horizontalScaling}"
                horizontal-inset="${t=>t.horizontalInset}"
                vertical-viewport-lock="${t=>t.horizontalViewportLock}"
                horizontal-viewport-lock="${t=>t.verticalViewportLock}"
                dir="${t=>t.currentDirection}"
                ${Me("region")}
            >
                <div class="tooltip" part="tooltip" role="tooltip">
                    <slot></slot>
                </div>
            </${s.tagFor(a)}>
        `)}
    `,P={top:"top",right:"right",bottom:"bottom",left:"left",start:"start",end:"end",topLeft:"top-left",topRight:"top-right",bottomLeft:"bottom-left",bottomRight:"bottom-right",topStart:"top-start",topEnd:"top-end",bottomStart:"bottom-start",bottomEnd:"bottom-end"};let v=class extends B{constructor(){super(...arguments),this.anchor="",this.delay=300,this.autoUpdateMode="anchor",this.anchorElement=null,this.viewportElement=null,this.verticalPositioningMode="dynamic",this.horizontalPositioningMode="dynamic",this.horizontalInset="false",this.verticalInset="false",this.horizontalScaling="content",this.verticalScaling="content",this.verticalDefaultPosition=void 0,this.horizontalDefaultPosition=void 0,this.tooltipVisible=!1,this.currentDirection=J.ltr,this.showDelayTimer=null,this.hideDelayTimer=null,this.isAnchorHoveredFocused=!1,this.isRegionHovered=!1,this.handlePositionChange=e=>{this.classList.toggle("top",this.region.verticalPosition==="start"),this.classList.toggle("bottom",this.region.verticalPosition==="end"),this.classList.toggle("inset-top",this.region.verticalPosition==="insetStart"),this.classList.toggle("inset-bottom",this.region.verticalPosition==="insetEnd"),this.classList.toggle("center-vertical",this.region.verticalPosition==="center"),this.classList.toggle("left",this.region.horizontalPosition==="start"),this.classList.toggle("right",this.region.horizontalPosition==="end"),this.classList.toggle("inset-left",this.region.horizontalPosition==="insetStart"),this.classList.toggle("inset-right",this.region.horizontalPosition==="insetEnd"),this.classList.toggle("center-horizontal",this.region.horizontalPosition==="center")},this.handleRegionMouseOver=e=>{this.isRegionHovered=!0},this.handleRegionMouseOut=e=>{this.isRegionHovered=!1,this.startHideDelayTimer()},this.handleAnchorMouseOver=e=>{if(this.tooltipVisible){this.isAnchorHoveredFocused=!0;return}this.startShowDelayTimer()},this.handleAnchorMouseOut=e=>{this.isAnchorHoveredFocused=!1,this.clearShowDelayTimer(),this.startHideDelayTimer()},this.handleAnchorFocusIn=e=>{this.startShowDelayTimer()},this.handleAnchorFocusOut=e=>{this.isAnchorHoveredFocused=!1,this.clearShowDelayTimer(),this.startHideDelayTimer()},this.startHideDelayTimer=()=>{this.clearHideDelayTimer(),this.tooltipVisible&&(this.hideDelayTimer=window.setTimeout(()=>{this.updateTooltipVisibility()},60))},this.clearHideDelayTimer=()=>{this.hideDelayTimer!==null&&(clearTimeout(this.hideDelayTimer),this.hideDelayTimer=null)},this.startShowDelayTimer=()=>{if(!this.isAnchorHoveredFocused){if(this.delay>1){this.showDelayTimer===null&&(this.showDelayTimer=window.setTimeout(()=>{this.startHover()},this.delay));return}this.startHover()}},this.startHover=()=>{this.isAnchorHoveredFocused=!0,this.updateTooltipVisibility()},this.clearShowDelayTimer=()=>{this.showDelayTimer!==null&&(clearTimeout(this.showDelayTimer),this.showDelayTimer=null)},this.getAnchor=()=>{const e=this.getRootNode();return e instanceof ShadowRoot?e.getElementById(this.anchor):document.getElementById(this.anchor)},this.handleDocumentKeydown=e=>{if(!e.defaultPrevented&&this.tooltipVisible)switch(e.key){case Fe:this.isAnchorHoveredFocused=!1,this.updateTooltipVisibility(),this.$emit("dismiss");break}},this.updateTooltipVisibility=()=>{if(this.visible===!1)this.hideTooltip();else if(this.visible===!0){this.showTooltip();return}else{if(this.isAnchorHoveredFocused||this.isRegionHovered){this.showTooltip();return}this.hideTooltip()}},this.showTooltip=()=>{this.tooltipVisible||(this.currentDirection=ee(this),this.tooltipVisible=!0,document.addEventListener("keydown",this.handleDocumentKeydown),me.queueUpdate(this.setRegionProps))},this.hideTooltip=()=>{this.tooltipVisible&&(this.clearHideDelayTimer(),this.region!==null&&this.region!==void 0&&(this.region.removeEventListener("positionchange",this.handlePositionChange),this.region.viewportElement=null,this.region.anchorElement=null,this.region.removeEventListener("mouseover",this.handleRegionMouseOver),this.region.removeEventListener("mouseout",this.handleRegionMouseOut)),document.removeEventListener("keydown",this.handleDocumentKeydown),this.tooltipVisible=!1)},this.setRegionProps=()=>{this.tooltipVisible&&(this.region.viewportElement=this.viewportElement,this.region.anchorElement=this.anchorElement,this.region.addEventListener("positionchange",this.handlePositionChange),this.region.addEventListener("mouseover",this.handleRegionMouseOver,{passive:!0}),this.region.addEventListener("mouseout",this.handleRegionMouseOut,{passive:!0}))}}visibleChanged(){this.$fastController.isConnected&&(this.updateTooltipVisibility(),this.updateLayout())}anchorChanged(){this.$fastController.isConnected&&(this.anchorElement=this.getAnchor())}positionChanged(){this.$fastController.isConnected&&this.updateLayout()}anchorElementChanged(e){if(this.$fastController.isConnected){if(e!=null&&(e.removeEventListener("mouseover",this.handleAnchorMouseOver),e.removeEventListener("mouseout",this.handleAnchorMouseOut),e.removeEventListener("focusin",this.handleAnchorFocusIn),e.removeEventListener("focusout",this.handleAnchorFocusOut)),this.anchorElement!==null&&this.anchorElement!==void 0){this.anchorElement.addEventListener("mouseover",this.handleAnchorMouseOver,{passive:!0}),this.anchorElement.addEventListener("mouseout",this.handleAnchorMouseOut,{passive:!0}),this.anchorElement.addEventListener("focusin",this.handleAnchorFocusIn,{passive:!0}),this.anchorElement.addEventListener("focusout",this.handleAnchorFocusOut,{passive:!0});const t=this.anchorElement.id;this.anchorElement.parentElement!==null&&this.anchorElement.parentElement.querySelectorAll(":hover").forEach(i=>{i.id===t&&this.startShowDelayTimer()})}this.region!==null&&this.region!==void 0&&this.tooltipVisible&&(this.region.anchorElement=this.anchorElement),this.updateLayout()}}viewportElementChanged(){this.region!==null&&this.region!==void 0&&(this.region.viewportElement=this.viewportElement),this.updateLayout()}connectedCallback(){super.connectedCallback(),this.anchorElement=this.getAnchor(),this.updateTooltipVisibility()}disconnectedCallback(){this.hideTooltip(),this.clearShowDelayTimer(),this.clearHideDelayTimer(),super.disconnectedCallback()}updateLayout(){switch(this.verticalPositioningMode="locktodefault",this.horizontalPositioningMode="locktodefault",this.position){case P.top:case P.bottom:this.verticalDefaultPosition=this.position,this.horizontalDefaultPosition="center";break;case P.right:case P.left:case P.start:case P.end:this.verticalDefaultPosition="center",this.horizontalDefaultPosition=this.position;break;case P.topLeft:this.verticalDefaultPosition="top",this.horizontalDefaultPosition="left";break;case P.topRight:this.verticalDefaultPosition="top",this.horizontalDefaultPosition="right";break;case P.bottomLeft:this.verticalDefaultPosition="bottom",this.horizontalDefaultPosition="left";break;case P.bottomRight:this.verticalDefaultPosition="bottom",this.horizontalDefaultPosition="right";break;case P.topStart:this.verticalDefaultPosition="top",this.horizontalDefaultPosition="start";break;case P.topEnd:this.verticalDefaultPosition="top",this.horizontalDefaultPosition="end";break;case P.bottomStart:this.verticalDefaultPosition="bottom",this.horizontalDefaultPosition="start";break;case P.bottomEnd:this.verticalDefaultPosition="bottom",this.horizontalDefaultPosition="end";break;default:this.verticalPositioningMode="dynamic",this.horizontalPositioningMode="dynamic",this.verticalDefaultPosition=void 0,this.horizontalDefaultPosition="center";break}}};l([u({mode:"boolean"})],v.prototype,"visible",void 0);l([u],v.prototype,"anchor",void 0);l([u],v.prototype,"delay",void 0);l([u],v.prototype,"position",void 0);l([u({attribute:"auto-update-mode"})],v.prototype,"autoUpdateMode",void 0);l([u({attribute:"horizontal-viewport-lock"})],v.prototype,"horizontalViewportLock",void 0);l([u({attribute:"vertical-viewport-lock"})],v.prototype,"verticalViewportLock",void 0);l([R],v.prototype,"anchorElement",void 0);l([R],v.prototype,"viewportElement",void 0);l([R],v.prototype,"verticalPositioningMode",void 0);l([R],v.prototype,"horizontalPositioningMode",void 0);l([R],v.prototype,"horizontalInset",void 0);l([R],v.prototype,"verticalInset",void 0);l([R],v.prototype,"horizontalScaling",void 0);l([R],v.prototype,"verticalScaling",void 0);l([R],v.prototype,"verticalDefaultPosition",void 0);l([R],v.prototype,"horizontalDefaultPosition",void 0);l([R],v.prototype,"tooltipVisible",void 0);l([R],v.prototype,"currentDirection",void 0);const gt=(s,e)=>O`
    ${K("block")} :host {
      box-sizing: content-box;
      height: 0;
      border: none;
      border-top: calc(${M} * 1px) solid ${ae};
    }

    :host([orientation="vertical"]) {
      border: none;
      height: 100%;
      margin: 0 calc(${Oe} * 1px);
      border-left: calc(${M} * 1px) solid ${ae};
  }
  `,ft=te.compose({baseName:"divider",template:ct,styles:gt}),mt=(s,e)=>O`
    ${K("block")} :host {
      --skeleton-fill-default: ${le};
      overflow: hidden;
      width: 100%;
      position: relative;
      background-color: var(--skeleton-fill, var(--skeleton-fill-default));
      --skeleton-animation-gradient-default: linear-gradient(
        270deg,
        var(--skeleton-fill, var(--skeleton-fill-default)) 0%,
        ${je} 51%,
        var(--skeleton-fill, var(--skeleton-fill-default)) 100%
      );
      --skeleton-animation-timing-default: ease-in-out;
    }

    :host(.rect) {
      border-radius: calc(${be} * 1px);
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

    ${K("block")} span.shimmer {
      position: absolute;
      width: 100%;
      height: 100%;
      background-image: var(--skeleton-animation-gradient, var(--skeleton-animation-gradient-default));
      background-size: 0px 0px / 90% 100%;
      background-repeat: no-repeat;
      background-color: var(--skeleton-animation-fill, ${le});
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
  `.withBehaviors(ye(O`
        :host{
          background-color: ${Ae.CanvasText};
        }
      `)),vt=A.compose({baseName:"skeleton",template:ut,styles:mt}),bt=(s,e)=>O`
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
      border-radius: calc(${be} * 1px);
      border: calc(${M} * 1px) solid ${W};
      background: ${Z};
      color: ${_e};
      padding: 4px 12px;
      height: fit-content;
      width: fit-content;
      ${Ue}
      white-space: nowrap;
      box-shadow: ${qe};
    }

    ${s.tagFor(a)} {
      display: flex;
      justify-content: center;
      align-items: center;
      overflow: visible;
      flex-direction: row;
    }

    ${s.tagFor(a)}.right,
    ${s.tagFor(a)}.left {
      flex-direction: column;
    }

    ${s.tagFor(a)}.top .tooltip::after,
    ${s.tagFor(a)}.bottom .tooltip::after,
    ${s.tagFor(a)}.left .tooltip::after,
    ${s.tagFor(a)}.right .tooltip::after {
      content: '';
      width: 12px;
      height: 12px;
      background: ${Z};
      border-top: calc(${M} * 1px) solid ${W};
      border-left: calc(${M} * 1px) solid ${W};
      position: absolute;
    }

    ${s.tagFor(a)}.top .tooltip::after {
      transform: translateX(-50%) rotate(225deg);
      bottom: 5px;
      left: 50%;
    }

    ${s.tagFor(a)}.top .tooltip {
      margin-bottom: 12px;
    }

    ${s.tagFor(a)}.bottom .tooltip::after {
      transform: translateX(-50%) rotate(45deg);
      top: 5px;
      left: 50%;
    }

    ${s.tagFor(a)}.bottom .tooltip {
      margin-top: 12px;
    }

    ${s.tagFor(a)}.left .tooltip::after {
      transform: translateY(-50%) rotate(135deg);
      top: 50%;
      right: 5px;
    }

    ${s.tagFor(a)}.left .tooltip {
      margin-right: 12px;
    }

    ${s.tagFor(a)}.right .tooltip::after {
      transform: translateY(-50%) rotate(-45deg);
      top: 50%;
      left: 5px;
    }

    ${s.tagFor(a)}.right .tooltip {
      margin-left: 12px;
    }
  `.withBehaviors(ye(O`
        :host([disabled]) {
          opacity: 1;
        }
        ${s.tagFor(a)}.top .tooltip::after,
        ${s.tagFor(a)}.bottom .tooltip::after,
        ${s.tagFor(a)}.left .tooltip::after,
        ${s.tagFor(a)}.right .tooltip::after {
          content: '';
          width: unset;
          height: unset;
        }
      `));class yt extends v{connectedCallback(){super.connectedCallback(),Z.setValueFor(this,Ve)}}const wt=yt.compose({baseName:"tooltip",baseClass:v,template:pt,styles:bt}),$={modified:"modified on",back:"Back",next:"Next",pages:"pages",page:"Page"},Rt=[He`
:host([hidden]){display:none}:host{display:block;font-family:var(--default-font-family, "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, "BlinkMacSystemFont", "Roboto", "Helvetica Neue", sans-serif);font-size:var(--default-font-size, 14px);--theme-primary-color:#0078d7;--theme-dark-color:#005a9e}:focus-visible{outline-color:var(--focus-ring-color,Highlight);outline-color:var(--focus-ring-color,-webkit-focus-ring-color);outline-style:var(--focus-ring-style,auto)}.ms-icon{display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;font-size:16px;-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;margin:4px 0}.error{background-color:#fde7e9;padding-block:8px;padding-inline:8px 12px}.ms-icon-chevron-down::before{content:"\\\e70d"}.ms-icon-chevron-up::before{content:"\\\e70e"}.ms-icon-contact::before{content:"\\\e77b"}.ms-icon-add-friend::before{content:"\\\e8fa"}.ms-icon-outlook-logo-inverser::before{content:"\\\eb6d"}.search-results{scroll-margin:100px}.search-result-grid{font-size:14px;margin:16px 4px;display:grid;grid-template-columns:32px 2fr 0fr;gap:.5rem}.search-result{font-size:14px;margin:16px 4px}.file-icon{--file-type-icon-height:32px}.search-result-info{margin:4px 0;display:inline-flex}.search-result-date{padding-top:3px}.search-result-date__shimmer{border-radius:4px;margin-top:2%;margin-left:5px;height:10px;width:200px}.search-result-summary{margin:4px 0;color:var(--body-subtle-text-color)}.search-result-thumbnail__shimmer{width:126px;height:72px}.search-result-thumbnail img{height:72px;max-width:126px;width:126px;object-fit:cover}.search-result-url{font-size:14px;font-weight:600;margin:4px 0}.search-result-url a{color:var(--color);text-decoration:none}.search-result-url a:hover{text-decoration:underline}.search-result-content__shimmer{border-radius:4px;margin-top:10px;height:10px}.search-result-name{font-size:16px;font-weight:600;margin:4px 0}.search-result-name__shimmer{border-radius:4px;margin-top:10px;height:10px;width:20%}.search-result-name a{color:var(--color);text-decoration:none}.search-result-name a:hover{text-decoration:underline}.search-result-author__shimmer{width:24px;height:24px}.search-result-icon{width:30px;height:30px}.search-result-icon__shimmer{width:32px;height:32px}.search-result-icon img{width:100%}.search-result-icon svg,.search-result-icon svg>path{width:100%;height:100%;fill:var(--neutral-foreground-rest);fill-rule:nonzero!important;clip-rule:nonzero!important}.search-results-page-active{border-bottom-style:solid;border-bottom-color:var(--accent-base-color)}.search-results-page svg,.search-results-page svg>path{height:12px;fill:var(--neutral-foreground-rest);fill-rule:nonzero!important;clip-rule:nonzero!important}.search-result-answer{box-shadow:var(--answer-box-shadow,0 3.2px 7.2px rgba(0,0,0,.132),0 .6px 1.8px rgba(0,0,0,.108));border-radius:var(--answer-border-radius,4px);border:var(--answer-border,none);padding:var(--answer-padding,10px)}.search-results-pages{margin:16px 4px}
`];var x=function(s,e,t,i){var r=arguments.length,o=r<3?e:i===null?i=Object.getOwnPropertyDescriptor(e,t):i,h;if(typeof Reflect=="object"&&typeof Reflect.decorate=="function")o=Reflect.decorate(s,e,t,i);else for(var c=s.length-1;c>=0;c--)(h=s[c])&&(o=(r<3?h(o):r>3?h(e,t,o):h(e,t))||o);return r>3&&o&&Object.defineProperty(e,t,o),o},y=function(s,e){if(typeof Reflect=="object"&&typeof Reflect.metadata=="function")return Reflect.metadata(s,e)},xt=function(s,e,t,i){function r(o){return o instanceof t?o:new t(function(h){h(o)})}return new(t||(t=Promise))(function(o,h){function c(m){try{g(i.next(m))}catch(p){h(p)}}function d(m){try{g(i.throw(m))}catch(p){h(p)}}function g(m){m.done?o(m.value):r(m.value).then(c,d)}g((i=i.apply(s,e||[])).next())})};const Pt=()=>{Ne(vt,Je,wt,ft),Pe(),Be(),We("search-results",w)};class w extends Xe{static get styles(){return Rt}get strings(){return $}get queryString(){return this._queryString}set queryString(e){this._queryString!==e&&(this._queryString=e,this.currentPage=1)}get from(){return(this.currentPage-1)*this.size}get size(){return this._size}set size(e){e>this.maxPageSize?this._size=this.maxPageSize:this._size=e}get searchEndpoint(){return"/search/query"}get maxPageSize(){return 1e3}constructor(){super(),this._size=10,this.entityTypes=["driveItem","listItem","site"],this.scopes=[],this.contentSources=[],this.version="v1.0",this.pagingMax=7,this.enableTopResults=!1,this.cacheEnabled=!1,this.cacheInvalidationPeriod=3e4,this.isRefreshing=!1,this.defaultFields=["webUrl","lastModifiedBy","lastModifiedDateTime","summary","displayName","name"],this.currentPage=1,this.renderContent=()=>{var e,t,i,r,o,h,c,d;let g=null,m=null,p=null;return this.hasTemplate("header")&&(m=this.renderTemplate("header",this.response)),p=this.renderFooter((t=(e=this.response)===null||e===void 0?void 0:e.value[0])===null||t===void 0?void 0:t.hitsContainers[0]),this.response&&this.hasTemplate("default")?g=this.renderTemplate("default",this.response)||f``:!((r=(i=this.response)===null||i===void 0?void 0:i.value[0])===null||r===void 0)&&r.hitsContainers[0]?g=f`${(d=(c=(h=(o=this.response)===null||o===void 0?void 0:o.value[0])===null||h===void 0?void 0:h.hitsContainers[0])===null||c===void 0?void 0:c.hits)===null||d===void 0?void 0:d.map(b=>this.renderResult(b))}`:this.hasTemplate("no-data")?g=this.renderTemplate("no-data",null):g=f``,f`
      ${m}
      <div class="search-results">
        ${g}
      </div>
      ${p}`},this.renderLoading=()=>this.renderTemplate("loading",null)||f`
        ${[...Array(this.size)].map(()=>f`
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
                ${this.fetchThumbnail&&f`
                    <div class="search-result-thumbnail">
                      <fluent-skeleton class="search-result-thumbnail__shimmer" shape="rect" shimmer></fluent-skeleton>
                    </div>
                  `}
              </div>
              <fluent-divider></fluent-divider>
            </div>
          `)}
       `,this.onFirstPageClick=()=>{this.currentPage=1,this.scrollToFirstResult()},this.onPageBackClick=()=>{this.currentPage--,this.scrollToFirstResult()},this.onPageNextClick=()=>{this.currentPage++,this.scrollToFirstResult()}}args(){return[this.providerState,this.queryString,this.queryTemplate,this.entityTypes,this.contentSources,this.scopes,this.version,this.size,this.fetchThumbnail,this.fields,this.enableTopResults,this.currentPage]}refresh(e=!1){this.isRefreshing=!0,e&&this.clearState(),this._task.run()}clearState(){this.response=null}loadState(){var e,t,i,r,o,h;return xt(this,void 0,void 0,function*(){const c=Ce.globalProvider;if(this.error=null,!(!c||c.state!==Se.SignedIn)){if(this.queryString){try{const d=this.getRequestOptions();let g;const m=JSON.stringify({endpoint:`${this.version}${this.searchEndpoint}`,requestOptions:d});let p=null;if(this.shouldRetrieveCache()){g=he.getCache(_.search,_.search.stores.responses);const b=X()?yield g.getValue(m):null;b&&Ye(this.cacheInvalidationPeriod)>Date.now()-b.timeCached&&(p=JSON.parse(b.response))}if(!p){const b=c.graph.forComponent(this);let S=b.api(this.searchEndpoint).version(this.version);if(!((e=this.scopes)===null||e===void 0)&&e.length&&(S=S.middlewareOptions(Ee(this.scopes))),p=yield S.post({requests:[d]}),this.fetchThumbnail){const ie=b.createBatch(),se=N.fromGraph(b).createBatch(),De=!((t=p.value)===null||t===void 0)&&t.length&&(!((i=p.value[0].hitsContainers)===null||i===void 0)&&i.length)?(o=(r=p.value[0].hitsContainers[0])===null||r===void 0?void 0:r.hits)!==null&&o!==void 0?o:[]:[];for(const C of De){const D=C.resource;(D.size>0||!((h=D.webUrl)===null||h===void 0)&&h.endsWith(".aspx"))&&(D["@odata.type"]==="#microsoft.graph.driveItem"||D["@odata.type"]==="#microsoft.graph.listItem")&&(D["@odata.type"]==="#microsoft.graph.listItem"?se.get(C.hitId.toString(),`/sites/${D.parentReference.siteId}/pages/${D.id}`):ie.get(C.hitId.toString(),`/drives/${D.parentReference.driveId}/items/${D.id}/thumbnails/0/medium`))}const re=C=>{if(C&&C.size>0)for(const[D,oe]of C){const ne=p.value[0].hitsContainers[0].hits[D],Te=ne.resource["@odata.type"]==="#microsoft.graph.listItem"?{url:oe.content.thumbnailWebUrl}:{url:oe.content.url};ne.resource.thumbnail=Te}};try{re(yield ie.executeAll()),re(yield se.executeAll())}catch{}}this.shouldUpdateCache()&&p&&(g=he.getCache(_.search,_.search.stores.responses),yield g.putValue(m,{response:JSON.stringify(p)}))}Ge(this.response,p)||(this.response=p)}catch(d){this.error=d}this.response&&(this.error=null)}else this.response=null;this.isRefreshing=!1,this.fireCustomEvent("dataChange",{response:this.response,error:this.error})}})}renderResult(e){const t=this.getResourceType(e.resource);if(this.hasTemplate(`result-${t}`))return this.renderTemplate(`result-${t}`,e,e.hitId);switch(e.resource["@odata.type"]){case"#microsoft.graph.driveItem":return this.renderDriveItem(e);case"#microsoft.graph.site":return this.renderSite(e);case"#microsoft.graph.person":return this.renderPerson(e);case"#microsoft.graph.drive":case"#microsoft.graph.list":return this.renderList(e);case"#microsoft.graph.listItem":return this.renderListItem(e);case"#microsoft.graph.search.bookmark":return this.renderBookmark(e);case"#microsoft.graph.search.acronym":return this.renderAcronym(e);case"#microsoft.graph.search.qna":return this.renderQnA(e);default:return this.renderDefault(e)}}renderFooter(e){if(this.pagingRequired(e)){const t=this.getActivePages(e.total);return f`
        <div class="search-results-pages">
          ${this.renderPreviousPage()}
          ${this.renderFirstPage(t)}
          ${this.renderAllPages(t)}
          ${this.renderNextPage()}
        </div>
      `}}pagingRequired(e){return(e==null?void 0:e.moreResultsAvailable)||this.currentPage*this.size<(e==null?void 0:e.total)}getActivePages(e){const t=()=>{const o=this.currentPage-Math.floor(this.pagingMax/2)-1;return o>=Math.floor(this.pagingMax/2)?o:0},i=[],r=t();if(r+1>this.pagingMax-this.currentPage||this.pagingMax===this.currentPage)for(let o=r+1;o<Math.ceil(e/this.size)&&o<this.pagingMax+(this.currentPage-1)&&i.length<this.pagingMax-2;++o)i.push(o+1);else for(let o=r;o<this.pagingMax;++o)i.push(o+1);return i}renderAllPages(e){return f`
      ${e.map(t=>f`
            <fluent-button
              title="${$.page} ${t}"
              appearance="stealth"
              class="${t===this.currentPage?"search-results-page-active":"search-results-page"}"
              @click="${()=>this.onPageClick(t)}">
                ${t}
            </fluent-button>`)}`}renderFirstPage(e){return f`
      ${e.some(t=>t===1)?E:f`
              <fluent-button
                 title="${$.page} 1"
                 appearance="stealth"
                 class="search-results-page"
                 @click="${this.onFirstPageClick}">
                 1
               </fluent-button>`?f`
              <fluent-button
                id="page-back-dot"
                appearance="stealth"
                class="search-results-page"
                title="${this.getDotButtonTitle()}"
                @click="${()=>this.onPageClick(this.currentPage-Math.ceil(this.pagingMax/2))}"
              >
                ...
              </fluent-button>`:E}`}getDotButtonTitle(){return`${$.back} ${Math.ceil(this.pagingMax/2)} ${$.pages}`}renderPreviousPage(){return this.currentPage>1?f`
          <fluent-button
            appearance="stealth"
            class="search-results-page"
            title="${$.back}"
            @click="${this.onPageBackClick}">
              ${T(z.ChevronLeft)}
            </fluent-button>`:E}renderNextPage(){return this.isLastPage()?E:f`
          <fluent-button
            appearance="stealth"
            class="search-results-page"
            title="${$.next}"
            aria-label="${$.next}"
            @click="${this.onPageNextClick}">
              ${T(z.ChevronRight)}
            </fluent-button>`}onPageClick(e){this.currentPage=e,this.scrollToFirstResult()}isLastPage(){return this.currentPage===Math.ceil(this.response.value[0].hitsContainers[0].total/this.size)}scrollToFirstResult(){this.renderRoot.querySelector(".search-results").scrollIntoView({block:"start",behavior:"smooth"})}getResourceType(e){return e["@odata.type"].split(".").pop()}renderDriveItem(e){var t,i;const r=e.resource;return U`
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
            <a href="${r.webUrl}?Web=1" target="_blank">${L(r.name)}</a>
          </div>
          <div class="search-result-info">
            <div class="search-result-author">
              <mgt-person
                person-query=${r.lastModifiedBy.user.email}
                view="oneline"
                person-card="hover"
                show-presence="true">
              </mgt-person>
            </div>
            <div class="search-result-date">
              &nbsp; ${$.modified} ${ce(new Date(r.lastModifiedDateTime))}
            </div>
          </div>
          <div class="search-result-summary" .innerHTML="${I(e.summary)}"></div>
        </div>
        ${((t=r.thumbnail)===null||t===void 0?void 0:t.url)&&f`
          <div class="search-result-thumbnail">
            <a href="${r.webUrl}" target="_blank"><img alt="${r.name}" src="${(i=r.thumbnail)===null||i===void 0?void 0:i.url}" /></a>
          </div>`}

      </div>
      <fluent-divider></fluent-divider>
    `}renderSite(e){const t=e.resource;return f`
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
          <div class="search-result-summary" .innerHTML="${I(e.summary)}"></div>
        </div>
      </div>
      <fluent-divider></fluent-divider>
    `}renderList(e){const t=e.resource;return U`
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
              ${L(t.name||Y(t.webUrl))}
            </a>
          </div>
          <div class="search-result-summary" .innerHTML="${I(e.summary)}"></div>
        </div>
      </div>
      <fluent-divider></fluent-divider>
    `}renderListItem(e){var t,i;const r=e.resource;return U`
      <div class="search-result-grid">
        <div class="search-result-icon">
          ${r.webUrl.endsWith(".aspx")?T(z.News):T(z.FileOuter)}
        </div>
        <div class="search-result-content">
          <div class="search-result-name">
            <a href="${r.webUrl}?Web=1" target="_blank">
              ${L(r.name||Y(r.webUrl))}
            </a>
          </div>
          <div class="search-result-info">
            <div class="search-result-author">
              <mgt-person
                person-query=${r.lastModifiedBy.user.email}
                view="oneline"
                person-card="hover"
                show-presence="true">
              </mgt-person>
            </div>
            <div class="search-result-date">
              &nbsp; ${$.modified} ${ce(new Date(r.lastModifiedDateTime))}
            </div>
          </div>
          <div class="search-result-summary" .innerHTML="${I(e.summary)}"></div>
        </div>
        ${((t=r.thumbnail)===null||t===void 0?void 0:t.url)&&f`
          <div class="search-result-thumbnail">
            <a href="${r.webUrl}" target="_blank"><img alt="${L(r.name||Y(r.webUrl))}" src="${((i=r.thumbnail)===null||i===void 0?void 0:i.url)||E}" /></a>
          </div>`}
      </div>
      <fluent-divider></fluent-divider>
    `}renderPerson(e){const t=e.resource;return U`
      <div class="search-result">
        <mgt-person
          view="fourlines"
          person-query=${t.userPrincipalName}
          person-card="hover"
          show-presence="true">
        </mgt-person>
      </div>
      <fluent-divider></fluent-divider>
    `}renderBookmark(e){return this.renderAnswer(e,z.DoubleBookmark)}renderAcronym(e){return this.renderAnswer(e,z.BookOpen)}renderQnA(e){return this.renderAnswer(e,z.BookQuestion)}renderAnswer(e,t){const i=e.resource;return f`
      <div class="search-result-grid search-result-answer">
        <div class="search-result-icon">
          ${T(t)}
        </div>
        <div class="search-result-content">
          <div class="search-result-name">
            <a href="${this.getResourceUrl(i)}?Web=1" target="_blank">${i.displayName}</a>
          </div>
          <div class="search-result-summary">${i.description}</div>
        </div>
      </div>
      <fluent-divider></fluent-divider>
    `}renderDefault(e){const t=e.resource,i=this.getResourceUrl(t);return f`
      <div class="search-result-grid">
        <div class="search-result-icon">
          ${this.getResourceIcon(t)}
        </div>
        <div class="search-result-content">
          <div class="search-result-name">
            ${i?f`
                  <a href="${i}?Web=1" target="_blank">${this.getResourceName(t)}</a>
                `:f`
                  ${this.getResourceName(t)}
                `}
          </div>
          <div class="search-result-summary" .innerHTML="${this.getResultSummary(e)}"></div>
        </div>
      </div>
      <fluent-divider></fluent-divider>
    `}getResourceUrl(e){return e.webUrl||e.webLink||null}getResourceName(e){return e.displayName||e.subject||L(e.name)}getResultSummary(e){var t;return I(e.summary||((t=e.resource)===null||t===void 0?void 0:t.description)||null)}getResourceIcon(e){switch(e["@odata.type"]){case"#microsoft.graph.site":return T(z.Globe);case"#microsoft.graph.message":return T(z.Email);case"#microsoft.graph.event":return T(z.Event);case"microsoft.graph.chatMessage":return T(z.SmallChat);default:return T(z.FileOuter)}}shouldRetrieveCache(){return X()&&this.cacheEnabled&&!this.isRefreshing}shouldUpdateCache(){return X()&&this.cacheEnabled}getRequestOptions(){const e={entityTypes:this.entityTypes,query:{queryString:this.queryString},from:this.from?this.from:void 0,size:this.size?this.size:void 0,fields:this.getFields(),enableTopResults:this.enableTopResults?this.enableTopResults:void 0};return this.entityTypes.includes("externalItem")&&(e.contentSources=this.contentSources),this.version==="beta"&&(e.query.queryTemplate=this.queryTemplate?this.queryTemplate:void 0),e}getFields(){if(this.fields)return this.defaultFields.concat(this.fields)}}x([k({attribute:"query-string",type:String}),y("design:type",String),y("design:paramtypes",[String])],w.prototype,"queryString",null);x([k({attribute:"query-template",type:String}),y("design:type",String)],w.prototype,"queryTemplate",void 0);x([k({attribute:"entity-types",converter:s=>s.split(",").map(e=>e.trim()),type:String}),y("design:type",Array)],w.prototype,"entityTypes",void 0);x([k({attribute:"scopes",converter:(s,e)=>s?s.toLowerCase().split(","):null}),y("design:type",Array)],w.prototype,"scopes",void 0);x([k({attribute:"content-sources",converter:(s,e)=>s?s.toLowerCase().split(","):null}),y("design:type",Array)],w.prototype,"contentSources",void 0);x([k({attribute:"version",reflect:!0,type:String}),y("design:type",Object)],w.prototype,"version",void 0);x([k({attribute:"size",reflect:!0,type:Number}),y("design:type",Number),y("design:paramtypes",[Object])],w.prototype,"size",null);x([k({attribute:"paging-max",reflect:!0,type:Number}),y("design:type",Object)],w.prototype,"pagingMax",void 0);x([k({attribute:"fetch-thumbnail",type:Boolean}),y("design:type",Boolean)],w.prototype,"fetchThumbnail",void 0);x([k({attribute:"fields",converter:s=>s.split(",").map(e=>e.trim()),type:String}),y("design:type",Array)],w.prototype,"fields",void 0);x([k({attribute:"enable-top-results",reflect:!0,type:Boolean}),y("design:type",Object)],w.prototype,"enableTopResults",void 0);x([k({attribute:"cache-enabled",reflect:!0,type:Boolean}),y("design:type",Object)],w.prototype,"cacheEnabled",void 0);x([k({attribute:"cache-invalidation-period",reflect:!0,type:Number}),y("design:type",Object)],w.prototype,"cacheInvalidationPeriod",void 0);x([Qe(),y("design:type",Object)],w.prototype,"response",void 0);x([k({attribute:!1}),y("design:type",Object)],w.prototype,"currentPage",void 0);const kt=we("file",Pe),j=we("search-results",Pt),zt=s=>n.jsx(n.Fragment,{children:s.searchTerm&&n.jsxs(n.Fragment,{children:[s.searchTerm!=="*"&&n.jsx(j,{entityTypes:["bookmark"],queryString:s.searchTerm,version:"beta",size:1,scopes:["Bookmark.Read.All"],children:n.jsx(Dt,{template:"no-data"})}),n.jsx(j,{entityTypes:["driveItem","listItem","site"],queryString:s.searchTerm,scopes:["Files.Read.All","Files.ReadWrite.All","Sites.Read.All"],fetchThumbnail:!0})]})}),Dt=s=>n.jsx(n.Fragment,{}),Tt=s=>n.jsx(n.Fragment,{children:n.jsx(j,{entityTypes:["people"],size:20,queryString:s.searchTerm,version:"beta"})}),$t=s=>n.jsx(n.Fragment,{children:s.searchTerm&&n.jsx(j,{entityTypes:["externalItem"],contentSources:["/external/connections/contosoBlogPosts"],queryString:s.searchTerm,scopes:["ExternalItem.Read.All"],version:"beta"})}),ke=Re({container:{...xe.gap("16px"),display:"flex",flexDirection:"row",flexWrap:"wrap"},card:{width:"300px",height:"fit-content",maxWidth:"100%"},caption:{color:de.colorNeutralForeground3},noDataSearchTerm:{fontWeight:de.fontWeightSemibold},emptyContainer:{display:"flex",justifyContent:"center",alignItems:"center",flexDirection:"column",height:"calc(100vh - 300px)"},fileContainer:{display:"flex"},fileTitle:{paddingLeft:"10px",alignSelf:"center"},noDataMessage:{paddingLeft:"10px"},noDataIcon:{fontSize:"128px"},row:{cursor:"pointer"}}),ze=30,Ct=s=>n.jsx(n.Fragment,{children:s.searchTerm&&n.jsxs(j,{entityTypes:["driveItem"],queryString:s.searchTerm,fetchThumbnail:!0,queryTemplate:"({searchTerms}) ContentTypeId:0x0101*",version:"beta",fields:["createdBy","lastModifiedDateTime","Title","DefaultEncodingURL"],size:ze,cacheEnabled:!0,children:[n.jsx(fe,{template:"default"}),n.jsx(fe,{template:"loading"}),n.jsx(Et,{template:"no-data"})]})}),St=(s,e)=>[G({columnId:"name",renderHeaderCell:()=>"Name",renderCell:i=>n.jsx(Q,{children:s?n.jsx(V,{shape:"rectangle",style:{width:"120px"}}):n.jsxs("div",{className:e.fileContainer,children:[n.jsx(kt,{fileDetails:i.resource,view:"image"}),n.jsx("span",{className:e.fileTitle,children:i.resource.listItem.fields.title})]})})}),G({columnId:"modified",renderHeaderCell:()=>"Modified",renderCell:i=>n.jsx(Q,{children:s?n.jsx(V,{shape:"rectangle",style:{width:"120px"}}):Lt(new Date(i.resource.lastModifiedDateTime))})}),G({columnId:"owner",renderHeaderCell:()=>"Owner",renderCell:i=>n.jsx(Q,{children:s?n.jsxs("div",{style:{display:"grid",alignItems:"center",position:"relative",gridTemplateColumns:"min-content 80%",gap:"10px"},children:[n.jsx(V,{shape:"circle",size:32}),n.jsx(V,{style:{width:"120px"}})]}):n.jsx(Ze,{personQuery:i.resource.createdBy.user.email,view:"oneline",personCardInteraction:"hover"})})})],fe=s=>{var r,o,h;const e=ke(),[t]=H.useState((h=(o=(r=s.dataContext.value)==null?void 0:r[0])==null?void 0:o.hitsContainers[0])==null?void 0:h.hits),i=c=>{const d=new URL(c.resource.listItem.fields.defaultEncodingURL);d.searchParams.append("Web","1"),window.open(d.toString(),"_blank")};return n.jsx("div",{children:n.jsxs(st,{columns:St(s.template==="loading",e),items:s.template==="loading"?[...Array(ze)]:t,children:[n.jsx(rt,{children:n.jsx(ge,{children:({renderHeaderCell:c})=>n.jsx(ot,{children:c()})})}),n.jsx(nt,{children:({item:c,rowId:d})=>n.jsx(ge,{className:e.row,onClick:()=>i(c),children:({renderCell:g})=>n.jsx(at,{children:g(c)})},d)})]})})},Et=s=>{var i;const e=ke(),[t]=H.useState((i=s.dataContext.value[0])==null?void 0:i.searchTerms);return n.jsxs("div",{className:e.emptyContainer,children:[n.jsx("div",{children:n.jsx(Ke,{className:e.noDataIcon})}),n.jsxs("div",{className:e.noDataMessage,children:["We couldn't find any results for ",n.jsx("span",{className:e.noDataSearchTerm,children:t.join(" ")})]})]})},Lt=s=>{const e=new Date,t=new Date(e.getFullYear(),e.getMonth(),e.getDate());if(s>=t)return s.toLocaleString("default",{hour:"numeric",minute:"numeric"});const i=new Date(t);if(i.setDate(e.getDate()-e.getDay()),s>=i)return s.toLocaleString("default",{hour:"numeric",minute:"numeric",weekday:"short"});const r=new Date(i);return r.setDate(i.getDate()-7),s>=r?s.toLocaleString("default",{day:"numeric",month:"numeric",weekday:"short"}):s.toLocaleString("default",{day:"numeric",month:"numeric",year:"numeric"})},It=Re({panels:{...xe.padding("10px")},container:{maxWidth:"1028px",width:"100%"}}),Vt=()=>{const s=It(),e=et(),[t]=H.useState(new URLSearchParams(window.location.search).get("q")),[i,r]=H.useState("allResults"),o=(h,c)=>{r(c.value)};return n.jsxs(n.Fragment,{children:[n.jsx(Le,{title:"Search",description:"Use this Search Center to test Microsot Graph Toolkit search components capabilities"}),n.jsxs("div",{className:s.container,children:[n.jsxs(tt,{selectedValue:i,onTabSelect:o,children:[n.jsx(q,{value:"allResults",children:"All Results"}),n.jsx(q,{value:"driveItems",children:"Files"}),n.jsx(q,{value:"externalItems",children:"External Items"}),n.jsx(q,{value:"people",children:"People"})]}),n.jsxs("div",{className:s.panels,children:[i==="allResults"&&n.jsx(zt,{searchTerm:t??e.state.searchTerm}),i==="driveItems"&&n.jsx(Ct,{searchTerm:t??e.state.searchTerm}),i==="externalItems"&&n.jsx($t,{searchTerm:t??e.state.searchTerm}),i==="people"&&n.jsx(Tt,{searchTerm:t??e.state.searchTerm})]})]})]})};export{Vt as default};
