/*
 * (c) Copyright Ascensio System SIA 2010-2019
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

// 3rdParty
/*!
 * perfect-scrollbar v1.4.0
 * (c) 2018 Hyunje Jun
 * @license MIT
 */
!function(t,e){"object"==typeof exports&&"undefined"!=typeof module?module.exports=e():"function"==typeof define&&define.amd?define(e):t.PerfectScrollbar=e()}(this,function(){"use strict";function t(t){return getComputedStyle(t)}function e(t,e){for(var i in e){var r=e[i];"number"==typeof r&&(r+="px"),t.style[i]=r}return t}function i(t){var e=document.createElement("div");return e.className=t,e}function r(t,e){if(!v)throw new Error("No element matching method supported");return v.call(t,e)}function l(t){t.remove?t.remove():t.parentNode&&t.parentNode.removeChild(t)}function n(t,e){return Array.prototype.filter.call(t.children,function(t){return r(t,e)})}function o(t,e){var i=t.element.classList,r=m.state.scrolling(e);i.contains(r)?clearTimeout(Y[e]):i.add(r)}function s(t,e){Y[e]=setTimeout(function(){return t.isAlive&&t.element.classList.remove(m.state.scrolling(e))},t.settings.scrollingThreshold)}function a(t,e){o(t,e),s(t,e)}function c(t){if("function"==typeof window.CustomEvent)return new CustomEvent(t);var e=document.createEvent("CustomEvent");return e.initCustomEvent(t,!1,!1,void 0),e}function h(t,e,i,r,l){var n=i[0],o=i[1],s=i[2],h=i[3],u=i[4],d=i[5];void 0===r&&(r=!0),void 0===l&&(l=!1);var f=t.element;t.reach[h]=null,f[s]<1&&(t.reach[h]="start"),f[s]>t[n]-t[o]-1&&(t.reach[h]="end"),e&&(f.dispatchEvent(c("ps-scroll-"+h)),e<0?f.dispatchEvent(c("ps-scroll-"+u)):e>0&&f.dispatchEvent(c("ps-scroll-"+d)),r&&a(t,h)),t.reach[h]&&(e||l)&&f.dispatchEvent(c("ps-"+h+"-reach-"+t.reach[h]))}function u(t){return parseInt(t,10)||0}function d(t){return r(t,"input,[contenteditable]")||r(t,"select,[contenteditable]")||r(t,"textarea,[contenteditable]")||r(t,"button,[contenteditable]")}function f(e){var i=t(e);return u(i.width)+u(i.paddingLeft)+u(i.paddingRight)+u(i.borderLeftWidth)+u(i.borderRightWidth)}function p(t,e){return t.settings.minScrollbarLength&&(e=Math.max(e,t.settings.minScrollbarLength)),t.settings.maxScrollbarLength&&(e=Math.min(e,t.settings.maxScrollbarLength)),e}function b(t,i){var r={width:i.railXWidth},l=Math.floor(t.scrollTop);i.isRtl?r.left=i.negativeScrollAdjustment+t.scrollLeft+i.containerWidth-i.contentWidth:r.left=t.scrollLeft,i.isScrollbarXUsingBottom?r.bottom=i.scrollbarXBottom-l:r.top=i.scrollbarXTop+l,e(i.scrollbarXRail,r);var n={top:l,height:i.railYHeight};i.isScrollbarYUsingRight?i.isRtl?n.right=i.contentWidth-(i.negativeScrollAdjustment+t.scrollLeft)-i.scrollbarYRight-i.scrollbarYOuterWidth:n.right=i.scrollbarYRight-t.scrollLeft:i.isRtl?n.left=i.negativeScrollAdjustment+t.scrollLeft+2*i.containerWidth-i.contentWidth-i.scrollbarYLeft-i.scrollbarYOuterWidth:n.left=i.scrollbarYLeft+t.scrollLeft,e(i.scrollbarYRail,n),e(i.scrollbarX,{left:i.scrollbarXLeft,width:i.scrollbarXWidth-i.railBorderXWidth}),e(i.scrollbarY,{top:i.scrollbarYTop,height:i.scrollbarYHeight-i.railBorderYWidth})}function g(t,e){function i(e){b[d]=g+Y*(e[a]-v),o(t,f),R(t),e.stopPropagation(),e.preventDefault()}function r(){s(t,f),t[p].classList.remove(m.state.clicking),t.event.unbind(t.ownerDocument,"mousemove",i)}var l=e[0],n=e[1],a=e[2],c=e[3],h=e[4],u=e[5],d=e[6],f=e[7],p=e[8],b=t.element,g=null,v=null,Y=null;t.event.bind(t[h],"mousedown",function(e){g=b[d],v=e[a],Y=(t[n]-t[l])/(t[c]-t[u]),t.event.bind(t.ownerDocument,"mousemove",i),t.event.once(t.ownerDocument,"mouseup",r),t[p].classList.add(m.state.clicking),e.stopPropagation(),e.preventDefault()})}var v="undefined"!=typeof Element&&(Element.prototype.matches||Element.prototype.webkitMatchesSelector||Element.prototype.mozMatchesSelector||Element.prototype.msMatchesSelector),m={main:"ps",element:{thumb:function(t){return"ps__thumb-"+t},rail:function(t){return"ps__rail-"+t},consuming:"ps__child--consume"},state:{focus:"ps--focus",clicking:"ps--clicking",active:function(t){return"ps--active-"+t},scrolling:function(t){return"ps--scrolling-"+t}}},Y={x:null,y:null},X=function(t){this.element=t,this.handlers={}},w={isEmpty:{configurable:!0}};X.prototype.bind=function(t,e){void 0===this.handlers[t]&&(this.handlers[t]=[]),this.handlers[t].push(e),this.element.addEventListener(t,e,!1)},X.prototype.unbind=function(t,e){var i=this;this.handlers[t]=this.handlers[t].filter(function(r){return!(!e||r===e)||(i.element.removeEventListener(t,r,!1),!1)})},X.prototype.unbindAll=function(){var t=this;for(var e in t.handlers)t.unbind(e)},w.isEmpty.get=function(){var t=this;return Object.keys(this.handlers).every(function(e){return 0===t.handlers[e].length})},Object.defineProperties(X.prototype,w);var y=function(){this.eventElements=[]};y.prototype.eventElement=function(t){var e=this.eventElements.filter(function(e){return e.element===t})[0];return e||(e=new X(t),this.eventElements.push(e)),e},y.prototype.bind=function(t,e,i){this.eventElement(t).bind(e,i)},y.prototype.unbind=function(t,e,i){var r=this.eventElement(t);r.unbind(e,i),r.isEmpty&&this.eventElements.splice(this.eventElements.indexOf(r),1)},y.prototype.unbindAll=function(){this.eventElements.forEach(function(t){return t.unbindAll()}),this.eventElements=[]},y.prototype.once=function(t,e,i){var r=this.eventElement(t),l=function(t){r.unbind(e,l),i(t)};r.bind(e,l)};var W=function(t,e,i,r,l){void 0===r&&(r=!0),void 0===l&&(l=!1);var n;if("top"===e)n=["contentHeight","containerHeight","scrollTop","y","up","down"];else{if("left"!==e)throw new Error("A proper axis should be provided");n=["contentWidth","containerWidth","scrollLeft","x","left","right"]}h(t,i,n,r,l)},L={isWebKit:"undefined"!=typeof document&&"WebkitAppearance"in document.documentElement.style,supportsTouch:"undefined"!=typeof window&&("ontouchstart"in window||window.DocumentTouch&&document instanceof window.DocumentTouch),supportsIePointer:"undefined"!=typeof navigator&&navigator.msMaxTouchPoints,isChrome:"undefined"!=typeof navigator&&/Chrome/i.test(navigator&&navigator.userAgent)},R=function(t){var e=t.element,i=Math.floor(e.scrollTop);t.containerWidth=e.clientWidth,t.containerHeight=e.clientHeight,t.contentWidth=e.scrollWidth,t.contentHeight=e.scrollHeight,e.contains(t.scrollbarXRail)||(n(e,m.element.rail("x")).forEach(function(t){return l(t)}),e.appendChild(t.scrollbarXRail)),e.contains(t.scrollbarYRail)||(n(e,m.element.rail("y")).forEach(function(t){return l(t)}),e.appendChild(t.scrollbarYRail)),!t.settings.suppressScrollX&&t.containerWidth+t.settings.scrollXMarginOffset<t.contentWidth?(t.scrollbarXActive=!0,t.railXWidth=t.containerWidth-t.railXMarginWidth,t.railXRatio=t.containerWidth/t.railXWidth,t.scrollbarXWidth=p(t,u(t.railXWidth*t.containerWidth/t.contentWidth)),t.scrollbarXLeft=u((t.negativeScrollAdjustment+e.scrollLeft)*(t.railXWidth-t.scrollbarXWidth)/(t.contentWidth-t.containerWidth))):t.scrollbarXActive=!1,!t.settings.suppressScrollY&&t.containerHeight+t.settings.scrollYMarginOffset<t.contentHeight?(t.scrollbarYActive=!0,t.railYHeight=t.containerHeight-t.railYMarginHeight,t.railYRatio=t.containerHeight/t.railYHeight,t.scrollbarYHeight=p(t,u(t.railYHeight*t.containerHeight/t.contentHeight)),t.scrollbarYTop=u(i*(t.railYHeight-t.scrollbarYHeight)/(t.contentHeight-t.containerHeight))):t.scrollbarYActive=!1,t.scrollbarXLeft>=t.railXWidth-t.scrollbarXWidth&&(t.scrollbarXLeft=t.railXWidth-t.scrollbarXWidth),t.scrollbarYTop>=t.railYHeight-t.scrollbarYHeight&&(t.scrollbarYTop=t.railYHeight-t.scrollbarYHeight),b(e,t),t.scrollbarXActive?e.classList.add(m.state.active("x")):(e.classList.remove(m.state.active("x")),t.scrollbarXWidth=0,t.scrollbarXLeft=0,e.scrollLeft=0),t.scrollbarYActive?e.classList.add(m.state.active("y")):(e.classList.remove(m.state.active("y")),t.scrollbarYHeight=0,t.scrollbarYTop=0,e.scrollTop=0)},T={"click-rail":function(t){t.event.bind(t.scrollbarY,"mousedown",function(t){return t.stopPropagation()}),t.event.bind(t.scrollbarYRail,"mousedown",function(e){var i=e.pageY-window.pageYOffset-t.scrollbarYRail.getBoundingClientRect().top>t.scrollbarYTop?1:-1;t.element.scrollTop+=i*t.containerHeight,R(t),e.stopPropagation()}),t.event.bind(t.scrollbarX,"mousedown",function(t){return t.stopPropagation()}),t.event.bind(t.scrollbarXRail,"mousedown",function(e){var i=e.pageX-window.pageXOffset-t.scrollbarXRail.getBoundingClientRect().left>t.scrollbarXLeft?1:-1;t.element.scrollLeft+=i*t.containerWidth,R(t),e.stopPropagation()})},"drag-thumb":function(t){g(t,["containerWidth","contentWidth","pageX","railXWidth","scrollbarX","scrollbarXWidth","scrollLeft","x","scrollbarXRail"]),g(t,["containerHeight","contentHeight","pageY","railYHeight","scrollbarY","scrollbarYHeight","scrollTop","y","scrollbarYRail"])},keyboard:function(t){function e(e,r){var l=Math.floor(i.scrollTop);if(0===e){if(!t.scrollbarYActive)return!1;if(0===l&&r>0||l>=t.contentHeight-t.containerHeight&&r<0)return!t.settings.wheelPropagation}var n=i.scrollLeft;if(0===r){if(!t.scrollbarXActive)return!1;if(0===n&&e<0||n>=t.contentWidth-t.containerWidth&&e>0)return!t.settings.wheelPropagation}return!0}var i=t.element,l=function(){return r(i,":hover")},n=function(){return r(t.scrollbarX,":focus")||r(t.scrollbarY,":focus")};t.event.bind(t.ownerDocument,"keydown",function(r){if(!(r.isDefaultPrevented&&r.isDefaultPrevented()||r.defaultPrevented)&&(l()||n())){var o=document.activeElement?document.activeElement:t.ownerDocument.activeElement;if(o){if("IFRAME"===o.tagName)o=o.contentDocument.activeElement;else for(;o.shadowRoot;)o=o.shadowRoot.activeElement;if(d(o))return}var s=0,a=0;switch(r.which){case 37:s=r.metaKey?-t.contentWidth:r.altKey?-t.containerWidth:-30;break;case 38:a=r.metaKey?t.contentHeight:r.altKey?t.containerHeight:30;break;case 39:s=r.metaKey?t.contentWidth:r.altKey?t.containerWidth:30;break;case 40:a=r.metaKey?-t.contentHeight:r.altKey?-t.containerHeight:-30;break;case 32:a=r.shiftKey?t.containerHeight:-t.containerHeight;break;case 33:a=t.containerHeight;break;case 34:a=-t.containerHeight;break;case 36:a=t.contentHeight;break;case 35:a=-t.contentHeight;break;default:return}t.settings.suppressScrollX&&0!==s||t.settings.suppressScrollY&&0!==a||(i.scrollTop-=a,i.scrollLeft+=s,R(t),e(s,a)&&r.preventDefault())}})},wheel:function(e){function i(t,i){var r=Math.floor(o.scrollTop),l=0===o.scrollTop,n=r+o.offsetHeight===o.scrollHeight,s=0===o.scrollLeft,a=o.scrollLeft+o.offsetWidth===o.scrollWidth;return!(Math.abs(i)>Math.abs(t)?l||n:s||a)||!e.settings.wheelPropagation}function r(t){var e=t.deltaX,i=-1*t.deltaY;return void 0!==e&&void 0!==i||(e=-1*t.wheelDeltaX/6,i=t.wheelDeltaY/6),t.deltaMode&&1===t.deltaMode&&(e*=10,i*=10),e!==e&&i!==i&&(e=0,i=t.wheelDelta),t.shiftKey?[-i,-e]:[e,i]}function l(e,i,r){if(!L.isWebKit&&o.querySelector("select:focus"))return!0;if(!o.contains(e))return!1;for(var l=e;l&&l!==o;){if(l.classList.contains(m.element.consuming))return!0;var n=t(l);if([n.overflow,n.overflowX,n.overflowY].join("").match(/(scroll|auto)/)){var s=l.scrollHeight-l.clientHeight;if(s>0&&!(0===l.scrollTop&&r>0||l.scrollTop===s&&r<0))return!0;var a=l.scrollWidth-l.clientWidth;if(a>0&&!(0===l.scrollLeft&&i<0||l.scrollLeft===a&&i>0))return!0}l=l.parentNode}return!1}function n(t){var n=r(t),s=n[0],a=n[1];if(!l(t.target,s,a)){var c=!1;e.settings.useBothWheelAxes?e.scrollbarYActive&&!e.scrollbarXActive?(a?o.scrollTop-=a*e.settings.wheelSpeed:o.scrollTop+=s*e.settings.wheelSpeed,c=!0):e.scrollbarXActive&&!e.scrollbarYActive&&(s?o.scrollLeft+=s*e.settings.wheelSpeed:o.scrollLeft-=a*e.settings.wheelSpeed,c=!0):(o.scrollTop-=a*e.settings.wheelSpeed,o.scrollLeft+=s*e.settings.wheelSpeed),R(e),(c=c||i(s,a))&&!t.ctrlKey&&(t.stopPropagation(),t.preventDefault())}}var o=e.element;void 0!==window.onwheel?e.event.bind(o,"wheel",n):void 0!==window.onmousewheel&&e.event.bind(o,"mousewheel",n)},touch:function(e){function i(t,i){var r=Math.floor(h.scrollTop),l=h.scrollLeft,n=Math.abs(t),o=Math.abs(i);if(o>n){if(i<0&&r===e.contentHeight-e.containerHeight||i>0&&0===r)return 0===window.scrollY&&i>0&&L.isChrome}else if(n>o&&(t<0&&l===e.contentWidth-e.containerWidth||t>0&&0===l))return!0;return!0}function r(t,i){h.scrollTop-=i,h.scrollLeft-=t,R(e)}function l(t){return t.targetTouches?t.targetTouches[0]:t}function n(t){return!(t.pointerType&&"pen"===t.pointerType&&0===t.buttons||(!t.targetTouches||1!==t.targetTouches.length)&&(!t.pointerType||"mouse"===t.pointerType||t.pointerType===t.MSPOINTER_TYPE_MOUSE))}function o(t){if(n(t)){var e=l(t);u.pageX=e.pageX,u.pageY=e.pageY,d=(new Date).getTime(),null!==p&&clearInterval(p)}}function s(e,i,r){if(!h.contains(e))return!1;for(var l=e;l&&l!==h;){if(l.classList.contains(m.element.consuming))return!0;var n=t(l);if([n.overflow,n.overflowX,n.overflowY].join("").match(/(scroll|auto)/)){var o=l.scrollHeight-l.clientHeight;if(o>0&&!(0===l.scrollTop&&r>0||l.scrollTop===o&&r<0))return!0;var s=l.scrollLeft-l.clientWidth;if(s>0&&!(0===l.scrollLeft&&i<0||l.scrollLeft===s&&i>0))return!0}l=l.parentNode}return!1}function a(t){if(n(t)){var e=l(t),o={pageX:e.pageX,pageY:e.pageY},a=o.pageX-u.pageX,c=o.pageY-u.pageY;if(s(t.target,a,c))return;r(a,c),u=o;var h=(new Date).getTime(),p=h-d;p>0&&(f.x=a/p,f.y=c/p,d=h),i(a,c)&&t.preventDefault()}}function c(){e.settings.swipeEasing&&(clearInterval(p),p=setInterval(function(){e.isInitialized?clearInterval(p):f.x||f.y?Math.abs(f.x)<.01&&Math.abs(f.y)<.01?clearInterval(p):(r(30*f.x,30*f.y),f.x*=.8,f.y*=.8):clearInterval(p)},10))}if(L.supportsTouch||L.supportsIePointer){var h=e.element,u={},d=0,f={},p=null;L.supportsTouch?(e.event.bind(h,"touchstart",o),e.event.bind(h,"touchmove",a),e.event.bind(h,"touchend",c)):L.supportsIePointer&&(window.PointerEvent?(e.event.bind(h,"pointerdown",o),e.event.bind(h,"pointermove",a),e.event.bind(h,"pointerup",c)):window.MSPointerEvent&&(e.event.bind(h,"MSPointerDown",o),e.event.bind(h,"MSPointerMove",a),e.event.bind(h,"MSPointerUp",c)))}}},H=function(r,l){var n=this;if(void 0===l&&(l={}),"string"==typeof r&&(r=document.querySelector(r)),!r||!r.nodeName)throw new Error("no element is specified to initialize PerfectScrollbar");this.element=r,r.classList.add(m.main),this.settings={handlers:["click-rail","drag-thumb","keyboard","wheel","touch"],maxScrollbarLength:null,minScrollbarLength:null,scrollingThreshold:1e3,scrollXMarginOffset:0,scrollYMarginOffset:0,suppressScrollX:!1,suppressScrollY:!1,swipeEasing:!0,useBothWheelAxes:!1,wheelPropagation:!0,wheelSpeed:1};for(var o in l)n.settings[o]=l[o];this.containerWidth=null,this.containerHeight=null,this.contentWidth=null,this.contentHeight=null;var s=function(){return r.classList.add(m.state.focus)},a=function(){return r.classList.remove(m.state.focus)};this.isRtl="rtl"===t(r).direction,this.isNegativeScroll=function(){var t=r.scrollLeft,e=null;return r.scrollLeft=-1,e=r.scrollLeft<0,r.scrollLeft=t,e}(),this.negativeScrollAdjustment=this.isNegativeScroll?r.scrollWidth-r.clientWidth:0,this.event=new y,this.ownerDocument=r.ownerDocument||document,this.scrollbarXRail=i(m.element.rail("x")),r.appendChild(this.scrollbarXRail),this.scrollbarX=i(m.element.thumb("x")),this.scrollbarXRail.appendChild(this.scrollbarX),this.scrollbarX.setAttribute("tabindex",0),this.event.bind(this.scrollbarX,"focus",s),this.event.bind(this.scrollbarX,"blur",a),this.scrollbarXActive=null,this.scrollbarXWidth=null,this.scrollbarXLeft=null;var c=t(this.scrollbarXRail);this.scrollbarXBottom=parseInt(c.bottom,10),isNaN(this.scrollbarXBottom)?(this.isScrollbarXUsingBottom=!1,this.scrollbarXTop=u(c.top)):this.isScrollbarXUsingBottom=!0,this.railBorderXWidth=u(c.borderLeftWidth)+u(c.borderRightWidth),e(this.scrollbarXRail,{display:"block"}),this.railXMarginWidth=u(c.marginLeft)+u(c.marginRight),e(this.scrollbarXRail,{display:""}),this.railXWidth=null,this.railXRatio=null,this.scrollbarYRail=i(m.element.rail("y")),r.appendChild(this.scrollbarYRail),this.scrollbarY=i(m.element.thumb("y")),this.scrollbarYRail.appendChild(this.scrollbarY),this.scrollbarY.setAttribute("tabindex",0),this.event.bind(this.scrollbarY,"focus",s),this.event.bind(this.scrollbarY,"blur",a),this.scrollbarYActive=null,this.scrollbarYHeight=null,this.scrollbarYTop=null;var h=t(this.scrollbarYRail);this.scrollbarYRight=parseInt(h.right,10),isNaN(this.scrollbarYRight)?(this.isScrollbarYUsingRight=!1,this.scrollbarYLeft=u(h.left)):this.isScrollbarYUsingRight=!0,this.scrollbarYOuterWidth=this.isRtl?f(this.scrollbarY):null,this.railBorderYWidth=u(h.borderTopWidth)+u(h.borderBottomWidth),e(this.scrollbarYRail,{display:"block"}),this.railYMarginHeight=u(h.marginTop)+u(h.marginBottom),e(this.scrollbarYRail,{display:""}),this.railYHeight=null,this.railYRatio=null,this.reach={x:r.scrollLeft<=0?"start":r.scrollLeft>=this.contentWidth-this.containerWidth?"end":null,y:r.scrollTop<=0?"start":r.scrollTop>=this.contentHeight-this.containerHeight?"end":null},this.isAlive=!0,this.settings.handlers.forEach(function(t){return T[t](n)}),this.lastScrollTop=Math.floor(r.scrollTop),this.lastScrollLeft=r.scrollLeft,this.event.bind(this.element,"scroll",function(t){return n.onScroll(t)}),R(this)};return H.prototype.update=function(){this.isAlive&&(this.negativeScrollAdjustment=this.isNegativeScroll?this.element.scrollWidth-this.element.clientWidth:0,e(this.scrollbarXRail,{display:"block"}),e(this.scrollbarYRail,{display:"block"}),this.railXMarginWidth=u(t(this.scrollbarXRail).marginLeft)+u(t(this.scrollbarXRail).marginRight),this.railYMarginHeight=u(t(this.scrollbarYRail).marginTop)+u(t(this.scrollbarYRail).marginBottom),e(this.scrollbarXRail,{display:"none"}),e(this.scrollbarYRail,{display:"none"}),R(this),W(this,"top",0,!1,!0),W(this,"left",0,!1,!0),e(this.scrollbarXRail,{display:""}),e(this.scrollbarYRail,{display:""}))},H.prototype.onScroll=function(t){this.isAlive&&(R(this),W(this,"top",this.element.scrollTop-this.lastScrollTop),W(this,"left",this.element.scrollLeft-this.lastScrollLeft),this.lastScrollTop=Math.floor(this.element.scrollTop),this.lastScrollLeft=this.element.scrollLeft)},H.prototype.destroy=function(){this.isAlive&&(this.event.unbindAll(),l(this.scrollbarX),l(this.scrollbarY),l(this.scrollbarXRail),l(this.scrollbarYRail),this.removePsClasses(),this.element=null,this.scrollbarX=null,this.scrollbarY=null,this.scrollbarXRail=null,this.scrollbarYRail=null,this.isAlive=!1)},H.prototype.removePsClasses=function(){this.element.className=this.element.className.split(" ").filter(function(t){return!t.match(/^ps([-_].+|)$/)}).join(" ")},H});
var Ps = null;

(function(window, undefined){

	function CIHelper(plugin)
	{
		this.plugin = plugin;
		this.ps;
		this.items = [];
		this.isVisible = false;
		this.isCurrentVisible = false;
	};

	CIHelper.prototype.createWindow = function()
	{
		var _body = document.body;
		var _head = document.getElementsByTagName('head')[0];
		if (!_body || !_head)
			return;

		var _style = document.createElement('style');
		_style.type      = 'text/css';

		var _style_body = ".ih_main { margin: 0px; padding: 0px; width: 100%; height: 100%; display: inline-block; overflow: hidden; box-sizing: border-box; user-select: none; position: fixed; border: 1px solid #cfcfcf; } ";
		_style_body += "ul { margin: 0px; padding: 0px; width: 100%; height: 100%; list-style-type: none; outline:none; } ";
		_style_body += "li { padding: 5px; font-family: \"Helvetica Neue\", Helvetica, Arial, sans-serif; font-size: 12px; font-weight: 400; color: #373737; } ";
		_style_body += "li:hover { background-color: #D8DADC; } ";
		_style_body += ".li_selected { background-color: #D8DADC; color: #373737; }";
		_style_body += ".li_selected:hover { background-color: #D8DADC; color: #373737; }";
		_style.innerHTML = _style_body;
		_head.appendChild(_style);

		document.body.style.background = "#FFFFFF";
		document.body.style.width = "100%";
		document.body.style.height = "100%";
		document.body.style.margin = "0";
		document.body.style.padding = "0";

		document.body.innerHTML = "<div class=\"ih_main\" id=\"ih_area\"><ul id=\"ih_elements_id\" role=\"listbox\"></ul></div>";

		this.ps = new PerfectScrollbar(document.getElementById("ih_area"), { minScrollbarLength: 20 });
		this.updateScrolls();

		this.createDefaultEvents();
	};

	CIHelper.prototype.setItems = function(items)
	{
		this.items = items;

		var _data = "";
		var _len = items.length;
		for (var i = 0; i < _len; i++)
		{
			if (undefined === items[i].id)
				items[i].id = "" + i;

			_data += "<li role=\"option\"";
			if (0 == i)
				_data += " class=\"li_selected\"";

			_data += " id=\"" + items[i].id + "\"";

			_data += " onclick=\"_private_on_ih_click(event)\">";
			_data += items[i].text;
			_data += "</li>";
		}

		document.getElementById("ih_elements_id").innerHTML = _data;
		this.updateScrolls();
		this.scrollToSelected();
	};

	CIHelper.prototype.createDefaultEvents = function()
	{
		this.plugin.onExternalMouseUp = function()
		{
			var evt = document.createEvent("MouseEvents");
			evt.initMouseEvent("mouseup", true, true, window, 1, 0, 0, 0, 0,
				false, false, false, false, 0, null);

			document.dispatchEvent(evt);
		};

		var _t = this;
		window.onkeydown = function(e) {
			switch (e.keyCode)
			{
				case 27: // Escape
				{
					if (_t.isVisible)
					{
						_t.isVisible = false;
						_t.plugin.executeMethod("UnShowInputHelper", [_t.plugin.info.guid, true]);
					}
					break;
				}
				case 38: // Up
				case 40: // Down
				case 9: // Tab
				case 36: // Home
				case 35: // End
				case 33: // PageUp
				case 34: // PageDown
				{
					var items = document.getElementsByTagName("li");
					var curIndex = -1;
					for (var i = 0; i < items.length; i++)
					{
						if (items[i].className == "li_selected")
						{
							curIndex = i;
							items[i].className = "";
							break;
						}
					}
					if (curIndex == -1)
					{
						curIndex = 0;
					}
					else
					{
						switch (e.keyCode)
						{
							case 38:
							{
								curIndex--;
								if (curIndex < 0)
									curIndex = 0;
								break;
							}
							case 40:
							{
								curIndex++;
								if (curIndex >= items.length)
									curIndex = items.length - 1;
								break;
							}
							case 9:
							{
								curIndex++;
								if (curIndex >= items.length)
									curIndex = 0;
								break;
							}
							case 36:
							{
								curIndex = 0;
								break;
							}
							case 35:
							{
								curIndex = items.length - 1;
								break;
							}
							case 33:
							case 34:
							{
								var _indexDif = 1;
								var _count = (document.getElementById("ih_area").clientHeight / 24) >> 0;
								if (_count > 1)
									_indexDif = _count;

								if (33 == e.keyCode)
								{
									curIndex -= _indexDif;
									if (curIndex < 0)
										curIndex = 0;
								}
								else
								{
									curIndex += _indexDif;
									if (curIndex >= items.length)
										curIndex = curIndex = items.length - 1;;
								}
								break;
							}
						}
					}

					if (curIndex < items.length)
					{
						items[curIndex].className = "li_selected";

						var _currentOffset = items[curIndex].offsetTop;
						var _currentHeight = items[curIndex].offsetHeight;

						var container = document.getElementById("ih_area");
						var _currentScroll = container.scrollTop;
						if (_currentOffset < _currentScroll)
						{
							if (container.scrollTo)
								container.scrollTo(0, _currentOffset);
							else
								container.scrollTop = _currentOffset;
						}
						else if ((_currentScroll + container.offsetHeight) < (_currentOffset + _currentHeight))
						{
							if (container.scrollTo)
								container.scrollTo(0, _currentOffset - (container.offsetHeight - _currentHeight));
							else
								container.scrollTop = _currentOffset - (container.offsetHeight - _currentHeight);
						}
					}
					break;
				}
				case 13:
				{
					_t.onSelectedItem();
					break;
				}
			}

			if (e.preventDefault)
				e.preventDefault();
			if (e.stopPropagation)
				e.stopPropagation();
			return false;
		};

		window.onresize = function(e)
		{
			_t.updateScrolls();
		};

		window._private_on_ih_click = function(e)
		{
			var items = document.getElementsByTagName("li");
			for (var i = 0; i < items.length; i++)
			{
				items[i].className = "";
			}
			e.target.className = "li_selected";
			var _id = e.target.getAttribute("id");
			_t.onSelectedItem();
		};

		this.plugin.event_onKeyDown = function(data)
		{
			window.onkeydown({ keyCode : data.keyCode });
		};
	};

	CIHelper.prototype.updateScrolls = function()
	{
		this.ps.update(); this.ps.update();

		var _elemV = document.getElementsByClassName("ps__rail-y")[0];
		var _elemH = document.getElementsByClassName("ps__rail-x")[0];

		if (!_elemH || !_elemV)
			return;

		var _styleV = window.getComputedStyle(_elemV);
		var _styleH = window.getComputedStyle(_elemH);

		var _visibleV = (_styleV && _styleV.display == "none") ? false : true;
		var _visibleH = (_styleH && _styleH.display == "none") ? false : true;

		if (_visibleH && _visibleV)
		{
			if ("13px" != _elemV.style.marginBottom)
				_elemV.style.marginBottom = "13px";
			if ("13px" != _elemH.style.marginRight)
				_elemH.style.marginRight = "13px";
		}
		else
		{
			if ("2px" != _elemV.style.marginBottom)
				_elemV.style.marginBottom = "2px";
			if ("2px" != _elemH.style.marginRight)
				_elemH.style.marginRight = "2px";
		}
	};

	CIHelper.prototype.scrollToSelected = function()
	{
		var items = document.getElementsByTagName("li");
		var curIndex = -1;
		for (var i = 0; i < items.length; i++)
		{
			if (items[i].className == "li_selected")
			{
				var container = document.getElementById("ih_area");
				if (container.scrollTo)
					container.scrollTo(0, items[i].offsetTop);
				else
					container.scrollTop = items[i].offsetTop;
				return;
			}
		}
	};

	CIHelper.prototype.getSelectedItem = function()
	{
		var items = document.getElementsByTagName("li");
		var curId = -1;
		for (var i = 0; i < items.length; i++)
		{
			if (items[i].className == "li_selected")
			{
				curId = items[i].getAttribute("id");
				break;
			}
		}

		if (-1 == curId)
			return null;

		var len = this.items.length;
		for (var i = 0; i < len; i++)
		{
			if (curId == this.items[i].id)
				return this.items[i];
		}

		return null;
	};

	CIHelper.prototype.onSelectedItem = function()
	{
		if (this.plugin.inputHelper_onSelectItem)
			this.plugin.inputHelper_onSelectItem(this.getSelectedItem());
	};

	CIHelper.prototype.show = function(w, h, isKeyboardTake)
	{
		this.isCurrentVisible = true;
		this.plugin.executeMethod("ShowInputHelper", [this.plugin.info.guid, w, h, isKeyboardTake], function() { window.Asc.plugin.ih.isVisible = true; });
	};

	CIHelper.prototype.unShow = function()
	{
		if (!this.isCurrentVisible && !this.isVisible)
			return;

		this.isCurrentVisible = false;
		window.Asc.plugin.executeMethod("UnShowInputHelper", [this.plugin.info.guid], function() { window.Asc.plugin.ih.isVisible = false; });
	};

	CIHelper.prototype.getItemHeight = function()
	{
		var _sizeItem = 24;
		var _items = document.getElementsByTagName("li");
		if (_items.length > 0 && _items[0].offsetHeight > 0)
			_sizeItem = _items[0].offsetHeight;
		return _sizeItem;
	};

	CIHelper.prototype.getItemsHeight = function(count)
	{
		return 2 + count * this.getItemHeight();
	};

	CIHelper.prototype.getItems = function()
	{
		return this.items;
	};

	CIHelper.prototype.getScrollSizes = function()
	{
		var _size = { w : 0, h : 0 };
		var _sizeItem = this.getItemHeight();

		var _elem = document.getElementById("ih_elements_id");
		if (_elem)
		{
			_size.w = _elem.scrollWidth;
			_size.h = 2 + this.items.length * _sizeItem;
		}
		return _size;
	};

	window.Asc = window.Asc || {};
	window.Asc.inputHelper = CIHelper;

})(window, undefined);

(function(window, undefined){

	function onReloadPage(isCtrl)
	{
		window.parent.postMessage(JSON.stringify({
			type : "reload",
			guid : window.Asc.plugin.guid,
			ctrl : isCtrl
		}), "*");
	}
	function onBaseKeyDown(e)
	{
		var isCtrl = (e.metaKey || e.ctrlKey) ? true : false;
		if (e.keyCode == 116)
		{
			onReloadPage(isCtrl);
			if (e.preventDefault)
				e.preventDefault();
			if (e.stopPropagation)
				e.stopPropagation();
			return false;
		}
	}

	if (window.addEventListener)
		window.addEventListener("keydown", onBaseKeyDown, false);
	else
		window.attachEvent("keydown", onBaseKeyDown);

})(window, undefined);

(function(window, undefined){

    var g_isMouseSendEnabled = false;
    var g_language = "";

    window.plugin_sendMessage = function sendMessage(data)
    {
        window.parent.postMessage(data, "*");
    };

    window.plugin_onMessage = function(event)
    {
        if (!window.Asc.plugin)
            return;

        if (typeof(event.data) == "string")
        {
            var pluginData = {};
            try
            {
                pluginData = JSON.parse(event.data);
            }
            catch(err)
            {
                pluginData = {};
            }

            var type = pluginData.type;

            if (pluginData.guid != window.Asc.plugin.guid)
            {
                if (undefined !== pluginData.guid)
                    return;

                switch (type)
                {
                    case "onExternalPluginMessage":
                        break;
                    default:
                        return;
                }
            }

            if (type == "init")
                window.Asc.plugin.info = pluginData;

            if (!window.Asc.plugin.tr || !window.Asc.plugin.tr_init)
            {
				window.Asc.plugin.tr_init = true;
                window.Asc.plugin.tr = function(val) {
                    if (!window.Asc.plugin.translateManager || !window.Asc.plugin.translateManager[val])
                        return val;
                    return window.Asc.plugin.translateManager[val];
                };
            }

            var newLang = "";
            if (window.Asc.plugin.info)
                newLang = window.Asc.plugin.info.lang;
            if (newLang == "" || newLang != g_language)
            {
                g_language = newLang;
                if (g_language == "en-EN" || g_language == "")
				{
					window.Asc.plugin.translateManager = {};
					if (window.Asc.plugin.onTranslate)
						window.Asc.plugin.onTranslate();
				}
				else
				{
					var _client = new XMLHttpRequest();
					_client.open("GET", "./translations/" + g_language + ".json");

					_client.onreadystatechange = function ()
					{
						if (_client.readyState == 4)
						{
						    if (_client.status == 200 || location.href.indexOf("file:") == 0)
						    {
                                try
                                {
                                    window.Asc.plugin.translateManager = JSON.parse(_client.responseText);
                                }
                                catch (err)
                                {
                                    window.Asc.plugin.translateManager = {};
                                }
                            }
                            else
                            {
                                window.Asc.plugin.translateManager = {};
                            }


                            if (window.Asc.plugin.onTranslate)
                                window.Asc.plugin.onTranslate();
						}
					};
					_client.send();
				}
            }

            switch (type)
            {
                case "init":
                {
                    pluginStart();
                    window.Asc.plugin.init(window.Asc.plugin.info.data);
                    break;
                }
                case "button":
                {
                    var _buttonId = parseInt(pluginData.button);
                    if (!window.Asc.plugin.button && -1 == _buttonId)
						window.Asc.plugin.executeCommand("close", "");
                    else
                        window.Asc.plugin.button(_buttonId);
                    break;
                }
                case "enableMouseEvent":
                {
                    g_isMouseSendEnabled = pluginData.isEnabled;
					if (window.Asc.plugin.onEnableMouseEvent)
						window.Asc.plugin.onEnableMouseEvent(g_isMouseSendEnabled);
                    break;
                }
                case "onExternalMouseUp":
                {
                    if (window.Asc.plugin.onExternalMouseUp)
                        window.Asc.plugin.onExternalMouseUp();
                    break;
                }
                case "onMethodReturn":
                {
					window.Asc.plugin.isWaitMethod = false;

					if (window.Asc.plugin.methodCallback)
					{
					    var methodCallback = window.Asc.plugin.methodCallback;
                        window.Asc.plugin.methodCallback = null;
                        methodCallback(pluginData.methodReturnData);
                        methodCallback = null;
					}
					else if (window.Asc.plugin.onMethodReturn)
					{
						window.Asc.plugin.onMethodReturn(pluginData.methodReturnData);
					}

					if (window.Asc.plugin.executeMethodStack && window.Asc.plugin.executeMethodStack.length > 0)
                    {
                        var obj = window.Asc.plugin.executeMethodStack.shift();
                        window.Asc.plugin.executeMethod(obj.name, obj.params, obj.callback);
                    }

                    break;
                }
				case "onCommandCallback":
				{
                    if (window.Asc.plugin.onCallCommandCallback)
                    {
                        window.Asc.plugin.onCallCommandCallback();
                        window.Asc.plugin.onCallCommandCallback = null;
                    }
                    else if (window.Asc.plugin.onCommandCallback)
						window.Asc.plugin.onCommandCallback();
					break;
				}
                case "onExternalPluginMessage":
                {
					if (window.Asc.plugin.onExternalPluginMessage && pluginData.data && pluginData.data.type)
						window.Asc.plugin.onExternalPluginMessage(pluginData.data);
					break;
                }
                case "onEvent":
                {
                    if (window.Asc.plugin["event_" + pluginData.eventName])
                        window.Asc.plugin["event_" + pluginData.eventName](pluginData.eventData);
                    break;
                }
                default:
                    break;
            }
        }
    };

    function pluginStart()
    {
        if (window.Asc.plugin.isStarted)
            return;

        window.Asc.plugin.isStarted = true;
        window.startPluginApi();

        window.Asc.plugin.checkPixelRatio = function(isAttack)
        {
            if (window.Asc.plugin.checkedPixelRatio && true !== isAttack)
                return;

            window.Asc.plugin.checkedPixelRatio = true;

            var userAgent = navigator.userAgent.toLowerCase();
            var isIE = (userAgent.indexOf("msie") > -1 || userAgent.indexOf("trident") > -1 || userAgent.indexOf("edge") > -1);
            var isChrome = !isIE && (userAgent.indexOf("chrome") > -1);
            var isMozilla = !isIE && (userAgent.indexOf("firefox") > -1);

            var zoom = 1.0;
            var isRetina = false;
            var retinaPixelRatio = 1;

            var isMobileVersion = window.Asc.plugin.info ? window.Asc.plugin.info.isMobileMode : false;

            // пока отключаем мозиллу... хотя почти все работает
            if ((/*isMozilla || */isChrome) && document && document.firstElementChild && document.body && !isMobileVersion)
            {
                if (window.devicePixelRatio > 0.1)
                {
                    if (window.devicePixelRatio < 1.99)
                    {
                        zoom = window.devicePixelRatio;
                    }
                    else
                    {
                        zoom = window.devicePixelRatio / 2;
                        retinaPixelRatio = 2;
                        isRetina = true;
                    }

                    document.firstElementChild.style.zoom = 1.0 / zoom;
                }
                else
                {
                    document.firstElementChild.style.zoom = "normal";
                }
            }
            else
            {
                isRetina = (Math.abs(2 - window.devicePixelRatio) < 0.01);
                if (isRetina)
                    retinaPixelRatio = 2;

                if (isMobileVersion)
                {
                    isRetina = (window.devicePixelRatio >= 1.9);
                    retinaPixelRatio = window.devicePixelRatio;
                }
            }

            window.Asc.plugin.zoom = zoom;
            window.Asc.plugin.retinaPixelRatio = retinaPixelRatio;
        };

        window.Asc.plugin.checkPixelRatio();
    }

    window.onmousemove = function(e)
    {
        if (!g_isMouseSendEnabled || !window.Asc.plugin || !window.Asc.plugin.executeCommand)
            return;

        var _x = (undefined === e.clientX) ? e.pageX : e.clientX;
        var _y = (undefined === e.clientY) ? e.pageY : e.clientY;

        window.Asc.plugin.executeCommand("onmousemove", JSON.stringify({ x : _x, y : _y }));

    };
    window.onmouseup   = function(e)
    {
        if (!g_isMouseSendEnabled || !window.Asc.plugin || !window.Asc.plugin.executeCommand)
            return;

        var _x = (undefined === e.clientX) ? e.pageX : e.clientX;
        var _y = (undefined === e.clientY) ? e.pageY : e.clientY;

        window.Asc.plugin.executeCommand("onmouseup", JSON.stringify({ x : _x, y : _y }));
    };

    window.plugin_sendMessage(JSON.stringify({ guid : window.Asc.plugin.guid, type : "initialize_internal" }));

})(window, undefined);
