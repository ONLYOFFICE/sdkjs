/*
 *
 * (c) Copyright Ascensio System Limited 2010-2016
 *
 * This program is freeware. You can redistribute it and/or modify it under the terms of the GNU 
 * General Public License (GPL) version 3 as published by the Free Software Foundation (https://www.gnu.org/copyleft/gpl.html). 
 * In accordance with Section 7(a) of the GNU GPL its Section 15 shall be amended to the effect that 
 * Ascensio System SIA expressly excludes the warranty of non-infringement of any third-party rights.
 *
 * THIS PROGRAM IS DISTRIBUTED WITHOUT ANY WARRANTY; WITHOUT EVEN THE IMPLIED WARRANTY OF MERCHANTABILITY OR
 * FITNESS FOR A PARTICULAR PURPOSE. For more details, see GNU GPL at https://www.gnu.org/copyleft/gpl.html
 *
 * You can contact Ascensio System SIA by email at sales@onlyoffice.com
 *
 * The interactive user interfaces in modified source and object code versions of ONLYOFFICE must display 
 * Appropriate Legal Notices, as required under Section 5 of the GNU GPL version 3.
 *
 * Pursuant to Section 7  3(b) of the GNU GPL you must retain the original ONLYOFFICE logo which contains 
 * relevant author attributions when distributing the software. If the display of the logo in its graphic 
 * form is not reasonably feasible for technical reasons, you must include the words "Powered by ONLYOFFICE" 
 * in every copy of the program you distribute. 
 * Pursuant to Section 7  3(e) we decline to grant you any rights under trademark law for use of our trademarks.
 *
*/
"use strict";

(
	/**
	 * @param {Window} window
	 * @param {undefined} undefined
	 */
	function (window, undefined) {
		/*
		 * Import
		 * -----------------------------------------------------------------------------
		*/
		var asc = window["Asc"],
			asc_typeOf = asc.typeOf;


		/** @constructor */
		function asc_CHandlersList(handlers) {
			this.handlers = handlers || {};
			return this;
		}

		asc_CHandlersList.prototype.hasTrigger = function (eventName) {
			return null != this.handlers[eventName];
		};

		asc_CHandlersList.prototype.trigger = function (eventName) {
			var h = this.handlers[eventName], t = asc_typeOf(h), a = Array.prototype.slice.call(arguments, 1), i;
			if (t === "function") {
				return h.apply(this, a);
			}
			if (t === "array") {
				for (i = 0; i < h.length; i += 1) {
					if (asc_typeOf(h[i]) === "function") {h[i].apply(this, a);}
				}
				return true;
			}
			return false;
		};
		asc_CHandlersList.prototype.add = function (eventName, eventHandler, replaceOldHandler) {
			var th = this.handlers, h, old, t;
			if (replaceOldHandler || !th.hasOwnProperty(eventName)) {
				th[eventName] = eventHandler;
			} else {
				old = h = th[eventName];
				t = asc_typeOf(old);
				if (t !== "array") {
					h = th[eventName] = [];
					if (t === "function") {h.push(old);}
				}
				h.push(eventHandler);
			}
		};
		asc_CHandlersList.prototype.remove = function (eventName, eventHandler) {
			var th = this.handlers, h = th[eventName], i;
			if (th.hasOwnProperty(eventName)) {
				if (asc_typeOf(h) !== "array" || asc_typeOf(eventHandler) !== "function") {
					delete th[eventName];
					return true;
				}
				for (i = h.length - 1; i >= 0; i -= 1) {
					if (h[i] === eventHandler) {
						delete h[i];
						return true;
					}
				}
			}
			return false;
		};

		//---------------------------------------------------------export---------------------------------------------------
		window['AscCommonExcel'] = window['AscCommonExcel'] || {};
		AscCommonExcel.asc_CHandlersList = asc_CHandlersList;
	}
)(window);