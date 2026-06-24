/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
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
					if (asc_typeOf(h[i]) === "function") {
						h[i].apply(this, a);
					}
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
					if (t === "function") {
						h.push(old);
					}
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
