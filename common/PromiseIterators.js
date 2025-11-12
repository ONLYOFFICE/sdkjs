/*
 * (c) Copyright Ascensio System SIA 2010-2024
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
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
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

(function (window) {
	function CPromiseGetterIterator(arrFPromiseGetters) {
		this.promiseGetters = arrFPromiseGetters;
	}

	CPromiseGetterIterator.prototype.getPromise = function (nIndex) {
		const fPromiseGetter = this.promiseGetters[nIndex];
		if (fPromiseGetter) {
			return fPromiseGetter();
		}
	};

	CPromiseGetterIterator.prototype.forEachValue = function (fSuccess, fReject, fAfterAll) {
		let nIndex = 0;
		const oThis = this;

		const resolveIndexPromise = function () {
			const oPromise = oThis.getPromise(nIndex);
			if (oPromise) {
				oPromise.then(function (oValue) {
					nIndex += 1;
					fSuccess && fSuccess(oValue);
					resolveIndexPromise();
				}, function (reason) {
					nIndex += 1;
					fReject && fReject(reason);
					resolveIndexPromise();
				});
			} else {
				fAfterAll && fAfterAll();
			}
		};
		resolveIndexPromise();
	};
	CPromiseGetterIterator.prototype.forAllSuccessValues = function (fCallback) {
		const arrValues = [];
		const fAfterAll = function () {
			fCallback(arrValues);
		};
		this.forEachValue(function (oValue) {
			arrValues.push(oValue);
		}, undefined, fAfterAll);
	};

	window.AscCommon = window.AscCommon || {};
	window.AscCommon.CPromiseGetterIterator = CPromiseGetterIterator;

})(window);
