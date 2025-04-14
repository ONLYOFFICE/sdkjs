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
QUnit.config.autostart = false;
$(function () {
	let api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function () {
		}
	};
	let docInfo = new Asc.asc_CDocInfo();
	docInfo.asc_putTitle("TeSt.xlsx");
	api.DocInfo = docInfo;
	api.initCollaborativeEditing({});
	window["Asc"]["editor"] = api;

	waitLoadModules(function () {
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		startTests();
	});

	function waitLoadModules(waitCallback) {
		Asc.spreadsheet_api.prototype._init = function () {
			this._loadModules();
		};
		Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
			callback();
		};
		Asc.spreadsheet_api.prototype.onEndLoadFile = function (fonts, callback) {
			waitCallback();
		};
		AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {
		};
		AscCommonExcel.WorkbookView.prototype._init = function () {
		};
		AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {
		};
		AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {
		};
		AscCommonExcel.WorksheetView.prototype._init = function () {
		};
		AscCommonExcel.WorksheetView.prototype.updateRanges = function () {
		};
		AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {
		};
		AscCommonExcel.WorksheetView.prototype.setSelection = function () {
		};
		AscCommonExcel.WorksheetView.prototype.draw = function () {
		};
		AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
		};
		AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
		};
	}

	function openDocument(file) {
		if (api.wbModel) {
			api.asc_CloseFile();
		}

		api.isOpenOOXInBrowser = false;
		api.OpenDocumentFromBin(null, AscCommon.Base64.decode(file["Editor.bin"]));
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
		return api.wbModel;
	}

	function prepareTest(assert, wb) {
		api.wb.model = wb;
		api.wbModel = wb;
		api.initGlobalObjects(wb);
		api.handlers.remove("getSelectionState");
		api.handlers.add("getSelectionState", function () {
			return null;
		});
		api.handlers.remove("asc_onError");
		api.handlers.add("asc_onError", function (code, level) {
			assert.equal(code, 0, "asc_onError");
		});
		AscCommon.History.Clear();
	}

	let memory = new AscCommon.CMemory();

	function Utf8ArrayToStr(array) {
		let out, i, len, c;
		let char2, char3;

		out = "";
		len = array.length;
		i = 0;
		while (i < len) {
			c = array[i++];
			switch (c >> 4) {
				case 0:
				case 1:
				case 2:
				case 3:
				case 4:
				case 5:
				case 6:
				case 7:
					// 0xxxxxxx
					out += String.fromCharCode(c);
					break;
				case 12:
				case 13:
					// 110x xxxx   10xx xxxx
					char2 = array[i++];
					out += String.fromCharCode(((c & 0x1F) << 6) | (char2 & 0x3F));
					break;
				case 14:
					// 1110 xxxx  10xx xxxx  10xx xxxx
					char2 = array[i++];
					char3 = array[i++];
					out += String.fromCharCode(((c & 0x0F) << 12) |
						((char2 & 0x3F) << 6) |
						((char3 & 0x3F) << 0));
					break;
			}
		}

		return out;
	}

	function getXml(pivot, addCacheDefinition) {
		memory.Seek(0);
		pivot.toXml(memory);
		if (addCacheDefinition) {
			memory.WriteXmlString('\n\n');
			pivot.cacheDefinition.toXml(memory);
		}
		let buffer = new Uint8Array(memory.GetCurPosition());
		for (let i = 0; i < memory.GetCurPosition(); i++) {
			buffer[i] = memory.data[i];
		}
		if (typeof TextDecoder !== "undefined") {
			return new TextDecoder("utf-8").decode(buffer);
		} else {
			return Utf8ArrayToStr(buffer);
		}

	}

	AscCommonExcel.WorksheetView.prototype._calcActiveCellOffset = function () {
		return {row: 0, col: 0}
	}
	AscCommonExcel.WorksheetView.prototype._calcActiveCellOffset = function () {
		return {row: 0, col: 0}
	}

	QUnit.module("Formulas Performance Tests");

	function startTests() {
		QUnit.start();
		QUnit.test('Test: editCell', function (assert) {
			const file = Asc.formulaPerformance;
			const wb = openDocument(file);
			prepareTest(assert, wb);
			const ws = wb.getWorksheetByName('WED Allocation_ Processed count');
			const range = ws.getRange3(3, 5, 3, 5)
			const fragment = new AscCommonExcel.Fragment();
			fragment.text = "NK2800592";
			fragment.charCodes = []

			for (let i = 0; i < fragment.text.length; i++) {
				fragment.charCodes.push(fragment.text.charCodeAt(i));
			}
			const standardTime = 3000;
			const testsCount = 3
			let averageDuration = 0;
			for (let i = 0; i < testsCount; i += 1) {
				const start = performance.now();
				range.setValue2([fragment], true);
				const end = performance.now();
				const duration = end - start;
				averageDuration += duration;
			}
			averageDuration /= testsCount;
			assert.ok(averageDuration < standardTime, `Set Value average time: ${averageDuration.toFixed(2)} ms`);
		});

	}
});
