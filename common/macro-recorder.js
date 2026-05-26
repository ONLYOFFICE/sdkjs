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

(function (window)
{
	let oMacroRecorderContext = {};

	function VarRegistry()
	{
		this.byId = {};
		this.counters = {};

		this.register = function(id, type)
		{
			if (!this.counters.hasOwnProperty(type))
				this.counters[type] = 0;
			this.counters[type]++;
			let varName = type + this.counters[type];
			this.byId[id] = varName;
			return varName;
		};
		this.inc = function(type)
		{
			if (!this.counters.hasOwnProperty(type))
				this.counters[type] = 0;
			this.counters[type]++;
			let varName = type + this.counters[type];
			this.byId[type] = varName;
			return varName;
		};
		this.get = function(id)
		{
			return this.byId[id];
		};
		this.reset = function()
		{
			this.byId = {};
			this.counters = {};
		};
	}
	let varRegistry = new VarRegistry();

	/**
	 * @param editor
	 * @constructor
	 */
	function MacroRecorder(editor)
	{
		this.editor = editor;
		this.inProgress = false;
		this.paused = false;
		this.macroName = "";
		this.result = "";

		this.prevChangeType = null;
		this.prevData = undefined;
		this.actionCount = 0;
		this.isFirstAction = null;
		this.currentDescription = null;

		varRegistry.reset();
		oMacroRecorderContext = this;
	}
	
	MacroRecorder.prototype.start = function(macroName)
	{
		if (this.inProgress)
			return;
		
		this.macroName = macroName;
		this.result = "";
		this.paused = false;
		this.inProgress = true;
		this.isFirstAction = true;
		varRegistry.reset();

		this.initEvents();
		this.editor.asc_registerCallback('asc_onKeyDown', this.onKeyDown);

		this.editor.sendEvent("asc_onMacroRecordingStart");
	};
	MacroRecorder.prototype.initEvents = function()
	{
		let _t = this;

		_t.onKeyDown = function(e)
		{
			if (e.KeyCode === 8) // BackSpace
			{
				if (_t.editor.editorId !== AscCommon.c_oEditorId.Word && _t.editor.editorId !== AscCommon.c_oEditorId.Presentation)
					return;

				_t.addStepData("remove", 1);
				// TODO when we have:
				// * Selection.Delete
				// * Selection.TypeBackspace
			}
			else if (e.KeyCode === 9) // Tab
			{
				if (_t.editor.editorId !== AscCommon.c_oEditorId.Word && _t.editor.editorId !== AscCommon.c_oEditorId.Presentation)
					return;

				let doc = _t.editor.getLogicDocument();
				if (doc.GetCurrentTable())
				{
					_t.addStepData("moveCursorRight", [{
						isRtl:			false,
						isAddSelect:	e.IsShift(),
						isWord:			e.IsCtrl()
					}]);
				}
			}
			else if (e.KeyCode === 13) // Enter
			{
			}
			else if (e.KeyCode === 27) // Esc
			{
			}
			else if (e.KeyCode === 32) // Space
			{
			}
			else if (e.KeyCode === 33) // PgUp
			{
			}
			else if (e.KeyCode === 34) // PgDn
			{
			}
			else if (e.KeyCode === 35) // End
			{
			}
			else if (e.KeyCode === 36) // Home
			{
			}
			else if (e.KeyCode === 37) // Left Arrow
			{
				if (_t.editor.editorId !== AscCommon.c_oEditorId.Word && _t.editor.editorId !== AscCommon.c_oEditorId.Presentation)
					return;

				let doc = _t.editor.getLogicDocument();
				let curPara = doc.GetCurrentParagraph(true);
				let isRtl = (curPara ? curPara.isRtlDirection() : false);

				let type = isRtl
					? "moveCursorRight"
					: "moveCursorLeft"

				_t.addStepData(type, [{
					isRtl:			isRtl,
					isAddSelect:	e.IsShift(),
					isWord:			e.IsCtrl()
				}]);
			}
			else if (e.KeyCode === 38) // Top Arrow
			{
				if (_t.editor.editorId !== AscCommon.c_oEditorId.Word && _t.editor.editorId !== AscCommon.c_oEditorId.Presentation)
					return;

				_t.addStepData('moveCursorUp', [{
					isAddSelect:	e.IsShift(),
					isWord:			e.IsCtrl()
				}]);
			}
			else if (e.KeyCode === 39) // Right Arrow
			{
				if (_t.editor.editorId !== AscCommon.c_oEditorId.Word && _t.editor.editorId !== AscCommon.c_oEditorId.Presentation)
					return;

				let doc = _t.editor.getLogicDocument();
				let curPara = doc.GetCurrentParagraph(true);
				let isRtl = (curPara ? curPara.isRtlDirection() : false);

				let type = isRtl
					? "moveCursorLeft"
					: "moveCursorRight"

				_t.addStepData(type, [{
					isRtl:			isRtl,
					isAddSelect:	e.IsShift(),
					isWord:			e.IsCtrl()
				}]);
			}
			else if (e.KeyCode === 40) // Bottom Arrow
			{
				if (_t.editor.editorId !== AscCommon.c_oEditorId.Word && _t.editor.editorId !== AscCommon.c_oEditorId.Presentation)
					return;

				_t.addStepData('moveCursorDown', [{
					isAddSelect:	e.IsShift(),
					isWord:			e.IsCtrl()
				}]);
			}
			else if (e.KeyCode === 46) // Delete
			{
			}
			else if (e.KeyCode === 144) // Num Lock
			{
			}
			else if (e.KeyCode === 145) // Scroll Lock
			{
			}
		};
	};
	MacroRecorder.prototype.stop = function()
	{
		if (!this.inProgress)
			return;

		this.editor.asc_unregisterCallback('asc_onKeyDown', this.onKeyDown);

		this.inProgress = false;
		this.paused = false;

		if (this.prevData !== undefined && this.prevChangeType)
		{
			this.getResultByType(this.prevChangeType, this.prevData);
			this.prevData = [];
			this.prevChangeType = null;
			varRegistry.reset();
		}

		let macroData = "";
		try
		{
			let data = this.editor.macros.GetData();
			if (data && "" !== data)
			{
				macroData = JSON.parse(this.editor.macros.GetData());
			}
			else
			{
				macroData = {
					"macrosArray" : [],
					"current"     : -1
				};
			}
		}
		catch (e)
		{
			return;
		}
		
		let name = this.macroName ? this.macroName : this.getNewName(macroData["macrosArray"]);
		let value = "(function()\n{\n" + this.result + "})();"
		macroData["macrosArray"].push({
			"guid" : AscCommon.CreateUUID(true),
			"name" : name,
			"autostart" : false,
			"value" : value
		});
		
		macroData["current"] = macroData["macrosArray"].length - 1;
		this.editor.asc_setMacros(JSON.stringify(macroData));
		this.editor.sendEvent("asc_onMacroRecordingStop");
	};
	MacroRecorder.prototype.cancel = function()
	{
		if (!this.inProgress)
			return;
		
		this.inProgress = false;
		this.paused = false;
		this.editor.sendEvent("asc_onMacroRecordingStop");
	};
	MacroRecorder.prototype.pause = function()
	{
		if (!this.inProgress || this.paused)
			return;
		
		this.paused = true;
		this.editor.sendEvent("asc_onMacroRecordingPause");
	};
	MacroRecorder.prototype.resume = function()
	{
		if (!this.inProgress || !this.paused)
			return;
		
		this.paused = false;
		this.editor.sendEvent("asc_onMacroRecordingResume");
	};
	MacroRecorder.prototype.isInProgress = function()
	{
		return this.inProgress;
	};
	MacroRecorder.prototype.isPaused = function()
	{
		return this.paused;
	};
	MacroRecorder.prototype.addDefualtVaribalesForEditor = function()
	{
		if (this.editor.editorId === AscCommon.c_oEditorId.Word)
			this.proceedDefualtVariablesForWord();
		else if (this.editor.editorId === AscCommon.c_oEditorId.Spreadsheet)
			this.proceedDefualtVariablesForSpreadsheet();
		else if (this.editor.editorId === AscCommon.c_oEditorId.Presentation)
			this.proceedDefualtVariablesForPresentation();
	};
	MacroRecorder.prototype.getMacrosListForEditor = function()
	{
		let actionsMacros = null;

		if (this.editor.editorId === AscCommon.c_oEditorId.Word)
			actionsMacros = DEActionMap;
		else if (this.editor.editorId === AscCommon.c_oEditorId.Spreadsheet)
			actionsMacros = SEActionMap;
		else if (this.editor.editorId === AscCommon.c_oEditorId.Presentation)
			actionsMacros = PEActionMap;

		return actionsMacros;
	}
	MacroRecorder.prototype.onFinalizeAction = function()
	{
		// todo implement onFinalizeAction
		this.getResultByType(this.prevChangeType, this.prevData);
		this.prevChangeType	= null;
		this.prevData		= undefined;
	};
	MacroRecorder.prototype.addStepData = function(type, additional, isStart)
	{
		let list = this.getMacrosListForEditor();
		if (list && !list[type])
			return;
		if (isStart === true)
			this.currentDescription = type;

		if (isStart === false)
			type = this.currentDescription;

		if (!this.isInProgress()
			|| this.isPaused()
			|| undefined === additional
			|| !AscCommon.g_oTableId.IsOn()
			|| (Array.isArray(additional) && additional.length === 0))
			return;

		this.filterAction(type, additional);
	};
	const ACTION = {
		NOT_WRITE_PREV: 0,
		NOT_WRITE_NOW: 1,
		EXIT: 2,
	};
	MacroRecorder.prototype.filterRules = [
		{
			prevType: "DeselectDrawing",
			editor: [AscCommon.c_oEditorId.Word],
			curType: AscDFH.historydescription_Document_AddNewShape,
			check: function(type, additional) {
				return this.prevData.name === additional.id || this.prevData.includes(additional.id)
			},
			action: ACTION.NOT_WRITE_PREV
		},
		{
			prevType: "DeselectDrawing",
			editor: [AscCommon.c_oEditorId.Word],
			curType: "SetDrawingPos",
			check: function(type, additional) {
				return this.prevData.name === additional.id || this.prevData.includes(additional.id)
			},
			action: ACTION.NOT_WRITE_PREV
		},
		{
			prevType: "SetDrawingPos",
			editor: [AscCommon.c_oEditorId.Word],
			curType: "SelectDrawing",
			check: function(type, additional) {
				let prevData = this.prevData;
				let checkFunc = function(shape) { return shape.getObjectName() === prevData.id };
				return additional.filter(checkFunc).length != 0
			},
			action: ACTION.NOT_WRITE_NOW
		},
		{
			prevType: "SelectDrawing",
			editor: [AscCommon.c_oEditorId.Word],
			curType: "SetDrawingPos",
			check: function(type, additional) {
				let prevData = this.prevData;
				let checkFunc = function(shape) { return prevData.getObjectName() === shape.id };
				return additional.filter(checkFunc).length != 0
			},
			action: ACTION.NOT_WRITE_NOW
		},
		{
			prevType: "SetDrawingRotation",
			editor: [AscCommon.c_oEditorId.Word],
			curType: "SetDrawingRotation",
			check: function(type, additional) {
				return this.prevData === additional || (Array.isArray(this.prevData) && this.prevData[this.prevData.length - 1] === additional)
			},
			action: ACTION.EXIT
		},
	];
	MacroRecorder.prototype.filterAction = function(type, additional)
	{
		if (type === "remove") {
			return this.getResultByType(type, additional);
		}
		else {
			for (let nFiler = 0; nFiler < this.filterRules.length; nFiler++) {
				let currentRule = this.filterRules[nFiler];

				if (currentRule.prevType === this.prevChangeType
					&& currentRule.curType === type
					&& currentRule.editor.includes(this.editor.editorId)) {
					let bAction = currentRule.check.call(this, type, additional);

					if (!bAction) continue;
					let action = currentRule.action;

					switch (action) {
						case ACTION.NOT_WRITE_PREV:
							this.prevData = undefined;
							this.prevChangeType = undefined;
							break;
						case ACTION.NOT_WRITE_NOW:
							type = "none";
							break;
						case ACTION.EXIT:
							return;
						default:
							break;
					}
					break;
				}
			}
		}
		if (this.prevChangeType === type)
		{
			this.prevData = this.joinDataForMacros(this.prevData, additional);
		}
		else if (additional !== undefined)
		{
			this.getResultByType(this.prevChangeType, this.prevData);
			this.prevChangeType	= type;
			this.prevData		= additional;
		}

		if (this.isFirstAction)
		{
			this.addDefualtVaribalesForEditor();
			this.isFirstAction = false;
		}
	};
	MacroRecorder.prototype.getResultByType = function(type, additional)
	{
		let actionsMacros = this.getMacrosListForEditor();
		let actionMacroFunction = actionsMacros[type];
		if (actionMacroFunction)
		{
			if (Array.isArray(additional)
				&& type !== AscDFH.historydescription_Document_AddLetter
				&& type !== AscDFH.historydescription_Presentation_ParagraphAdd
				&& type !== 'moveCursorLeft'
				&& type !== 'moveCursorRight'
				&& type !== 'moveCursorUp'
				&& type !== 'moveCursorDown'
				&& type !== 'SelectDrawing'
				&& type !== 'SetTableBorder'
			)
			{
				for (let i = 0; i < additional.length; i++)
				{
					this.result += actionMacroFunction(additional[i], type);
				}
			}
			else
				this.result += actionMacroFunction(additional, type);
		}
	};
	MacroRecorder.prototype.joinDataForMacros = function(prevData, currentData) {
		if (Array.isArray(prevData))
		{
			if (Array.isArray(currentData))
			{
				for (let i = 0; i < currentData.length; i++)
				{
					prevData.push(currentData[i]);
				}
			}
			else
			{
				prevData.push(currentData);
			}
			return prevData
		}
		else
		{
			if (prevData === undefined || prevData === null)
			{
				return [currentData];
			}
			else
			{
				return [prevData, currentData];
			}
		}
	};
	MacroRecorder.prototype.proceedDataBefoeApply = function (data) {
		if (!data || typeof data !== 'object')
			return data;

		if (this.editor.editorId === AscCommon.c_oEditorId.Presentation && Array.isArray(data) && data.length === 0)
			return [[undefined]];

		var out = {};
		for (var key in data)
		{
			if (!Object.prototype.hasOwnProperty.call(data, key))
				continue;

			var val = data[key];
			var isArr = Array.isArray
				? Array.isArray(val)
				: Object.prototype.toString.call(val) === '[object Array]';

			if (isArr)
				out[key] = val;
			else
				out[key] = [val];
		}
		return out;
	};
	MacroRecorder.prototype.getNewName = function(macros)
	{
		let maxId = 0;
		for (let i = 0, count = macros.length; i < count; ++i)
		{
			if (0 !== macros[i].name.indexOf("Macro "))
				continue;
			
			let curId = parseInt(macros[i].name.substr(6));
			if (isNaN(curId))
				continue;
			
			maxId = Math.max(curId, maxId);
		}
		
		return "Macro " + (maxId + 1);
	};
	MacroRecorder.prototype.proceedDefualtVariablesForWord = function()
	{
		this.result += "\tlet doc = Api.GetDocument();\n"
			+ "";
	};
	MacroRecorder.prototype.proceedDefualtVariablesForSpreadsheet = function()
	{
		this.result += "\tlet worksheet = Api.GetActiveSheet();\n"
			+ "\tlet workbook = Api.GetActiveWorkbook();\n";
	};
	MacroRecorder.prototype.proceedDefualtVariablesForPresentation = function()
	{
		this.result += "\tlet presentation = Api.GetPresentation();\n"
			+ "";
	};

	function private_getTableVar(tableId)
	{
		let editorId = oMacroRecorderContext.editor.editorId;

		let varName = varRegistry.get(tableId);
		if (varName)
			return { varName: varName, declaration: "" };

		varName = varRegistry.register(tableId, "table");

		let getSelectedTables = "";
		if (editorId === AscCommon.c_oEditorId.Word)
		{
			getSelectedTables = " = doc.GetSelectedTable();\n";
		}
		else if (editorId === AscCommon.c_oEditorId.Presentation)
		{
			getSelectedTables = " = presentation.GetActiveTable();\n";
		}

		return { varName: varName, declaration: "\tlet " + varName + getSelectedTables };
	}

	function private_GetAlignH(nAlign)
	{
		if (c_oAscAlignH.Left === nAlign)
			return "left";
		else if (c_oAscAlignH.Right === nAlign)
			return "right";
		else if (c_oAscAlignH.Center === nAlign)
			return "center";

		return "left";
	}

	function private_GetAlignV(nAlign)
	{
		if (c_oAscAlignV.Top === nAlign)
			return "top";
		else if (c_oAscAlignV.Bottom === nAlign)
			return "bottom";
		else if (c_oAscAlignV.Center === nAlign)
			return "center";

		return "center";
	}

	function iterByDataFn(object, key, templateFn, type) {
		if (!object && !Array.isArray(object[0]))
			return "";

		let out = "";
		if (Array.isArray(object[0]))
		{
			for (let i = 0; i < object[0].length; i++) {
				out += templateFn(object[0][i]);
			}
		}
		else
		{
			if (!object[key])
			{
				out += templateFn(object, type);
				return out;
			}

			let arr = object[key];

			if (Array.isArray(arr))
			{
				for (let i = 0; i < arr.length; i++) {
					out += templateFn(arr[i], type);
				}
			}
			else
			{
				out += templateFn(object, type);
			}
		}

		return out;
	};
	function makeAction(key, templateFn) {
		return function(additional, type) {
		  return iterByDataFn(additional || {}, key, templateFn, type);
		};
	};

	function private_ChartInternalTypeToBuilder(sType) {
		switch (sType) {
			case Asc.c_oAscChartTypeSettings.unknown: {
				return "unknown";
			}
			case Asc.c_oAscChartTypeSettings.barNormal: {
				return "bar";
			}
			case Asc.c_oAscChartTypeSettings.barStacked: {
				return "barStacked";
			}
			case Asc.c_oAscChartTypeSettings.barStackedPer: {
				return "barStackedPercent";
			}
			case Asc.c_oAscChartTypeSettings.barNormal3d: {
				return "bar3D";
			}
			case Asc.c_oAscChartTypeSettings.barStacked3d: {
				return "barStacked3D";
			}
			case Asc.c_oAscChartTypeSettings.barStackedPer3d: {
				return "barStackedPercent3D";
			}
			case Asc.c_oAscChartTypeSettings.barNormal3dPerspective: {
				return "barStackedPercent3DPerspective";
			}
			case Asc.c_oAscChartTypeSettings.hBarNormal: {
				return "horizontalBar";
			}
			case Asc.c_oAscChartTypeSettings.hBarStacked: {
				return "horizontalBarStacked";
			}
			case Asc.c_oAscChartTypeSettings.hBarStackedPer: {
				return "horizontalBarStackedPercent";
			}
			case Asc.c_oAscChartTypeSettings.hBarNormal3d: {
				return "horizontalBar3D";
			}
			case Asc.c_oAscChartTypeSettings.hBarStacked3d: {
				return "horizontalBarStacked3D";
			}
			case Asc.c_oAscChartTypeSettings.hBarStackedPer3d: {
				return "horizontalBarStackedPercent3D";
			}
			case Asc.c_oAscChartTypeSettings.lineNormal: {
				return "lineNormal";
			}
			case Asc.c_oAscChartTypeSettings.lineStacked: {
				return "lineStacked";
			}
			case Asc.c_oAscChartTypeSettings.lineStackedPer: {
				return "lineStackedPercent";
			}
			case Asc.c_oAscChartTypeSettings.line3d: {
				return "line3D";
			}
			case Asc.c_oAscChartTypeSettings.pie: {
				return "pie";
			}
			case Asc.c_oAscChartTypeSettings.pie3d: {
				return "pie3D";
			}
			case Asc.c_oAscChartTypeSettings.doughnut: {
				return "doughnut";
			}
			case Asc.c_oAscChartTypeSettings.scatter: {
				return "scatter";
			}
			case Asc.c_oAscChartTypeSettings.stock: {
				return "stock";
			}
			case Asc.c_oAscChartTypeSettings.areaNormal: {
				return "area";
			}
			case Asc.c_oAscChartTypeSettings.areaStacked: {
				return "areaStacked";
			}
			case Asc.c_oAscChartTypeSettings.areaStackedPer: {
				return "areaStackedPercent";
			}
			case Asc.c_oAscChartTypeSettings.comboBarLine: {
				return "comboBarLine";
			}
			case Asc.c_oAscChartTypeSettings.comboBarLineSecondary: {
				return "comboBarLineSecondary";
			}
			case Asc.c_oAscChartTypeSettings.comboCustom: {
				return "comboCustom";
			}
			case Asc.c_oAscChartTypeSettings.radar: {
				return "radar";
			}
			case Asc.c_oAscChartTypeSettings.radarMarker: {
				return "radarMarker";
			}
			case Asc.c_oAscChartTypeSettings.radarFilled: {
				return "radarFilled";
			}
		}
		return "unknown";
	}

	function sameConditional(a, b, ignoreRtl) {
		if (a.isAddSelect !== b.isAddSelect) return false;
		if (a.isWord !== b.isWord) return false;

		if (!ignoreRtl) {
			var aHas = a.hasOwnProperty('isRtl');
			var bHas = b.hasOwnProperty('isRtl');
			if (aHas && bHas && a.isRtl !== b.isRtl) return false;
		}
		return true;
	}

	function pickPattern(obj, ignoreRtl) {
		var p = {
			isAddSelect: obj.isAddSelect,
			isWord: obj.isWord
		};
		if (!ignoreRtl && obj.hasOwnProperty('isRtl')) p.isRtl = obj.isRtl;
		return p;
	}

	function groupDataForCursor(arr, ignoreRtl) {
		if (!arr || arr.length === 0) return [];
		var ignoreRtl = !!ignoreRtl;

		var runs = [];
		var current = pickPattern(arr[0], ignoreRtl);
		var count = 1;

		for (var i = 1; i < arr.length; i++) {
			var item = arr[i];
			var view = pickPattern(item, ignoreRtl);

			if (sameConditional(view, current, ignoreRtl)) {
				count++;
			} else {
				runs.push({ pattern: current, count: count });
				current = view;
				count = 1;
			}
		}
		runs.push({ pattern: current, count: count });

		return runs;
	}

	function getWrappingStyleName(props) {
		var type   = props.type;
		var style  = props.style;
		var behind = props.behind;

		if (type === 0 && style === Asc.c_oAscWrapStyle2.Inline)
			return "inline";
		if (style === Asc.c_oAscWrapStyle2.Behind && behind)
			return "behind";
		if (style === Asc.c_oAscWrapStyle2.InFront && !behind)
			return "inFront";

		switch (style) {
			case Asc.c_oAscWrapStyle2.Square:
				return "square";
			case Asc.c_oAscWrapStyle2.Tight:
				return "tight";
			case Asc.c_oAscWrapStyle2.Through:
				return "through";
			case Asc.c_oAscWrapStyle2.TopAndBottom:
				return "topAndBottom";
			default:
				return "square";
		}
	}

	const wordActions = {
		setTextBold				: function(bold){return "\tdoc.GetRangeBySelect().SetBold(" + bold + ");\n"},
		setTextItalic			: function(italic){return "\tdoc.GetRangeBySelect().SetItalic(" + italic + ");\n"},
		setTextUnderline		: function(underline){return "\tdoc.GetRangeBySelect().SetUnderline(" + underline + ");\n"},
		setTextStrikeout		: function(strikeout){return "\tdoc.GetRangeBySelect().SetStrikeout(" + strikeout + ");\n"},
		setTextFontName			: function(name){return "\tdoc.GetRangeBySelect().SetFontFamily(\"" + name + "\");\n"},
		setTextFontSize			: function(size){return "\tdoc.GetRangeBySelect().SetFontSize(\"" + size + "\");\n"},
		setTextHighlightColor	: function(highlight){
			let highlightColor = "";
			if (highlight)
			{
				let color = new AscCommonWord.CDocumentColor(highlight.r, highlight.g, highlight.b);
				highlightColor = color.ToHighlightColor();
			}

			if (highlightColor === "")
				highlightColor = 'none';

			return "\tdoc.GetRangeBySelect().SetHighlight(\"" + highlightColor + "\");\n";
		},
		setTextHighlightNone	: function(){return "\tdoc.GetRangeBySelect().SetHighlight(\"none\");\n"},
		setTextVertAlign		: function(baseline, type){
			let align = "baseline";
			if (baseline === true)
				align = "baseline";
			else if (AscDFH.historydescription_Document_SetTextVertAlignHotKey3 === type)
				align = "subscript";
			else if (AscDFH.historydescription_Document_SetTextVertAlignHotKey2 === type)
				align = "superscript";

			return "\tApi.GetDocument().GetRangeBySelect().SetVertAlign(\"" + align + "\");\n";
		},
		setTextColor			: function(color){return "\tdoc.GetRangeBySelect().SetColor(" + color.r + "," + color.g + "," + color.b  + ");\n"},
		setStyleHeading			: function(name){return "\tdoc.GetRangeBySelect().SetStyle(doc.GetStyle(\"" + name + "\"));\n"},
		clearFormat				: function(){return "\tdoc.GetRangeBySelect().ClearFormating()\n"},
		cut						: function(){return "\tdoc.GetRangeBySelect().Cut();\n"},
		changeTextCase			: function(changeType){return ""; return "\tdoc.GetRangeBySelect().SetTextCase(\"" + changeType + "\");\n"},
		incFontSize				: function(){return ""; return "\tdoc.GetRangeBySelect().Grow();\n"},
		addLetter				: function(textArr){
			let textStr = "";
			for (let i = 0; i < textArr.length; ++i)
			{
				textStr += String.fromCodePoint(textArr[i]);
			}

			return "\tdoc.EnterText(\"" + textStr + "\");\n"
		},
		setAlign				: function(align){
			switch (align) {
				case AscCommon.align_Left:		align = 'left';		break;
				case AscCommon.align_Right:		align = 'right';	break;
				case AscCommon.align_Justify:	align = 'both';		break;
				case AscCommon.align_Center:	align = 'center';	break;
				default:						align = 'center';
			}
			return "\tdoc.GetCurrentParagraph().SetJc(\"" + align + "\");\n"
		},
		setParagraphShd			: function(color){return "\tdoc.GetRangeBySelect().SetShd(\"clear\", "+ color.asc_getR() + " , " + color.asc_getG() +", "+ color.asc_getB() +", false);\n"},
		setLineSpacing			: function(lineSpacing){
			let type = lineSpacing.type;
			let value = lineSpacing.value;

			switch(type)
			{
				case Asc.linerule_Auto		: type = "auto";	break;
				case Asc.linerule_AtLeast	: type = "atLeast";	break;
				case Asc.linerule_Exact		: type = "exact";	break;
				default						: type = "auto";	break;
			}

			return "\tdoc.GetRangeBySelect().GetAllParagraphs().forEach(function(para) {\n" +
				"\t\tpara.SetSpacingLine(" + value + " * 240, \"" + type + "\");\n" +
				"\t});\n";
		},
		// incIndentetLineSpacing	: function(){
		// 	// for now we don't have relative increaee/decrease for api
		// 	//"\tdoc.GetRangeBySelect().GetAllParagraphs().forEach(para => para.SetIndFirstLine());\n"
		// 	//paragraph.SetIndFirstLine(1440);
		// },
		setParagraphNumbering	: function(num){
			return "\tlet " + varRegistry.inc('numbering') + " = doc.CreateNumbering(\"" + num.Type + "\");\n" +
				"\tdoc.GetRangeBySelect().GetAllParagraphs().forEach(function(para) {\n" +
				"\t\tpara.SetNumbering(" + varRegistry.get('numbering') + ".GetLevel(0));\n" +
				"\t\tpara.SetContextualSpacing(true);\n" +
				"\t});\n";
		},
		addParagraph			: function(){
			return "\tdoc.InsertParagraphBreak();\n";
		},
		addBlankPage			: function(){return "\tdoc.InsertBlankPage();\n"},
		addPageBreak			: function(type){
			if (type === AscWord.break_Page)
				return "\tdoc.GetCurrentParagraph().AddPageBreak();\n" // to api selection
			else if (type === AscWord.break_Column)
				return "\tdoc.GetCurrentParagraph().AddColumnBreak();\n" // to api selection
		},
		addSectionBreak			: function(type){
			if (type === Asc.c_oAscSectionBreakType.NextPage)
				return "\tdoc.CreateSection(doc.GetCurrentParagraph()).SetType(\"nextPage\");\n";
			else if (type === Asc.c_oAscSectionBreakType.Column)
				return "\tdoc.CreateSection(doc.GetCurrentParagraph()).SetType(\"nextColumn\");\n";
			else if (type === Asc.c_oAscSectionBreakType.Continuous)
				return "\tdoc.CreateSection(doc.GetCurrentParagraph()).SetType(\"continuous\");\n";
			else if (type === Asc.c_oAscSectionBreakType.EvenPage)
				return "\tdoc.CreateSection(doc.GetCurrentParagraph()).SetType(\"evenPage\");\n";
			else if (type === Asc.c_oAscSectionBreakType.OddPage)
				return "\tdoc.CreateSection(doc.GetCurrentParagraph()).SetType(\"oddPage\");\n";
		},
		addTable				: function(prop){
			let varName = prop.tableId
				? varRegistry.register(prop.tableId, "table")
				: varRegistry.inc('table');
			let width = prop.tableGrid.map(function(num) { return AscCommon.MMToTwips(num, -1) });
			let styleStr = prop.style ? prop.style : "";
			let styleVar = varRegistry.inc('tableStyle');

			return "\tlet " + varName + " = (function()\n"
				+ "\t{\n"
				+ "\t\tlet table = Api.CreateTable(" + prop.col + ", " + prop.row + ");\n"
				+ "\t\tlet " + styleVar + " = doc.GetStyle(\"" + styleStr + "\");\n"
				+ "\t\tif (" + styleVar + ") table.SetStyle(" + styleVar + ");\n"
				+ "\t\ttable.SetTableLayout(\"autofit\");\n"
				+ "\t\tlet row = table.GetRow(0);\n"
				+ "\t\t[" + width.toString() + "].forEach(function(width, index){\n"
				+ "\t\t\tlet cellColumnFirst = row.GetCell(index);\n"
				+ "\t\t\tlet cellPr = Api.CreateTableCellPr();\n"
				+ "\t\t\tcellPr.SetWidth(\"twips\", width);\n"
				+ "\t\t\tcellColumnFirst.SetCellPr(cellPr);\n"
				+ "\t\t});\n"
				+ "\t\tdoc.Push(table);\n"
				+ "\t\ttable.GetRange(0, 0).Select();\n"
				+ "\t\treturn table;\n"
				+ "\t}());\n";
		},
		applyTablePr			: function(style)
		{
			return "\t(function () {\n"
					+ "\t\tlet table = doc.GetCurrentParagraph().GetParentTable();\n"
					+ "\t\tlet tableStyle = doc.GetStyle(\"" + style + "\");\n"
					+ "\t\t(table && tableStyle) && table.SetStyle(tableStyle);\n"
				+ "\t}());\n";
		},
		addImage				: function(image){
			function PxToEMU96(px){ return px * 9525; }
			function CmToPx96(cm){ return Math.round(cm * 96 / 2.54); }
			function CmToEMU(cm){ return Math.round(cm * 360000); }

			function SizeByWidthThreshold(origWpx, origHpx, targetWidthCm)
			{
				const thresholdPx = CmToPx96(targetWidthCm);
				if (origWpx < thresholdPx){
				  return { wEMU: PxToEMU96(origWpx), hEMU: PxToEMU96(origHpx), scaled: false };
				}
				else
				{
				  const aspect = origHpx / origWpx;
				  return {
					wEMU: CmToEMU(targetWidthCm),
					hEMU: CmToEMU(targetWidthCm * aspect),
					scaled: true
				  };
				}
			}

			let size = SizeByWidthThreshold(image.Image.naturalWidth, image.Image.naturalHeight, 16.5);
			let text = "";
			if (image instanceof AscFonts.CImage)
			{
				text += "\tlet " + varRegistry.inc('image') + " = Api.CreateImage(\"" + image.src + "\", " + size.wEMU + ", " + size.hEMU + ");\n"
				text += "\tdoc.GetCurrentParagraph().AddDrawing("+ varRegistry.get('image') + ");\n"
			}
			return text;
		},
		addChart				: function(chart){ //todo title
			let series = chart.getAllSeries();
			let seriesNames = [];
			let seriesData = [];
			let numformat	= [];
			for (let i = 0; i < series.length; i++)
			{
				let currSer = series[i];
				let name = currSer.asc_getSeriesName();
				let value = currSer.asc_getValuesArr();
				let format = currSer.getCatSourceNumFormat();
				seriesNames.push(name);
				seriesData.push(value);
				numformat.push(format);
			}

			let categories	= chart.getCatValues();
			let chartType	= private_ChartInternalTypeToBuilder(chart.getChartType());
			let width		= chart.GetWidth() * 36000.0;		//mm2emu
			let height		= chart.GetHeight() * 36000.0;		//mm2emu
			let style		= chart.getChartStyleIdx();
			let title		= chart.getChartTitle().getDocContent().GetText();
			title = title ? title.replace(/[\r\n\t]+/g, '') : "";

			let value = "\tlet " + varRegistry.inc('chart') + " = Api.CreateChart(\n"
				+ "\t\t\"" + chartType + "\",\n"
				+ "\t\t" + JSON.stringify(seriesData) + ",\n"
				+ "\t\t" + JSON.stringify(seriesNames) + ",\n"
				+ "\t\t" + JSON.stringify(categories) + ",\n"
				+ "\t\t" + width + ",\n"
				+ "\t\t" + height + ",\n"
				+ "\t\t" + style + "\n"
				+ "\t);\n"
			//+ "\t" + varRegistry.get('chart') + ".SetTitle(\"" + title + "\", " + 14 + ");\n"
			+ "\tdoc.GetCurrentParagraph().AddDrawing(" + varRegistry.get('chart') + ");\n"
			return value;
		},
		addHyperlink			: function(hl){
			// create hyperlink text
			return "\tdoc.GetRangeBySelect().AddHyperlink(\"" + (hl.Value ? hl.Value : hl.Text) + "\", \"" + hl.ToolTip + "\");\n"
		},
		addShape				: function(shapeProps){
			let fill = shapeProps.fill.getRGBAColor();
			let border = shapeProps.border;
			let borderwidth = border.w / 36000;
			let borderColor = border.Fill.getRGBAColor();
			let varName = varRegistry.register(shapeProps.id, 'shape');

			return "\tlet " + varName + " = (function () {\n" +
					"\t\tlet fill = Api.CreateSolidFill(Api.CreateRGBColor("+ fill.R +", " + fill.G + ", " + fill.B + "));\n" +
					"\t\tlet stroke = Api.CreateStroke(" + borderwidth +" * 36000, Api.CreateSolidFill(Api.CreateRGBColor("+ borderColor.R +", " + borderColor.G + ", " + borderColor.B + ")));\n" +
					"\t\tlet shape = Api.CreateShape(\"" + shapeProps.type + "\", " + shapeProps.extX + " * 36000, " + shapeProps.extY + " * 36000, fill, stroke);\n" +
					"\t\tshape.SetWrappingStyle(\"inFront\");\n" +
					"\t\tdoc.AddDrawingToPage(shape, doc.GetCurrentPage(), " + shapeProps.pos.x + " * 36000.0, " + shapeProps.pos.y + " * 36000.0 )\n" +
					"\t\tshape.Select(true);\n" +
					"\t\treturn shape\n" +
				"\t}());\n";
		},
		removeHdr				: function(hdr){
			if (hdr.isHeader)
				return "\tdoc.GetFinalSection().RemoveHeader(\"default\");\n";
			else
				return "\tdoc.GetFinalSection().RemoveFooter(\"default\");\n";
		},
		addComment				: function(commentData){
			return "\tlet " + varRegistry.inc('comment') + " = doc.AddComment(\"" + commentData.m_sText + "\", \"" + commentData.m_sUserName + "\", \"" + commentData.m_sUserId + "\");\n"
				+ "\t" + varRegistry.get('comment') + ".SetTime(" + commentData.m_sTime + ")\n";
			// todo add time
		},
		addMath					: function(type){
			let mathText = "";
			AscFormat.ExecuteNoHistory(function () {
				let paraMath = new AscCommonWord.ParaMath();
				paraMath.Root.Load_FromMenu(type);
				paraMath.Root.Correct_Content(true);
				mathText = paraMath.GetTextOfElement().GetText();
			}, this, []);

			return "\tdoc.AddMathEquation(\"" + mathText + "\", \"unicode\");\n"
		},
		addMathHotkey			: function(obj){
			let type = 'unicode';
			if (obj.type === 1)
				type === "latex";
			else if (obj.type === 2)
				type === "mathml"

			return "\tdoc.AddMathEquation(\"" + obj.math + "\", \"" + type + "\");\n";
		},
		addBlockContentControl	: function(strPlaceholder){
			return "\tlet " + varRegistry.inc('block') + " = Api.CreateBlockLvlSdt();\n"
				+ "\t" + varRegistry.get('block') + ".SetPlaceholderText(\"" + strPlaceholder + "\")\n"
				+ "\tdoc.Push(" + varRegistry.get('block') + ");\n"
				+ "\t" + varRegistry.get('block') + ".Select();\n";
		},
		addInlineContentControl	: function(strPlaceholder){
				return "\tlet " + varRegistry.inc('inline') + " = Api.CreateInlineLvlSdt();\n"
					+ "\t" + varRegistry.get('inline') + ".SetPlaceholderText(\"" + strPlaceholder + "\")\n"
					+ "\tdoc.GetCurrentParagraph().Push(" + varRegistry.get('inline') + ");\n"
					+ "\t" + varRegistry.get('inline') + ".Select();\n";
		},
		addContentControlList	: function(props){
			if (props.isComboBox === true)
				return "\tdoc.AddComboBoxContentControl();\n";
			else if (props.isComboBox === false)
				return "\tdoc.AddDropDownListContentControl();\n"
		},
		addContentControlCheckBox	: function(){
			return "\tdoc.AddCheckBoxContentControl({checked : false});\n";
		},
		addContentControlDatePicker	: function(){
			return "\tdoc.AddDatePickerContentControl();\n"
		},
		addContentControlPicture: function(){
			return "\tdoc.AddPictureContentControl(180 * 10000, 180 * 10000);\n";
		},
		moveCursorLeft			: function(arrData){
			let data = groupDataForCursor(arrData);
			let text = "";

			for (let i = 0; i < data.length; i++)
			{
				let currentChange = data[i];
				let nCount = currentChange.count;
				let pattern = currentChange.pattern;

				if (pattern.isRtl)
					text += "\tdoc.MoveCursorRight(" + nCount + ", " + pattern.isAddSelect + ", " + pattern.isWord + ");\n";
				else
					text += "\tdoc.MoveCursorLeft(" + nCount + ", " + pattern.isAddSelect + ", " + pattern.isWord + ");\n";
			}

			return text;
		},
		moveCursorRight			: function(arrData){
			let data = groupDataForCursor(arrData);
			let text = "";

			for (let i = 0; i < data.length; i++)
			{
				let currentChange	= data[i];
				let nCount			= currentChange.count;
				let pattern			= currentChange.pattern;

				if (pattern.isRtl)
					text += "\tdoc.MoveCursorLeft(" + nCount + ", " + pattern.isAddSelect + ", " + pattern.isWord + ");\n";
				else
					text += "\tdoc.MoveCursorRight(" + nCount + ", " + pattern.isAddSelect + ", " + pattern.isWord + ");\n";
			}

			return text;
		},
		moveCursorUp			: function(arrData){
			let data = groupDataForCursor(arrData, true);
			let text = "";

			for (let i = 0; i < data.length; i++)
			{
				let currentChange	= data[i];
				let nCount			= currentChange.count;
				let pattern			= currentChange.pattern;

				text += "\tdoc.MoveCursorUp(" + nCount + ", " + pattern.isAddSelect + ", " + pattern.isWord + ");\n";
			}
			return text;
		},
		moveCursorDown			: function(arrData){
			let data = groupDataForCursor(arrData, true);
			let text = "";

			for (let i = 0; i < data.length; i++)
			{
				let currentChange	= data[i];
				let nCount			= currentChange.count;
				let pattern			= currentChange.pattern;

				text += "\tdoc.MoveCursorDown(" + nCount + ", " + pattern.isAddSelect + ", " + pattern.isWord + ");\n";
			}
			return text;
		},
		backSpaceButton			: function(){
			return "";
		},
		deleteButton			: function(){
			return "";
		},
		remove					: function(){
			if (oMacroRecorderContext.prevChangeType === AscDFH.historydescription_Document_AddLetter && oMacroRecorderContext.prevData.length)
				oMacroRecorderContext.prevData.pop();
			return "";
		},
		selectDrawing			: function(shapes)
		{
			let str = "";
			let macroShapes = [];

			shapes.forEach(function(element) {
				if (!macroShapes.includes(element) && varRegistry.get(element.getObjectName()) !== undefined)
					macroShapes.push(element);
			});

			let ids = [];

			for (let i = 0; i < macroShapes.length; i++) {
				let varName = varRegistry.get(macroShapes[i].getObjectName());
				if (varName !== undefined)
					ids.push(varName);
			}

			ids = ids.join(", ");

			if (macroShapes.length !== 0)
				str = str + "\t[" + ids + "]\n\t\t.forEach(function(drawing, index){drawing.Select(index === 0)});\n"

			return str;
		},
		deselectDrawing			: function(oAdditional)
		{
			// for non-stored shape don't write anything
			let varName = varRegistry.get(oAdditional.name);
			if (varName === undefined) {
				return "";
			}
			return "\t[" + varName + "]\n\t\t.forEach(function(drawing){drawing.Unselect()});\n"
		},
		setDrawingFill			: function(unifill)
		{
			//TODO add transparent when add api
			if (unifill.fill.type === Asc.c_oAscFill.FILL_TYPE_SOLID)
			{
				let color = unifill.fill.color.color.RGBA;
				//let transparent = unifill.transparent;

				return "\tdoc.GetSelectedDrawings()\n" +
					"\t\t.forEach(function(shape) {\n" +
					"\t\t\tshape.Fill(Api.CreateSolidFill(Api.CreateRGBColor(" + color.R + ", " + color.G + ", " + color.B + ")));\n" +
					"\t\t});\n";
			}
			else if (unifill.fill instanceof AscFormat.CGradFill)
			{
				let colors = unifill.fill.colors;
				let strColor = "";

				for(let nColor = 0; nColor < colors.length; nColor++)
				{
					let gs = colors[nColor];
					let color = gs.color.color.RGBA;
					strColor += (nColor !== 0)
						? ",\n\t\t\t\t"
						: "\n\t\t\t\t";
					strColor += "Api.CreateGradientStop(Api.CreateRGBColor(" + color.R + ", " + color.G + ", " + color.B + "), " + gs.pos + ")";
				}

				let gradient = "";
				if (unifill.fill.lin) // leinear
					gradient = "\n\t\t\tApi.CreateLinearGradientFill([" + strColor + "\n\t\t\t], " + unifill.fill.lin.angle + ")";
				else if (unifill.fill.path) // radial
					gradient = "\n\t\t\tApi.CreateRadialGradientFill([" + strColor + "\n\t\t\t])";

				return "\tdoc.GetSelectedDrawings()\n"
      				+ "\t\t.forEach(function(shape) { shape.Fill(" + gradient + "); }\n\t\t);\n";
			}
			else if (unifill.fill instanceof AscFormat.CBlipFill)
			{
				let base64data = unifill.fill.getBase64RasterImageId()
				let blipFill = "\n\t\t\tApi.CreateBlipFill(\"" + base64data + "\", \"" + (unifill.fill.tile ? "tile" : "stretch") + "\")";
				return "\tdoc.GetSelectedDrawings()\n"
    		 		+ "\t\t.forEach(function(shape) { shape.Fill(" + blipFill + "); }\n\t\t);\n";
			}
			else if (unifill.fill instanceof AscFormat.CPattFill)
			{
				let bgClr = unifill.fill.bgClr.color.RGBA;
				let fgClr = unifill.fill.fgClr.color.RGBA;
				let ftype = unifill.fill.ftype;
				let type = AscCommon.global_hatch_names[ftype];

				let blipFill = "\n\t\t\tApi.CreatePatternFill(\n\t\t\t\t\"" + type + "\",\n\t\t\t\t" + "Api.CreateRGBColor(" + bgClr.R + ", " + bgClr.G + ", " + bgClr.B + "),\n\t\t\t\tApi.CreateRGBColor(" + fgClr.R + ", " + fgClr.G + ", " + fgClr.B + ")";
				return "\tdoc.GetSelectedDrawings()\n"
       				+ "\t\t.forEach(function(shape) { shape.Fill(" + blipFill + "); }\n\t\t\t)\n\t\t);\n";
			}
			else if (unifill.fill instanceof AscFormat.CNoFill)
			{
				return "\tdoc.GetSelectedDrawings()\n"
       				+ "\t\t.forEach(function(shape) { shape.Fill(Api.CreateNoFill()); });\n";
			}
		},
		setDrawingLine			: function(line)
		{
			let strStrokeShape = "\t(function () {\n";
			let type = AscFormat.CLn.prototype.GetDashByCode(line.prstDash);
			let color = line.Fill.fill.color.color.RGBA;
			strStrokeShape += "\t\tlet stroke = Api.CreateStroke(\n" +
					"\t\t\t" + line.w / 12700.0 + " * 12700.0,\n" +
					"\t\t\t" + "Api.CreateSolidFill(Api.CreateRGBColor(" + color.R + ", " + color.G + ", " + color.B + ")),\n" +
					"\t\t\t\"" + type + "\"\n" +
				"\t\t);\n";

			strStrokeShape += "\t\tdoc.GetSelectedDrawings()\n"
                 + "\t\t\t.forEach(function(draw) { draw.SetOutLine(stroke); });\n";

			strStrokeShape += "\t}());\n";

			return strStrokeShape;
		},
		setDrawingRotation		: function(nRot)
		{
			return "\tdoc.GetSelectedDrawings()\n"
    			+ "\t\t.forEach(function(shape) { shape.SetRotation(" + nRot * 180 / Math.PI + "); });\n";
		},
		setDrawingFlipH			: function(isFlip)
		{
			return "\tdoc.GetSelectedDrawings()\n"
    			+ "\t\t.forEach(function(shape) { shape.SetHorFlip(" + isFlip + "); });\n";
		},
		setDrawingFlipV			: function(isFlip)
		{
			return "\tdoc.GetSelectedDrawings()\n"
    			+ "\t\t.forEach(function(shape) { shape.SetVertFlip(" + isFlip + "); });\n";
		},
		setDrawingWrapping		: function(props)
		{
			return "\tdoc.GetSelectedDrawings().forEach(function(draw) { draw.SetWrappingStyle(\"" + getWrappingStyleName(props) + "\"); })\n";
		},
		setDrawingPos			: function(pos)
		{
			return "\tdoc.GetDrawingsByName([\"" + pos.name + "\"]).forEach(function(draw) {\n" +
				"\t\tdraw.SetVerPosition(\"page\", " + pos.y + " * 36000.0);\n" +
				"\t\tdraw.SetHorPosition(\"page\", " + pos.x + " * 36000.0);\n" +
				"\t});\n"
		},
		setPositionH			: function(data)
		{
			let relative = "";
			switch(data.relativeFrom)
			{
				 case Asc.c_oAscRelativeFromH.Character:	relative = "character";		break;
				 case Asc.c_oAscRelativeFromH.Column:		relative = "column";		break;
				 case Asc.c_oAscRelativeFromH.LeftMargin:	relative = "leftMargin";	break;
				 case Asc.c_oAscRelativeFromH.Margin:		relative = "margin";		break;
				 case Asc.c_oAscRelativeFromH.RightMargin:	relative = "rightMargin";	break;
				 case Asc.c_oAscRelativeFromH.Page:			relative = "page";			break;
				 default:									relative = "page";
			}

			if (data.useAlign)
			{
				let useAlign = "";
				switch (data.value) {
					case Asc.c_oAscAlignH.Center:	useAlign = "center";	break;
					case Asc.c_oAscAlignH.Left:		useAlign = "left";		break;
					case Asc.c_oAscAlignH.Right:	useAlign = "right";		break;
					default:						useAlign = "center";
				}

				return "\tdoc.GetSelectedDrawings().forEach(function(draw) {\n" +
					"\t\tdraw.SetHorAlign(\"" + relative + "\", \"" + useAlign + "\");\n" +
					"\t});\n"

			}

			if (data.percent)
			{
				return "\tdoc.GetSelectedDrawings().forEach(function(draw) {\n" +
					"\t\tdraw.SetHorPosition(\"" + relative + "\", " + data.value + ", " + data.percent + ");\n" +
					"\t});\n"
			}
			else
			{
				return "\tdoc.GetSelectedDrawings().forEach(function(draw) {\n" +
					"\t\tdraw.SetHorPosition(\"" + relative + "\", " + data.value + " * 36000.0, " + data.percent + ");\n" +
					"\t});\n"
			}
		},
		setPositionV			: function(data)
		{
			let relative = "";
			switch(data.relativeFrom)
			{
				 case Asc.c_oAscRelativeFromV.BottomMargin:	relative = "bottomMargin";	break;
				 case Asc.c_oAscRelativeFromV.TopMargin:	relative = "topMargin";		break;
				 case Asc.c_oAscRelativeFromV.Margin:		relative = "margin";		break;
				 case Asc.c_oAscRelativeFromV.Page:			relative = "page";			break;
				 case Asc.c_oAscRelativeFromV.Line:			relative = "line";			break;
				 case Asc.c_oAscRelativeFromV.Paragraph:	relative = "paragraph";		break;
				 default:									relative = "paragraph";
			}

			if (data.useAlign)
			{
				let useAlign = "";
				switch (data.value) {
					case Asc.c_oAscAlignV.Bottom:	useAlign = "bottom";	break;
					case Asc.c_oAscAlignV.Center:	useAlign = "center";	break;
					case Asc.c_oAscAlignV.Top:		useAlign = "top";		break;
					default:						useAlign = "center";
				}

				return "\tdoc.GetSelectedDrawings().forEach(function(draw) {\n" +
					"\tdraw.SetVerAlign(\"" + relative + "\", \"" + useAlign + "\");\n" +
					"\t});\n"

			}

			if (data.percent)
			{
				return "\tdoc.GetSelectedDrawings().forEach(function(draw) {\n" +
       				"\t\tdraw.SetVerPosition(\"" + relative + "\", " + data.value + ", " + data.percent + ");\n" +
       			"\t});\n";
			}
			else
			{
				return "\tdoc.GetSelectedDrawings().forEach(function(draw) {\n" +
					"\t\tdraw.SetVerPosition(\"" + relative + "\", " + data.value + " * 36000.0, " + data.percent + ");\n" +
					"\t});\n";
			}
		},
		setShapeSize			: function(oSize)
		{
			return "\tdoc.GetSelectedDrawings().forEach(function(draw) {\n" +
				"\t\tdraw.SetSize(" + oSize.width + " * 36000.0, " + oSize.height + " * 36000.0);\n" +
				"});\n";

		},
		setDrawingDistances		: function(oDistances)
		{
			return "\tdoc.GetSelectedDrawings().forEach(function(draw) {\n"
				+ "\t\tdraw.SetDistances(" + oDistances.Left + " * 36000.0, " + oDistances.Top + " * 36000.0, " + oDistances.Right + " * 36000.0, " + oDistances.Bottom + " * 36000.0);\n"
				+ "\t});\n";
		},
		setShapeInnerPadding	: function(oPadding)
		{
			let left =	(oPadding.set.Left !== null)	? oPadding.set.Left		: oPadding.original.Left;
			let right =	(oPadding.set.Right !== null)	? oPadding.set.Right	: oPadding.original.Right;
			let top =	(oPadding.set.Top !== null)		? oPadding.set.Top		: oPadding.original.Top;
			let bottom =(oPadding.set.Bottom !== null)	? oPadding.set.Bottom	: oPadding.original.Bottom;

			return "\tdoc.GetSelectedDrawings()\n"
				+ "\t\t.filter(function(item) { return item.GetClassType() === \"shape\"; })\n"
				+ "\t\t.forEach(function(draw) {\n"
				+ "\t\t\tdraw.SetPaddings(" + left + " * 36000.0, " + top + " * 36000.0, " + right + " * 36000.0, " + bottom + " * 36000.0);\n"
				+ "\t\t});\n";
		},
		setShapeRelSizeH		: function(oSize)
		{
			let relative = "";
			switch(oSize.RelativeFrom)
			{
				 case Asc.c_oAscRelativeFromH.Character:	relative = "character";		break;
				 case Asc.c_oAscRelativeFromH.Column:		relative = "column";		break;
				 case Asc.c_oAscRelativeFromH.LeftMargin:	relative = "leftMargin";	break;
				 case Asc.c_oAscRelativeFromH.Margin:		relative = "margin";		break;
				 case Asc.c_oAscRelativeFromH.RightMargin:	relative = "rightMargin";	break;
				 case Asc.c_oAscRelativeFromH.Page:			relative = "page";			break;
				 default:									relative = "page";
			}
			return "\tdoc.GetSelectedDrawings().forEach(function(draw) {\n" +
				"\t\tdraw.SetRelativeWidth(\"" + relative + "\", " + oSize.Value + ");\n" +
				"\t});\n";
		},
		setShapeRelSizeV		: function(oSize)
		{
			let relative = "";
			switch(oSize.RelativeFrom)
			{
				 case Asc.c_oAscRelativeFromV.BottomMargin:	relative = "bottomMargin";		break;
				 case Asc.c_oAscRelativeFromV.Line:			relative = "line";				break;
				 case Asc.c_oAscRelativeFromV.TopMargin:	relative = "topMargin";			break;
				 case Asc.c_oAscRelativeFromV.Margin:		relative = "margin";			break;
				 case Asc.c_oAscRelativeFromV.Paragraph:	relative = "paragraph";			break;
				 case Asc.c_oAscRelativeFromV.Page:			relative = "page";				break;
				 default:									relative = "page";
			}
			return "\tdoc.GetSelectedDrawings().forEach(function(draw) {\n" +
				"\t\tdraw.SetRelativeHeight(\"" + relative + "\", " + oSize.Value + ");\n" +
				"\t});\n";
		},
		setGeometry			: function(type)
		{
			return "\tdoc.GetSelectedDrawings()\n"
				+ "\t\t.filter(function(item) { return item.GetClassType() === \"shape\"; })\n"
				+ "\t\t.forEach(function(draw) {\n"
				+ "\t\t\tdraw.SetGeometry(Api.CreatePresetGeometry(\""+ type +"\"));\n"
				+ "\t\t});\n";
		},
		setDrawingTitle			: function(title){
			return "\tdoc.GetSelectedDrawings().forEach(function(draw) { draw.SetTitle(\"" + title + "\") });\n";
		},
		setDrawingDescription	: function(description){
			return "\tdoc.GetSelectedDrawings().forEach(function(draw) { draw.SetDescription(\"" + description + "\") });\n";
		},
		setDrawingAspectRatio	: function(aspect){
			return "\tdoc.GetSelectedDrawings().forEach(function(draw) { draw.SetLockAspect(" + aspect + ") });\n";
		},
		setTableWidth			: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let oWidth = oAdditional.data;
			let strType = (oWidth.type === 3) ? "percent" : "twips";
			let value = (oWidth.type === 3) ? oWidth.width : AscCommon.MMToTwips(oWidth.width, -1);
			return reg.declaration + "\t" + reg.varName + ".SetWidth(\"" + strType + "\", " + value + ");\n";
		},
		setTableLayout			: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let strLayout = oAdditional.data === c_oAscTableLayout.AutoFit ? "autofit" : "fixed";
			return reg.declaration + "\t" + reg.varName + ".SetTableLayout(\"" + strLayout + "\");\n";
		},
		setTableDefaultMargin	: function(oAdditional){
			let d = oAdditional.data;
			if (!d || !d.Bottom || !d.Left || !d.Right || !d.Top)
				return "";
			let reg = private_getTableVar(oAdditional.tableId);
			let bottom = d.Bottom;
			let top = d.Top;
			let left = d.Left;
			let right = d.Right;
			return reg.declaration +
				(bottom ? "\t" + reg.varName + ".SetTableCellMarginBottom(" + AscCommon.MMToTwips(bottom, -1) + ");\n" : "") +
				(left ? "\t" + reg.varName + ".SetTableCellMarginLeft(" + AscCommon.MMToTwips(left -1) + ");\n" : "") +
				(right ? "\t" + reg.varName + ".SetTableCellMarginRight(" + AscCommon.MMToTwips(right, -1) + ");\n" : "") +
				(top ? "\t" + reg.varName + ".SetTableCellMarginTop(" + AscCommon.MMToTwips(top, -1) + ");\n" : "");
		},
		setCellMargins			: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let d = oAdditional.data;
			let bottom = d.Bottom;
			let top = d.Top;
			let left = d.Left;
			let right = d.Right;
			return reg.declaration + "\t" + reg.varName + ".GetSelectedCells().forEach(function(cell){\n" +
				(bottom ? "\t\tcell.SetCellMarginBottom(" + AscCommon.MMToTwips(bottom, -1) + ");\n" : "") +
				(left ? "\t\tcell.SetCellMarginLeft(" + AscCommon.MMToTwips(left, -1) + ");\n" : "") +
				(right ? "\t\tcell.SetCellMarginRight(" + AscCommon.MMToTwips(right, -1) + ");\n" : "") +
				(top ? "\t\tcell.SetCellMarginTop(" + AscCommon.MMToTwips(top, -1) + ");\n" : "") +
				"\t});\n";
		},
		setTableCellSpacing		: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			return reg.declaration + "\t" + reg.varName + ".SetCellSpacing(" + AscCommon.MMToTwips(oAdditional.data, -1) + ");\n";
		},
		setTableBorders			: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let borders = oAdditional.data;
			let bottom = borders.Bottom;
			let top = borders.Top;
			let left = borders.Left;
			let right = borders.Right;
			let insideH = borders.InsideH;
			let insideV = borders.InsideV;
			let insideHStr = insideH
				? "\t" + reg.varName + ".SetTableBorderInsideH(\"" + (insideH.Value ? 'single' : 'none') + "\", " + insideH.Size * 72 * 8 / 25.4 + ", " + insideH.Space + ", " + insideH.Color.r + ", " + insideH.Color.g + ", " + insideH.Color.b + ");\n"
				: "";
			let insideVStr = insideV
				? "\t" + reg.varName + ".SetTableBorderInsideV(\"" + (insideV.Value ? 'single' : 'none') + "\", " + insideV.Size * 72 * 8 / 25.4 + ", " + insideV.Space + ", " + insideV.Color.r + ", " + insideV.Color.g + ", " + insideV.Color.b + ");\n"
				: "";
			return reg.declaration +
				"\t" + reg.varName + ".SetTableBorderBottom(\"" + (bottom.Value ? 'single' : 'none') + "\", " + bottom.Size * 72 * 8 / 25.4 + ", " + bottom.Space + ", " + bottom.Color.r + ", " + bottom.Color.g + ", " + bottom.Color.b + ");\n" +
				"\t" + reg.varName + ".SetTableBorderLeft(\"" + (left.Value ? 'single' : 'none') + "\", " + left.Size * 72 * 8 / 25.4 + ", " + left.Space + ", " + left.Color.r + ", " + left.Color.g + ", " + left.Color.b + ");\n" +
				"\t" + reg.varName + ".SetTableBorderRight(\"" + (right.Value ? 'single' : 'none') + "\", " + right.Size * 72 * 8 / 25.4 + ", " + right.Space + ", " + right.Color.r + ", " + right.Color.g + ", " + right.Color.b + ");\n" +
				"\t" + reg.varName + ".SetTableBorderTop(\"" + (top.Value ? 'single' : 'none') + "\", " + top.Size * 72 * 8 / 25.4 + ", " + top.Space + ", " + top.Color.r + ", " + top.Color.g + ", " + top.Color.b + ");\n" +
				insideHStr +
				insideVStr;
		},
		setCellBorders			: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let borders = oAdditional.data;
			let bottom = borders.Bottom;
			let top = borders.Top;
			let left = borders.Left;
			let right = borders.Right;
			return reg.declaration + "\t" + reg.varName + ".GetSelectedCells().forEach(function(cell){\n" +
				"\t\tcell.SetCellBorderBottom(\"" + (bottom.Value ? 'single' : 'none') + "\", " + bottom.Size * 72 * 8 / 25.4 + ", " + bottom.Space + ", " + bottom.Color.r + ", " + bottom.Color.g + ", " + bottom.Color.b + ");\n" +
				"\t\tcell.SetCellBorderLeft(\"" + (left.Value ? 'single' : 'none') + "\", " + left.Size * 72 * 8 / 25.4 + ", " + left.Space + ", " + left.Color.r + ", " + left.Color.g + ", " + left.Color.b + ");\n" +
				"\t\tcell.SetCellBorderRight(\"" + (right.Value ? 'single' : 'none') + "\", " + right.Size * 72 * 8 / 25.4 + ", " + right.Space + ", " + right.Color.r + ", " + right.Color.g + ", " + right.Color.b + ");\n" +
				"\t\tcell.SetCellBorderTop(\"" + (top.Value ? 'single' : 'none') + "\", " + top.Size * 72 * 8 / 25.4 + ", " + top.Space + ", " + top.Color.r + ", " + top.Color.g + ", " + top.Color.b + ");\n" +
				"\t});\n";
		},
		setTableAlign			: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let align = oAdditional.data;
			let type;
			if (AscCommon.align_Center === align)
				type = "center";
			else if (AscCommon.align_Right === align)
				type = "right";
			else
				type = "left";
			return reg.declaration + "\t" + reg.varName + ".SetJc(\"" + type + "\");\n";
		},
		setTableBackground		: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let color = oAdditional.data.Color;
			return reg.declaration + "\t" + reg.varName + ".SetBackgroundColor(" + color.r + ", " + color.g + ", " + color.b + ");\n";
		},
		setCellBackground		: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let color = oAdditional.data.Color;
			return reg.declaration + "\t" + reg.varName + ".GetSelectedCells().forEach(function(cell){\n" +
				"\t\tcell.SetBackgroundColor(" + color.r + ", " + color.g + ", " + color.b + ");\n" +
				"\t});\n";
		},
		setCellWidth			: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let cellPrVar = varRegistry.inc('cellPr');
			return reg.declaration
				+ "\tlet " + cellPrVar + " = Api.CreateTableCellPr();\n"
				+ "\t" + cellPrVar + ".SetWidth(\"twips\", " + AscCommon.MMToTwips(oAdditional.data, -1) + ");\n"
				+ "\t" + reg.varName + ".GetSelectedColumnsCells().forEach(function(cell){\n"
				+ "\t\tcell.SetCellPr(" + cellPrVar + ");\n"
				+ "\t});\n";
		},
		setCellWrap				: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let cellPrVar = varRegistry.inc('cellPr');
			return reg.declaration
				+ "\tlet " + cellPrVar + " = Api.CreateTableCellPr();\n"
				+ "\t" + cellPrVar + ".SetNoWrap(" + oAdditional.data + ");\n"
				+ "\t" + reg.varName + ".GetSelectedCells().forEach(function(cell){\n"
				+ "\t\tcell.SetCellPr(" + cellPrVar + ");\n"
				+ "\t});\n";
		},
		setTableInd				: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			return reg.declaration + "\t" + reg.varName + ".SetTableInd(" + AscCommon.MMToTwips(oAdditional.data, -1) + ");\n";
		},
		setTableStyle			: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);

			let style = oMacroRecorderContext.editor.getLogicDocument().GetStyleManager().GetName(oAdditional.data);
			if (!style)
				return "";
			let styleVar = varRegistry.inc('tableStyle');
			return reg.declaration
				+ "\tlet " + styleVar + " = doc.GetStyle(\"" + style + "\");\n"
				+ "\tif (" + styleVar + ") " + reg.varName + ".SetStyle(" + styleVar + ");\n";
		},
		setTableWrappingStyle	: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			return reg.declaration + "\t" + reg.varName + ".SetWrappingStyle(" + !!oAdditional.data + ");\n";
		},
		setRowHeightTable		: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let rowPrVar = varRegistry.inc('rowPr');
			return reg.declaration
				+ "\tlet " + rowPrVar + " = Api.CreateTableRowPr();\n"
				+ "\t" + rowPrVar + ".SetHeight(\"atLeast\", " + AscCommon.MMToTwips(oAdditional.data, -1) + ");\n"
				+ "\t" + reg.varName + ".GetSelectedRows().forEach(function(row){\n"
				+ "\t\trow.SetRowPr(" + rowPrVar + ");\n"
				+ "\t});\n";
		},
		selectTable				: function(data){
			let id = data.id;
			let type = data.type;
			let reg = private_getTableVar(id);

			if (!type || type === c_oAscTableSelectionType.Table) {
				return reg.declaration + "\ttable.Select()\n"
			}
			else if (type === c_oAscTableSelectionType.Row) {
				return reg.declaration + "\ttable.GetSelectedRows().forEach(function(row){console.log(row)})"
			}
			else if (type === c_oAscTableSelectionType.Column) {

			}
			else if (type === c_oAscTableSelectionType.Cell) {

			}
		},
		setTableHPos			: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let relative = oAdditional.data.relative;
			let align = oAdditional.data.align;
			let value = oAdditional.data.value;
			let strRelate;

			if (relative === c_oAscHAnchor.Text)
				strRelate = "character"
			else if (relative === c_oAscHAnchor.Margin)
				strRelate = "margin";
			else
				strRelate = "page";

			return reg.declaration + "\t" + reg.varName + ".SetHorPosition(\"" + strRelate + "\", " + align + ", " + (align ? value : AscCommon.MMToTwips(value, -1)) + ");\n";
		},
		setTableVPos			: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let relative = oAdditional.data.relative;
			let align = oAdditional.data.align;
			let value = oAdditional.data.value;
			let strRelate;

			if (relative === c_oAscVAnchor.Text)
				strRelate = "character"
			else if (relative === c_oAscVAnchor.Margin)
				strRelate = "margin";
			else
				strRelate = "page";

			return reg.declaration + "\t" + reg.varName + ".SetVerPosition(\"" + strRelate + "\", " + align + ", " + AscCommon.MMToTwips(value, -1) + ");\n";
		},
		setTableVAlign			: function(oAdditional)
		{
			let reg = private_getTableVar(oAdditional.tableId);
			let relative = oAdditional.data.relative;
			let value = oAdditional.data.value;
			let strRelate;

			if (relative === c_oAscVAnchor.Text)
				strRelate = "character"
			else if (relative === c_oAscVAnchor.Margin)
				strRelate = "margin";
			else
				strRelate = "page";

			return reg.declaration + "\t" + reg.varName + ".SetVerAlign(\"" + strRelate + "\", \"" + private_GetAlignV(value) + "\");\n";
		},
		setTableHAlign			: function(oAdditional)
		{
			let reg = private_getTableVar(oAdditional.tableId);
			let relative = oAdditional.data.relative;
			let value = oAdditional.data.value;
			let strRelate;

			if (relative === c_oAscVAnchor.Text)
				strRelate = "character"
			else if (relative === c_oAscVAnchor.Margin)
				strRelate = "margin";
			else
				strRelate = "page";

			return reg.declaration + "\t" + reg.varName + ".SetHorAlign(\"" + strRelate + "\", \"" + private_GetAlignH(value) + "\");\n";
		},
		setTableResize		: function(oAdditional)
		{
			let reg = private_getTableVar(oAdditional.tableId);
			let width = AscCommon.MMToTwips(oAdditional.width, -1);
			let height = AscCommon.MMToTwips(oAdditional.height, -1);

			return reg.declaration
				+ "\t" + reg.varName + ".SetWidth(\"twips\", " + width + ");\n"
				//+ reg.varName + ".SetHeight(\"twips\", " + height + ");\n"   // no api for height
		},
	};

	const DEActionMap = {};
	// home tab and general changes
	DEActionMap[AscDFH.historydescription_Document_SetTextBold]				= wordActions.setTextBold;
	DEActionMap[AscDFH.historydescription_Document_SetTextBoldHotKey]			= wordActions.setTextBold;
	DEActionMap[AscDFH.historydescription_Document_SetTextItalic]				= wordActions.setTextItalic;
	DEActionMap[AscDFH.historydescription_Document_SetTextItalicHotKey]		= wordActions.setTextItalic;
	DEActionMap[AscDFH.historydescription_Document_SetTextUnderline]			= wordActions.setTextUnderline;
	DEActionMap[AscDFH.historydescription_Document_SetTextUnderlineHotKey]		= wordActions.setTextUnderline;
	DEActionMap[AscDFH.historydescription_Document_SetTextStrikeout]			= wordActions.setTextStrikeout;
	DEActionMap[AscDFH.historydescription_Document_SetTextStrikeoutHotKey]		= wordActions.setTextStrikeout;
	DEActionMap[AscDFH.historydescription_Document_SetTextFontName]			= wordActions.setTextFontName;
	DEActionMap[AscDFH.historydescription_Document_SetTextFontNameLong]		= wordActions.setTextFontName;
	DEActionMap[AscDFH.historydescription_Document_SetTextFontSize]			= wordActions.setTextFontSize;
	DEActionMap[AscDFH.historydescription_Document_SetTextHighlightColor]		= wordActions.setTextHighlightNone;
	DEActionMap[AscDFH.historydescription_Document_SetTextHighlightNone]		= wordActions.setTextHighlightColor;
	DEActionMap[AscDFH.historydescription_Document_SetTextHighlight]			= wordActions.setTextHighlightColor;
	DEActionMap[AscDFH.historydescription_Document_SetTextVertAlignHotKey2]	= wordActions.setTextVertAlign;
	DEActionMap[AscDFH.historydescription_Document_SetTextVertAlignHotKey3]	= wordActions.setTextVertAlign;
	DEActionMap[AscDFH.historydescription_Document_SetTextVertAlignHotKey]		= wordActions.setTextVertAlign;
	DEActionMap[AscDFH.historydescription_Document_SetTextColor]				= wordActions.setTextColor;
	DEActionMap[AscDFH.historydescription_Document_SetStyleHeading]			= wordActions.setStyleHeading;
	DEActionMap[AscDFH.historydescription_Document_SetParagraphStyle]			= wordActions.setStyleHeading;
	DEActionMap[AscDFH.historydescription_Document_Shortcut_ClearFormatting]	= wordActions.clearFormat;
	DEActionMap[AscDFH.historydescription_Document_ClearFormatting]			= wordActions.clearFormat;
	DEActionMap[AscDFH.historydescription_Cut]									= wordActions.cut;
	DEActionMap[AscDFH.historydescription_Document_ChangeTextCase]				= wordActions.changeTextCase;
	DEActionMap[AscDFH.historydescription_Document_AddLetter]					= wordActions.addLetter;
	DEActionMap[AscDFH.historydescription_Document_SetParagraphAlign]			= wordActions.setAlign;
	DEActionMap[AscDFH.historydescription_Document_SetParagraphAlignHotKey]	= wordActions.setAlign;
	DEActionMap[AscDFH.historydescription_Document_SetParagraphShd]			= wordActions.setParagraphShd;
	DEActionMap[AscDFH.historydescription_Document_SetParagraphLineSpacing]	= wordActions.setLineSpacing;
	//DEActionMap[AscDFH.historydescription_Document_IncParagraphIndent]		= wordActions.incIndentetLineSpacing;
	//DEActionMap[AscDFH.historydescription_Document_DecParagraphIndent]		= wordActions.decIndentetLineSpacing;
	//DEActionMap[AscDFH.historydescription_Document_IncFontSize]				= wordActions.incFontSize;
	DEActionMap[AscDFH.historydescription_Document_SetParagraphNumbering]		= wordActions.setParagraphNumbering;
	DEActionMap[AscDFH.historydescription_Document_SetParagraphNumberingHotKey]= wordActions.setParagraphNumbering;
	DEActionMap[AscDFH.historydescription_Document_AddParagraph]				= wordActions.addParagraph;
	//DEActionMap[AscDFH.historydescription_Document_AddPageNumHotKey]			= wordActions.addPageNum;
	// DEActionMap[AscDFH.historydescription_Document_FormatPasteHotKey]		= wordActions;
	// DEActionMap[AscDFH.historydescription_Document_PasteHotKey]				= wordActions;
	// DEActionMap[AscDFH.historydescription_Document_PasteSafariHotKey]		= wordActions;
	// DEActionMap[AscDFH.historydescription_Document_CutHotKey]				= wordActions;
	DEActionMap['moveCursorLeft']												= wordActions.moveCursorLeft;
	DEActionMap['moveCursorRight']												= wordActions.moveCursorRight;
	DEActionMap['moveCursorUp']												= wordActions.moveCursorUp;
	DEActionMap['moveCursorDown']												= wordActions.moveCursorDown;
	DEActionMap['remove']														= wordActions.remove;

	//DEActionMap[AscDFH.historydescription_Document_BackSpaceButton]			= wordActions.backSpaceButton;
	// DEActionMap[AscDFH.historydescription_Document_DeleteButton]			= wordActions.deleteButton;
	
	// input tab
	DEActionMap[AscDFH.historydescription_Document_AddBlankPage]				= wordActions.addBlankPage;
	DEActionMap[AscDFH.historydescription_Document_AddPageBreak]				= wordActions.addPageBreak;
	DEActionMap[AscDFH.historydescription_Document_AddSectionBreak]			= wordActions.addSectionBreak;
	DEActionMap[AscDFH.historydescription_Document_AddTable]					= wordActions.addTable;
	//DEActionMap[AscDFH.historydescription_Document_ApplyTablePr]				= wordActions.applyTablePr;
	DEActionMap[AscDFH.historydescription_Document_AddChart]					= wordActions.addChart;
	DEActionMap[AscDFH.historydescription_Document_AddImages]					= wordActions.addImage;
	DEActionMap[AscDFH.historydescription_Document_AddHyperlink]				= wordActions.addHyperlink;
	DEActionMap[AscDFH.historydescription_Document_AddNewShape]				= wordActions.addShape;
	DEActionMap['SelectDrawing']												= wordActions.selectDrawing;
	DEActionMap['DeselectDrawing']												= wordActions.deselectDrawing;
	DEActionMap['SetDrawingFill']												= wordActions.setDrawingFill;
	DEActionMap['SetDrawingLine']												= wordActions.setDrawingLine;
	DEActionMap['SetDrawingRotation']											= wordActions.setDrawingRotation;
	DEActionMap['SetDrawingPos']												= wordActions.setDrawingPos;
	DEActionMap['SetDrawingFlipH']												= wordActions.setDrawingFlipH;
	DEActionMap['SetDrawingFlipV']												= wordActions.setDrawingFlipV;
	DEActionMap['SetDrawingWrapping']											= wordActions.setDrawingWrapping;
	DEActionMap['SetPositionH']												= wordActions.setPositionH;
	DEActionMap['SetPositionV']												= wordActions.setPositionV;
	DEActionMap['SetShapeSize']												= wordActions.setShapeSize;
	DEActionMap['SetDrawingDistances']											= wordActions.setDrawingDistances;
	DEActionMap['SetShapeInnerPadding']										= wordActions.setShapeInnerPadding;
	DEActionMap['SetRelSizeH']													= wordActions.setShapeRelSizeH;
	DEActionMap['SetRelSizeV']													= wordActions.setShapeRelSizeV;
	DEActionMap['SetGeometry']													= wordActions.setGeometry;
	DEActionMap[AscDFH.historydescription_Document_RemoveHdrFtr]				= wordActions.removeHdr;
	DEActionMap[AscDFH.historydescription_Document_AddComment]					= wordActions.addComment;
	//DEActionMap[AscDFH.AscDFH.historydescription_Document_AddTextArt]		= wordActions.addTextArt;
	//DEActionMap[AscDFH.AscDFH.historydescription_Document_AddDropCap]		= wordActions.addDropCap;
	//DEActionMap[AscDFH.AscDFH.historydescription_Document_AddDateTimeField]	= wordActions.addDateTimeField;
	DEActionMap[AscDFH.historydescription_Document_AddMath]					= wordActions.addMath;
	DEActionMap[AscDFH.historydescription_Document_AddMathHotKey]				= wordActions.addMathHotkey;
	//DEActionMap[AscDFH.historydescription_Document_AddTextWithProperties]	= wordActions.addMathHotkey;
	DEActionMap[AscDFH.historydescription_Document_AddBlockLevelContentControl]= wordActions.addBlockContentControl;
	DEActionMap[AscDFH.historydescription_Document_AddInlineLevelContentControl]= wordActions.addInlineContentControl;
	DEActionMap[AscDFH.historydescription_Document_AddContentControlList]		= wordActions.addContentControlList;
	DEActionMap[AscDFH.historydescription_Document_AddContentControlCheckBox]	= wordActions.addContentControlCheckBox;
	DEActionMap[AscDFH.historydescription_Document_AddContentControlDatePicker]= wordActions.addContentControlDatePicker;
	DEActionMap[AscDFH.historydescription_Document_AddContentControlPicture]	= wordActions.addContentControlPicture;
	DEActionMap['SetDrawingTitle']												= wordActions.setDrawingTitle;
	DEActionMap['SetDrawingDescription']										= wordActions.setDrawingDescription;
	DEActionMap['SetDrawingAspectRatio']										= wordActions.setDrawingAspectRatio;
	DEActionMap["SetTableWidth"]												= wordActions.setTableWidth;
	DEActionMap["SetTableLayout"]												= wordActions.setTableLayout;
	DEActionMap["SetTableDefaultMargin"]										= wordActions.setTableDefaultMargin;
	DEActionMap["SetTableCellSpacing"]											= wordActions.setTableCellSpacing;
	DEActionMap["SetTableBorders"]												= wordActions.setTableBorders;
	DEActionMap["SetCellBorders"]												= wordActions.setCellBorders;
	DEActionMap["SetTableAlign"]												= wordActions.setTableAlign;
	DEActionMap["SetTableBackground"]											= wordActions.setTableBackground;
	DEActionMap["SetCellBackground"]											= wordActions.setCellBackground;
	DEActionMap["SetCellWidth"]													= wordActions.setCellWidth;
	DEActionMap["SetCellMargins"]												= wordActions.setCellMargins;
	DEActionMap["SetCellWrap"]													= wordActions.setCellWrap;
	DEActionMap["SetTableInd"]													= wordActions.setTableInd;
	DEActionMap["SetTableStyle"]												= wordActions.setTableStyle;
	DEActionMap["SetTableWrappingStyle"]										= wordActions.setTableWrappingStyle;
	DEActionMap["SetRowHeightTable"]											= wordActions.setRowHeightTable;
	DEActionMap["SelectTable"]													= wordActions.selectTable;
	DEActionMap["SetTableHPosition"]											= wordActions.setTableHPos;
	DEActionMap["SetTableVPosition"]											= wordActions.setTableVPos;
	DEActionMap["SetTableVAlign"]												= wordActions.setTableVAlign;
	DEActionMap["SetTableHAlign"]												= wordActions.setTableHAlign;
	DEActionMap["SetTableResize"]												= wordActions.setTableResize;

	const cellActions = {
		setCellIncreaseFontSize	: function(){return "\tApi.GetSelection().FontIncrease();\n"},
		setCellDecreaseFontSize	: function(){return "\tApi.GetSelection().FontDecrease();\n"},
		setCellFontSize			: function(fontSize){return "\tApi.GetSelection().SetFontSize(\"" + fontSize + "\");\n"},
		setCellFontName			: function(fontName){return "\tApi.GetSelection().SetFontName(\"" + fontName + "\");\n"},
		setCellBold				: function(bold){return "\tApi.GetSelection().SetBold(" + bold + ");\n"},
		setCellItalic			: function(italic){return "\tApi.GetSelection().SetItalic(" + italic + ");\n"},
		setCellUnderline		: function(underline){
			let underlineType = null;

			switch (underline) {
				case Asc.EUnderline.underlineSingle:				underlineType = 'single';				break;
				case Asc.EUnderline.underlineSingleAccounting:		underlineType = 'singleAccounting';		break;
				case Asc.EUnderline.underlineDouble:				underlineType = 'double';				break;
				case Asc.EUnderline.underlineDoubleAccounting:		underlineType = 'doubleAccounting';		break;
				case Asc.EUnderline.underlineNone:
				default:											underlineType = 'none';					break;
			}

			return "\tApi.GetSelection().SetUnderline(\"" + underlineType + "\");\n"
		},
		setCellStrikeout		: function(strikeout){return "\tApi.GetSelection().SetStrikeout(" + (!!strikeout) + ");\n"},
		setCellSubscript		: function(subscript){return "\tApi.GetSelection().GetCharacters().GetFont().SetSubscript(" + subscript + ");\n"},
		setCellSuperscript		: function(supscript){return "\tApi.GetSelection().GetCharacters().GetFont().SetSuperscript(" + supscript + ");\n"},
		setCellReadingOrder		: function(dir){
			let direction = null;
			switch (dir) {
				case 0:		direction = 'context';	break;
				case 1:		direction = 'ltr';		break;
				case 2:		direction = 'rtl';		break;
				default:	return "";
			}
			return "\tApi.GetSelection().SetReadingOrder(\"" + direction + "\");\n"
		},
		setCellAlign			: function(al){
			let align = null;
			switch (al) {
				case AscCommon.align_Left:		align = 'left';		break;
				case AscCommon.align_Right:		align = 'right';	break;
				case AscCommon.align_Justify:	align = 'both';		break;
				case AscCommon.align_Center:	align = 'center';	break;
				default:						align = 'center';
			}
			return "\tApi.GetSelection().SetAlignHorizontal(\"" + align + "\");\n"
		},
		setCellVerticalAlign	: function(al){
			let align = null;
			switch (al) {
				case Asc.c_oAscVAlign.Center:	align = 'center';		break;
				case Asc.c_oAscVAlign.Bottom:	align = 'bottom';		break;
				case Asc.c_oAscVAlign.Top:		align = 'top';			break;
				case Asc.c_oAscVAlign.Dist:		align = 'distributed';	break;
				case Asc.c_oAscVAlign.Just:		align = 'justify';		break;
				default:						align = 'center';
			}
			return "\tApi.GetSelection().SetAlignVertical(\"" + align + "\");\n"
		},
		setCellTextColor		: function(clr){
			let color = clr === null
				? "Api.CreateColorFromRGB( 255, 255, 255)"
				: "Api.CreateColorFromRGB(" + clr.getR() + ", " + clr.getG() + ", " + clr.getB() + ")";
			return "\tApi.GetSelection().SetFontColor(" + color + ");\n"
		},
		setCellBackgroundColor	: function(clr){
			let color = clr === null
				? "Api.CreateColorFromRGB( 255, 255, 255)"
				: "Api.CreateColorFromRGB(" + clr.getR() + ", " + clr.getG() + ", " + clr.getB() + ")";
			return "\tApi.GetSelection().SetBackgroundColor(" + color + ");\n"
		},
		setCellWrap				: function(wrap){return "\tApi.GetSelection().SetWrap(" + wrap + ");\n"},
		//setCellShrinkToFit	: function(additional){ return (additional && additional.val !== undefined) ? "Api.GetSelection().SetShrinkToFit(" + additional.val + ");\n" : "";},
		setCellValue			: function(value){
			if (typeof value === 'string')
				value = '"' + value.replace(/"/g, '\\"') + '"';
			else
				value = value.toString();

			return "\tworksheet.GetActiveCell().SetValue(" + value + ");\n"
		},
		setCellFormula 			: function(value){
			if (typeof value === 'string')
				value = '"' + value.replace(/"/g, '\\"') + '"';
			else
				value = value.toString();

			return "\tworksheet.GetActiveCell().SetFormulaArray(" + value + ");\n"
		},
		setCellAngle			: function(angle){
			switch (angle) {
				case -90:	return "\tApi.GetSelection().SetOrientation('xlDownward');\n";
				case 0:		return "\tApi.GetSelection().SetOrientation('xlHorizontal');\n";
				case 90:	return "\tApi.GetSelection().SetOrientation('xlUpward');\n";
				case 255:	return "\tApi.GetSelection().SetOrientation('xlVertical');\n";
			}

			return "";
		},
		setCellChangeTextCase	: function(textCase){return "\tApi.GetSelection().ChangeTextCase(" + textCase + ");\n"},
		setCellChangeFontSize	: function(isInc){
			// todo create api
			return "";
			return isInc ? "\tApi.asc_increaseFontSize();\n" : "\tApi.asc_decreaseFontSize();\n";
		},
		setCellBorder			: function(borderArray){
			if (!Array.isArray(borderArray) || borderArray.length === 0) {
				return "";
			}

			let result = "";

			for (let i = 0; i < borderArray.length; i++) {
				let border = borderArray[i];
				if (border && border.style !== undefined) {

					let positionStr = null;
					switch (i) {
						case 0: positionStr = 'Top'; break;
						case 1: positionStr = 'Right'; break;
						case 2: positionStr = 'Bottom'; break;
						case 3: positionStr = 'Left'; break;
						case 4: positionStr = 'DiagonalDown'; break;
						case 5: positionStr = 'DiagonalUp'; break;
						case 6: positionStr = 'InsideVertical'; break;
						case 7: positionStr = 'InsideHorizontal'; break;
						default: continue;
					}

					let styleStr = null;
					switch (border.style) {
						case window['Asc'].c_oAscBorderStyles.None: styleStr = 'None'; break;
						case window['Asc'].c_oAscBorderStyles.Double: styleStr = 'Double'; break;
						case window['Asc'].c_oAscBorderStyles.Hair: styleStr = 'Hair'; break;
						case window['Asc'].c_oAscBorderStyles.DashDotDot: styleStr = 'DashDotDot'; break;
						case window['Asc'].c_oAscBorderStyles.DashDot: styleStr = 'DashDot'; break;
						case window['Asc'].c_oAscBorderStyles.Dotted: styleStr = 'Dotted'; break;
						case window['Asc'].c_oAscBorderStyles.Dashed: styleStr = 'Dashed'; break;
						case window['Asc'].c_oAscBorderStyles.Thin: styleStr = 'Thin'; break;
						case window['Asc'].c_oAscBorderStyles.MediumDashDotDot: styleStr = 'MediumDashDotDot'; break;
						case window['Asc'].c_oAscBorderStyles.SlantDashDot: styleStr = 'SlantDashDot'; break;
						case window['Asc'].c_oAscBorderStyles.MediumDashDot: styleStr = 'MediumDashDot'; break;
						case window['Asc'].c_oAscBorderStyles.MediumDashed: styleStr = 'MediumDashed'; break;
						case window['Asc'].c_oAscBorderStyles.Medium: styleStr = 'Medium'; break;
						case window['Asc'].c_oAscBorderStyles.Thick: styleStr = 'Thick'; break;
						default: continue;
					}

					let colorStr = "Api.CreateColorFromRGB(0, 0, 0)";
					if (border.color) {
						if (typeof border.color === 'string') {
							let hex = border.color.replace('#', '');
							if (hex.length === 3) {
								hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
							}
							let r = parseInt(hex.substr(0, 2), 16) || 0;
							let g = parseInt(hex.substr(2, 2), 16) || 0;
							let b = parseInt(hex.substr(4, 2), 16) || 0;
							colorStr = "Api.CreateColorFromRGB(" + r + ", " + g + ", " + b + ")";
						} else if (typeof border.color === 'object') {
							colorStr = "Api.CreateColorFromRGB(" + (border.color.r || 0) + ", " + (border.color.g || 0) + ", " + (border.color.b || 0) + ")";
						}
					}

					result += "\tApi.GetSelection().SetBorders(\"" + positionStr + "\", \"" + styleStr + "\", " + colorStr + ");\n";
				}
			}

			return result;
		},
		// setCellHyperlinkAdd		: function(additional) {return (additional && additional.url) ? "" : ""},
		// setCellHyperlinkModify	: function(additional) {return (additional && additional.url) ? "" : ""},
		// setCellHyperlinkRemove	: function(additional) {return (additional && additional.url) ? "" : ""},
		// cut						: function(){return "ApiApi.GetSelection().Cut();\n"},
		setCellStyle			: function(style){return ""},
		setCellFormat			: function(format){
			return "\tlet " + varRegistry.inc('format') + " = Api.Format(worksheet.GetActiveCell().GetValue(), \'" + format + "\')\n"
			+ "\tworksheet.GetActiveCell().SetValue(" + varRegistry.get('format') + ");\n";
		},
		setCellHyperlinkRemove	: function(data){return ""},
		setCellMerge			: function(data){
			if (data === Asc.c_oAscMergeOptions.MergeCenter)
				return "\tApi.GetSelection().Merge(false);\n"; // + set shrink / indent
			else if (data === Asc.c_oAscMergeOptions.None)
				return "\tApi.GetSelection().UnMerge();\n";
			else if (data === Asc.c_oAscMergeOptions.MergeAcross)
				return "\tApi.GetSelection().Merge(true);\n";
			else if (data === Asc.c_oAscMergeOptions.Merge)
				return "\tApi.GetSelection().Merge(false);\n";
		},
		setCellSort				: function(obj){
			let range = "\tlet " + varRegistry.inc('range') + " = Api.GetSelection().GetAddress(true, true);\n";

			if (obj.type === Asc.c_oAscSortOptions.Ascending)
				range += "\tApi.GetSelection().SetSort(" + varRegistry.get('range') + " , \"xlAscending\", undefined, undefined, undefined, undefined, \"xlYes\");\n";
			else if (obj.type === Asc.c_oAscSortOptions.Descending)
				range += "\tApi.GetSelection().SetSort(" + varRegistry.get('range') + " , \"xlDescending\", undefined, undefined, undefined, undefined, \"xlYes\");\n";

			return range;
		},
		setCellEmpty			: function(){ return "\tApi.GetSelection().Clear();\n";},
		setNumberFormat			: function(format){ return "\tApi.GetSelection().SetNumberFormat(\"" + format + "\");\n";},
		setCellPaste			: function(){return "\tApi.GetSelection().Paste();\n";},
		addChart				: function(chart){ //todo title
			let range = chart.parent.dataRefs.getRange();
			let type = private_ChartInternalTypeToBuilder(chart.getChartType());
			let nStyle = chart.getChartSpace().style;
			let from = chart.parent.drawingBase.from;
			let x = chart.parent.extX;
			let y = chart.parent.extY;
			return "\tworksheet.AddChart(\""+ range + "\", true, \"" + type + "\", " + nStyle + ", " + x + " * 36000, " + y + " * 36000, " + from.col + ", " + from.colOff + " * 36000, " + from.row + ", " + from.rowOff + " * 36000);\n"
		},
		addComment				: function(comment){
			let col		= comment.coords.nCol;
			let row		= comment.coords.nRow;
			let time	= comment.sTime;
			let rangeVar   = varRegistry.inc('range');
			let commentVar = varRegistry.inc('comment');

			return "\tlet " + rangeVar + " = worksheet.GetRangeByNumber(" + row + ", " + col + ");\n"
				+ "\tlet " + commentVar + " = " + rangeVar + ".AddComment(\""+ comment.sText + "\");\n"
				+ "\t" + commentVar + ".SetAuthorName(\"" + comment.sUserName +"\");\n"
				+ "\t" + commentVar + ".SetUserId(\"" + comment.sUserId +"\");\n"
				+ "\t" + commentVar + ".SetTime(" + time +");\n"
		},
		addHyperlink			: function(hp){
			let box		= hp.hyperlinkModel.Ref.bbox;
			let name = box.getName().split(":")[0];
			let loc = hp.hyperlinkModel.LocationSheet + "!" + hp.hyperlinkModel.LocationRangeBbox.getAbsName();
			if (hp.hyperlinkModel.Location)
				return "\tworksheet.SetHyperlink(\"" + name + "\", \"" + loc + "\", \"" + loc +"\", \"" + (hp.hyperlinkModel.Tooltip ? hp.hyperlinkModel.Tooltip : "") + "\", \"" + (hp.text ? hp.text : "") + "\");\n"
			else if (hp.hyperlinkModel.Hyperlink)
				return "\tworksheet.SetHyperlink(\"" + name + "\", \"" + hp.hyperlinkModel.Hyperlink + "\", \"\", \"" + (hp.hyperlinkModel.Tooltip ? hp.hyperlinkModel.Tooltip : "") + "\", \"" + (hp.text ? hp.text : "") + "\");\n"
		},
		addImageUrls			: function(image){
			let from = image.from;
			return "\tworksheet.AddImage(\"" + image.src + "\", " + image.width + " * 36000, " + image.height + " * 36000, "+ from.col + ", " + from.colOff + " * 36000, " + from.row + ", " + from.rowOff + " * 36000);\n"
		},
		addAutoFilter			: function(data){
			return "\tApi.GetSelection().SetAutoFilter();\n";
		},
		removeAutoFilter		: function(){
			return "";
		},
		selectRange				: function(ar){
			// todo update document on select
			let selectName = ar.getName();
			return "\tApi.GetRange(\""+ selectName + "\").Select();\n";
		},
		addShape				: function(shapeProps){
			let fill = shapeProps.fill.getRGBAColor();
			let border = shapeProps.border;
			let borderwidth = border.w / 36000;
			let borderColor = border.Fill.getRGBAColor();
			let varName = varRegistry.register(shapeProps.id, 'shape');
			let from = shapeProps.base.from;
			return "\tlet " + varName + " = (function () {\n" +
					"\t\tlet fill = Api.CreateSolidFill(Api.CreateRGBColor("+ fill.R +", " + fill.G + ", " + fill.B + "));\n" +
					"\t\tlet stroke = Api.CreateStroke(" + borderwidth +" * 36000, Api.CreateSolidFill(Api.CreateRGBColor("+ borderColor.R +", " + borderColor.G + ", " + borderColor.B + ")));\n" +
					"\t\tlet shape = worksheet.AddShape(\"" + shapeProps.type + "\", " + shapeProps.extX + " * 36000, " + shapeProps.extY + " * 36000, fill, stroke, " + from.col + ", " + from.colOff * 36000 + ", " + from.row + ", " + from.rowOff * 36000 + ");\n" +
					"\t\tshape.Select();\n" +
					"\t\treturn shape\n" +
				"\t}());\n";
		},
		deselectDrawing				: function(oAdditional)
		{
			let varName = varRegistry.get(oAdditional.name);
			if (varName === undefined) {
				return "\tworksheet.GetSelectedDrawings().forEach(function(drawing){drawing.Unselect()});\n";
			}
			return "\t[" + varName + "]\n\t\t.forEach(function(drawing){drawing.Unselect()});\n"
		},
		selectDrawing				: function(shapes)
		{
			let str = "";
			let shapesSort = [];

			shapes.forEach(function(element) {
				if (!shapesSort.includes(element))
					shapesSort.push(element);
			});

			shapesSort = shapesSort.filter(function(element) {
				return varRegistry.get(element.getObjectName()) !== undefined;
			});

			if (shapesSort.length === 0)
				return "";

			let ids = "";
			for (let i = 0; i < shapesSort.length; i++) {
				let varName = varRegistry.get(shapesSort[i].getObjectName());
				if (varName !== undefined)
					ids += varName + (i != (shapesSort.length - 1) && shapesSort.length > 1 ? ", " : "")
			}

			return str + "\t[" + ids + "]\n\t\t.forEach(function(drawing, index){drawing.Select(index === 0)});\n"
		},
		setDrawingFill			: function(unifill)
		{
			//TODO add transparent when add api
			if (unifill.fill.type === Asc.c_oAscFill.FILL_TYPE_SOLID)
			{
				let color = unifill.fill.color.color.RGBA;
				//let transparent = unifill.transparent;

				return "\tworksheet.GetSelectedShapes()\n"
						+ "\t\t.forEach(shape => shape.Fill(Api.CreateSolidFill(Api.CreateRGBColor(" + color.R + ", " + color.G + ", " + color.B + "))));\n";
			}
			else if (unifill.fill instanceof AscFormat.CGradFill)
			{
				let colorScheme = Asc.editor.getCurrentColorScheme();

				let colors = unifill.fill.colors;
				let strColor = "";

				for(let nColor = 0; nColor < colors.length; nColor++)
				{
					let gs = colors[nColor];
					let color = gs.color.color.id !== undefined ? colorScheme.colors[gs.color.color.id].RGBA : gs.color.color.RGBA;

					strColor += (nColor !== 0)
						? ",\n\t\t\t\t"
						: "\n\t\t\t\t";

					strColor += "Api.CreateGradientStop(Api.CreateRGBColor(" + color.R + ", " + color.G + ", " + color.B + "), " + gs.pos + ")";
				}

				let gradient = "";
				if (unifill.fill.lin) // leinear
					gradient = "\n\t\t\tApi.CreateLinearGradientFill([" + strColor + "\n\t\t\t], " + unifill.fill.lin.angle + ")";
				else if (unifill.fill.path) // radial
					gradient = "\n\t\t\tApi.CreateRadialGradientFill([" + strColor + "\n\t\t\t])";

				return "\tworksheet.GetSelectedShapes()\n"
						+ "\t\t.forEach(shape => shape.Fill(" + gradient + ")\n\t\t);\n";

			}
			else if (unifill.fill instanceof AscFormat.CBlipFill)
			{
				let base64data = unifill.fill.getBase64RasterImageId()
				let blipFill = "\n\t\t\tApi.CreateBlipFill(\"" + base64data + "\", \"" + (unifill.fill.tile ? "tile" : "stretch") + "\")";
				return "\tworksheet.GetSelectedShapes()\n"
						+ "\t\t.forEach(shape => shape.Fill(" + blipFill + ")\n\t\t);\n";

			}
			else if (unifill.fill instanceof AscFormat.CPattFill)
			{
				let bgClr = unifill.fill.bgClr.color.RGBA;
				let fgClr = unifill.fill.fgClr.color.RGBA;
				let ftype = unifill.fill.ftype;
				let type = AscCommon.global_hatch_names[ftype];

				let blipFill = "\n\t\t\tApi.CreatePatternFill(\n\t\t\t\t\"" + type + "\",\n\t\t\t\t" + "Api.CreateRGBColor(" + bgClr.R + ", " + bgClr.G + ", " + bgClr.B + "),\n\t\t\t\tApi.CreateRGBColor(" + fgClr.R + ", " + fgClr.G + ", " + fgClr.B + ")";
				return "\tworksheet.GetSelectedShapes()\n"
						+ "\t\t.forEach(shape => shape.Fill(" + blipFill + ")\n\t\t\t)\n\t\t);\n";
			}
			else if (unifill.fill instanceof AscFormat.CNoFill)
			{
				return "\tworksheet.GetSelectedShapes()\n"
						+ "\t\t.forEach(shape => shape.Fill(Api.CreateNoFill()));\n";
			}
		},
		setDrawingLine			: function(line)
		{
			let strStrokeShape = "\t(function () {\n";
			let type = AscFormat.CLn.prototype.GetDashByCode(line.prstDash);
			let color = line.Fill.fill.color.color.RGBA;

			strStrokeShape += "\t\tlet stroke = Api.CreateStroke(\n" +
					"\t\t\t" + line.w / 12700.0 + " * 12700.0,\n" +
					"\t\t\t" + "Api.CreateSolidFill(Api.CreateRGBColor(" + color.R + ", " + color.G + ", " + color.B + ")),\n" +
					"\t\t\t\"" + type + "\"\n" +
				"\t\t);\n";

			strStrokeShape += "\t\tworksheet.GetSelectedShapes()\n"
				+ "\t\t\t.forEach(shape => shape.SetOutLine(stroke));\n";

			strStrokeShape += "\t}());\n";

			return strStrokeShape;
		},
		setDrawingRotation		: function(nRot)
		{
			return "\tworksheet.GetSelectedShapes()\n"
				+ "\t\t.forEach(shape => {shape.SetRotation(" + nRot * 180 / Math.PI + ")});\n";

		},
		setDrawingFlipH			: function(isFlip)
		{
			return ""
			return "\tworksheet.GetSelectedShapes()\n"
				+ "\t\t.forEach(shape => {shape.SetHorFlip(" + isFlip +")});\n";

		},
		setDrawingFlipV			: function(isFlip)
		{
			return ""
			return "\tworksheet.GetSelectedShapes()\n"
				+ "\t\t.forEach(shape => {shape.SetVertFlip(" + isFlip + ")});\n";

		},
		setDrawingPos			: function(obj)
		{
			let fromCol	= obj.from.col;
			let colOff	= obj.from.colOff;
			let fromRow	= obj.from.row;
			let rowOff	= obj.from.rowOff;

			return "\tworksheet.GetSelectedDrawings().forEach(draw => {\n" +
				"\t\tdraw.SetPosition(" + fromCol + ", " + colOff + " * 36000, " + fromRow + ", " + rowOff + " * 36000);\n" +
				"\t});\n"
		},
		setShapeSize			: function(oSize)
		{
			return "\tworksheet.GetSelectedDrawings().forEach(draw => {" +
				"draw.SetSize(" + oSize.width + " * 36000.0, " + oSize.height + " * 36000.0)" +
			"});\n"
		},
		setShapeInnerPadding	: function(oPadding)
		{
			let left =	(oPadding.set.Left !== null)	? oPadding.set.Left		: oPadding.original.Left;
			let right =	(oPadding.set.Right !== null)	? oPadding.set.Right	: oPadding.original.Right;
			let top =	(oPadding.set.Top !== null)		? oPadding.set.Top		: oPadding.original.Top;
			let bottom =(oPadding.set.Bottom !== null)	? oPadding.set.Bottom	: oPadding.original.Bottom;

			return "\tworksheet.GetSelectedShapes()\n"
				+ "\t\t.forEach(draw => {\n"
					+ "\t\t\tdraw.SetPaddings(" + left + " * 36000.0, " + top + " * 36000.0, " + right + " * 36000.0, " + bottom + " * 36000.0)\n"
				+ "\t\t});\n"

		},
		setGeometry			: function(type)
		{
			return "\tworksheet.GetSelectedShapes()\n"
				+ "\t\t.forEach(draw => {\n"
					+ "\t\t\tdraw.SetGeometry(Api.CreatePresetGeometry(\""+ type +"\"))\n"
				+ "\t\t});\n"
		},
		setDrawingTitle		: function(title){
			return "\tworksheet.GetSelectedDrawings().forEach(function(draw) { draw.SetTitle(\"" + title + "\") });\n";
		},
		setDrawingDescription	: function(description){
			return "\tworksheet.GetSelectedDrawings().forEach(function(draw) { draw.SetDescription(\"" + description + "\") });\n";
		},
		setDrawingAspectRatio	: function(aspect){
			return "\tworksheet.GetSelectedDrawings().forEach(function(draw) { draw.SetLockAspect(" + aspect + ") });\n";
		}
	};
	const SEActionMap = {};
	//SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellIncreaseFontSize]	= cellActions.setCellIncreaseFontSize,
	//SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellDecreaseFontSize]	= cellActions.setCellDecreaseFontSize,
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellFontSize]				= cellActions.setCellFontSize;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellFontName]				= cellActions.setCellFontName;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellBold]					= cellActions.setCellBold;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellItalic]				= cellActions.setCellItalic;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellUnderline]			= cellActions.setCellUnderline;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellStrikeout]			= cellActions.setCellStrikeout;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellSubscript]			= cellActions.setCellSubscript;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellSuperscript]			= cellActions.setCellSuperscript;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellReadingOrder]			= cellActions.setCellReadingOrder;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellAlign]				= cellActions.setCellAlign;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellVertAlign]			= cellActions.setCellVerticalAlign;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellTextColor]			= cellActions.setCellTextColor;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellBackgroundColor]	    = cellActions.setCellBackgroundColor;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellWrap]				    = cellActions.setCellWrap;
	//SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellShrinkToFit]		= cellActions.setCellShrinkToFit;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellBorder]				= cellActions.setCellBorder;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellValue]				= cellActions.setCellValue;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellAngle]				= cellActions.setCellAngle;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellMerge]				= cellActions.setCellMerge;
	//SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellStyle]				= cellActions.setCellStyle;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellChangeTextCase]		= cellActions.setCellChangeTextCase;
	//SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellChangeFontSize]		= cellActions.setCellChangeFontSize;
	//SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellHyperlinkAdd]		= cellActions.setCellHyperlinkAdd;
	//SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellHyperlinkModify]	= cellActions.setCellHyperlinkModify;
	//SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellHyperlinkRemove]	= cellActions.setCellHyperlinkRemove;
	//SEActionMap[AscDFH.historydescription_Cut]									= cellActions.cut;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellFormat]				= cellActions.setCellFormat;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellSort]					= cellActions.setCellSort;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellEmpty]				= cellActions.setCellEmpty;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellChangeDigNum]			= cellActions.setNumberFormat;
	//SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellPaste]				= cellActions.setCellPaste;
	SEActionMap[AscDFH.historydescription_Spreadsheet_AddShape]					= cellActions.addShape;
	SEActionMap[AscDFH.historydescription_Spreadsheet_AddChart]					= cellActions.addChart;
	SEActionMap[AscDFH.historydescription_Spreadsheet_AddComment]					= cellActions.addComment;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellHyperlink]			= cellActions.addHyperlink;
	SEActionMap[AscDFH.historydescription_Spreadsheet_AddImageUrls]				= cellActions.addImageUrls;
	SEActionMap[AscDFH.historydescription_Spreadsheet_AddAutoFilter]				= cellActions.addAutoFilter;
	SEActionMap[AscDFH.historydescription_Spreadsheet_RemoveAutoFilter]			= cellActions.removeAutoFilter;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SelectRange]					= cellActions.selectRange;
	SEActionMap[AscDFH.historydescription_Spreadsheet_SetCellFormula]				= cellActions.setCellFormula;

	SEActionMap['DeselectDrawing']													= cellActions.deselectDrawing;
	SEActionMap['SelectDrawing']													= cellActions.selectDrawing;
	SEActionMap['SetDrawingFill']													= cellActions.setDrawingFill;
	SEActionMap['SetDrawingLine']													= cellActions.setDrawingLine;
	SEActionMap['SetDrawingRotation']												= cellActions.setDrawingRotation;
	SEActionMap['SetDrawingFlipH']													= cellActions.setDrawingFlipH;
	SEActionMap['SetDrawingFlipV']													= cellActions.setDrawingFlipV;
	SEActionMap['SetDrawingPos']													= cellActions.setDrawingPos;
	SEActionMap['SetShapeSize']													= cellActions.setShapeSize;
	SEActionMap['SetShapeInnerPadding']											= cellActions.setShapeInnerPadding;
	SEActionMap['SetGeometry']														= cellActions.setGeometry;
	SEActionMap['SetDrawingTitle']													= cellActions.setDrawingTitle;
	SEActionMap['SetDrawingDescription']											= cellActions.setDrawingDescription;
	SEActionMap['SetDrawingAspectRatio']											= cellActions.setDrawingAspectRatio;

	const presActions = {
		setParagraphAlign		: function(align){
			switch (align) {
				case AscCommon.align_Left:		align = 'left';		break;
				case AscCommon.align_Right:		align = 'right';	break;
				case AscCommon.align_Justify:	align = 'justify';	break;
				case AscCommon.align_Center:	align = 'center';	break;
				default:						align = 'center';
			}
			return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
				"\t\tshape.GetDocContent().GetContent().forEach(function(para) {\n" +
				"\t\t\tpara.SetJc(\"" + align + "\");\n" +
				"\t\t});\n" +
				"\t});\n";
		},
		paragraphAdd			: function(additional){
			if (!additional.length)
				return "";

			let text = "";
			for (let nChar = 0; nChar < additional.length; nChar++)
				text += String.fromCodePoint(additional[nChar]);

			return "\tpresentation.EnterText(\"" + text + "\");\n"	
		},
		remove					: function(){
			if (oMacroRecorderContext.prevChangeType === AscDFH.historydescription_Presentation_ParagraphAdd && oMacroRecorderContext.prevData.length)
				oMacroRecorderContext.prevData.pop();
			return "";
		},
		putTextPrBold			: function(bold){return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
			"\t\tshape.GetDocContent().GetContent().forEach(function(para) {\n" +
			"\t\t\tpara.SetBold(" + bold + ");\n" +
			"\t\t});\n" +
			"\t});\n"
			},
		putTextPrItalic			: function(italic){return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
			"\t\tshape.GetDocContent().GetContent().forEach(function(para) {\n" +
			"\t\t\tpara.SetItalic(" + italic + ");\n" +
			"\t\t});\n" +
			"\t});\n"},
		putTextPrUnderline		: function(underline){return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
			"\t\tshape.GetDocContent().GetContent().forEach(function(para) {\n" +
			"\t\t\tpara.SetUnderline(" + underline + ");\n" +
			"\t\t});\n" +
			"\t});\n"
			},
		putTextPrStrikeout		: function(strikeout){return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
			"\t\tshape.GetDocContent().GetContent().forEach(function(para) {\n" +
			"\t\t\tpara.SetStrikeout(" + strikeout + ");\n" +
			"\t\t});\n" +
			"\t});\n"
			},
		putTextPrFontName		: function(fontName){return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
			"\t\tshape.GetDocContent().GetContent().forEach(function(para) {\n" +
			"\t\t\tpara.SetFontName(" + fontName + ");\n" +
			"\t\t});\n" +
			"\t});\n"
			},
		putTextPrFontSize		: function(fontsize){return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
			"\t\tshape.GetDocContent().GetContent().forEach(function(para) {\n" +
			"\t\t\tpara.SetFontSize(" + fontsize + ");\n" +
			"\t\t});\n" +
			"\t});\n"
			},
		//putTextPrIncreaseFontSize : function(){return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.SetFontSize(" + fontsize + "));\n\t})\n"},
		//incDecFontSize			: makeAction("", function(){return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.SetFontSize(" + fontsize + "));\n\t})\n"}),
		setTextVertAlign		: function(vertalign){
			let textOfVertAlign = "baseline";

			if (AscCommon.vertalign_Baseline === vertalign)
				textOfVertAlign = "baseline";
			else if (AscCommon.vertalign_SubScript === vertalign)
				textOfVertAlign = "subscript";
			else if (AscCommon.vertalign_SuperScript === vertalign)
				textOfVertAlign = "superscript";

			return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
				"\t\tshape.GetDocContent().GetContent().forEach(function(para) {\n" +
				"\t\t\tpara.SetVertAlign(\"" + textOfVertAlign + "\");\n" +
				"\t\t});\n" +
				"\t});\n"
		},
		addNextSlide			: function(data){
			if (data === undefined)
			{
				return "\tpresentation.AddSlide(Api.CreateSlide());\n"
			}
			else
			{
				return "\tlet " + varRegistry.inc('slide') +" = Api.CreateSlide();\n"
					+ "\tlet " + varRegistry.inc('master') +" = presentation.GetMaster(0);\n"
					+ "\tlet " + varRegistry.inc('layout') +" = " + varRegistry.get('master') + ".GetLayout(" + data + ");\n"
					+ "\t" + varRegistry.get('slide') + ".ApplyLayout(" + varRegistry.get('layout') + ");\n"
					+ "\tpresentation.AddSlide(" + varRegistry.get('slide') + ");\n";
			}
		},
		deleteSlides			: function(index){
			return "\tlet " + varRegistry.inc('slide') +" = presentation.GetSlideByIndex(" + index + ");\n"
				+ "\tif (" + varRegistry.get('slide') + ") " + varRegistry.get('slide') + ".Delete();\n";
		},
		changeLayout			: function(changeObj) {
			return "\t[" + changeObj.slides.toString() + "].forEach(function(index) {\n"
				+ "\t\tlet " + varRegistry.inc('slide') +" = presentation.GetSlideByIndex(index);\n"
				+ "\t\tlet " + varRegistry.inc('master') +" = presentation.GetMaster(0);\n"
				+ "\t\tlet " + varRegistry.inc('layout') +" = " + varRegistry.get('master') + ".GetLayout(" + changeObj.layout + ");\n"
				+ "\t\tif (" + varRegistry.get('slide') + ") " + varRegistry.get('slide') + ".ApplyLayout(" + varRegistry.get('layout') + ");\n"
				+ "\t});\n";
		},
		//showfrom				: function(){},
		setTextHighlight		: function(highlight){
			let highlightColor = "";
			if (highlight)
			{
				let color = new AscCommonWord.CDocumentColor(highlight.r, highlight.g, highlight.b);
				highlightColor = color.ToHighlightColor();
			}

			if (highlightColor === "") highlightColor = 'none';
			return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
				"\t\tshape.GetDocContent().GetContent().forEach(function(para) {\n" +
				"\t\t\tpara.SetHighlight(\"" + highlightColor + "\");\n" +
				"\t\t});\n" +
				"\t});\n";

		},
		putTextColor			: function(color){
			return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
				"\t\tshape.GetDocContent().GetContent().forEach(function(para) {\n" +
				"\t\t\tpara.SetColor(" + color.r + ", " + color.g + ", " + color.b + ");\n" +
				"\t\t});\n" +
				"\t});\n";
		},
		clearFormatting			: function(isClear){return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
			"\t\tshape.GetDocContent().GetContent().forEach(function(para) {\n" +
			"\t\t\tpara.ClearFormating(" + isClear + ");\n" +
			"\t\t});\n" +
			"\t});\n"
		},
		putTextPrLineSpacing	: function(lineSpacing){
			let type = lineSpacing.type;
			let value = lineSpacing.value;

			switch(type)
			{
				case Asc.linerule_Auto		: type = "auto";	break;
				case Asc.linerule_AtLeast	: type = "atLeast";	break;
				case Asc.linerule_Exact		: type = "exact";	break;
				default						: type = "auto";	break;
			}
			return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
				"\t\tshape.GetDocContent().GetContent().forEach(function(para) {\n" +
				"\t\t\tpara.SetSpacingLine(" + value + " * 240, \"" + type + "\");\n" +
				"\t\t});\n" +
				"\t});\n";
		},
		paragraphRemove			: function(args){
			return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
				"\t\tshape.GetDocContent().GetContent().forEach(function(para) {\n" +
				"\t\t\tpara.RemoveAllElements();\n" +
				"\t\t});\n" +
				"\t});\n";
		},
		setVerticalAlign		: function(align){
			let typeOfVertAlign = "";
			switch(align.verticalTextAlign)
            {
                case 4:
                {
                    typeOfVertAlign = "top";
                    break;
                }
                case 1:
                {
					typeOfVertAlign = "center";
                    break;
                }
                case 0:
                {
					typeOfVertAlign = "bottom";
                    break;
                }
            }

			return "\tApi.GetSelection().GetShapes().forEach(function(shape) {\n" +
				"\t\tshape.SetVerticalTextAlign(\"" + typeOfVertAlign + "\");\n" +
				"\t});\n";
		},
		bringForward			: function(){
			// no api
		},
		bringToFront			: function(){
			// no api
		},
		bringBackward			: function(){
			// no api
		},
		sendToBack				: function(){
			// no api
		},
		createGroup				: function(){
			return "\tpresentation.GetCurrentSlide().GroupDrawings(Api.GetSelection().GetShapes());\n";
		},
		unGroup					: function(){
			return "\tApi.GetSelection().GetShapes().forEach(function(shape) { shape.Ungroup(); });\n";
		},
		addFlowTable			: function(prop){
			let varName = prop.tableId
				? varRegistry.register(prop.tableId, "table")
				: varRegistry.inc('table');
			let width = prop.tableGrid.map(function(num) { return AscCommon.MMToTwips(num, -1) });
			let styleStr = prop.style ? prop.style : "";
			let styleVar = varRegistry.inc('tableStyle');

			return "\tlet " + varName + " = (function()\n"
				+ "\t{\n"
				+ "\t\tlet table = Api.CreateTable(" + prop.col + ", " + prop.row + ");\n"
				+ "\t\tlet " + styleVar + " = presentation.GetStyle(\"" + styleStr + "\");\n"
				+ "\t\tif (" + styleVar + ") table.SetStyle(" + styleVar + ");\n"
				+ "\t\tlet row = table.GetRow(0);\n"
				+ "\t\t[" + width.toString() + "].forEach(function(width, index){\n"
				+ "\t\t\tlet cellColumnFirst = row.GetCell(index);\n"
				+ "\t\t\tlet cellPr = Api.CreateTableCellPr();\n"
				+ "\t\t\tcellPr.SetWidth(\"twips\", width);\n"
				+ "\t\t\tcellColumnFirst.SetCellPr(cellPr);\n"
				+ "\t\t});\n"
				+ "\t\tpresentation.GetCurrentSlide().AddObject(table);\n"
				//+ "\t\ttable.GetRange(0, 0).Select();\n"
				+ "\t\ttable.SetPosition(" + prop.x + " * 36000.0, " + prop.y + " * 36000.0);\n"
				+ "\t\treturn table;\n"
				+ "\t}());\n";
		},
		addFlowImage			: function(image){
			let text = "";
			for (let i = 0; i < image.src.length; i++)
			{
				let curImageUrl = image.src[i].src;
				let size = image.data[i];
				let xfrm = size.getXfrm();

				let width = xfrm.extX;
				let height = xfrm.extY;

				let posX = xfrm.offX;
				let posY = xfrm.offY;

				text += "\tlet " + varRegistry.inc('image') + " = Api.CreateImage(\"" + curImageUrl + "\", " + width + " * 36000, " + height + " * 36000);\n" +
					"\t" + varRegistry.get('image') + ".SetPosition(" + posX + " * 36000, " + posY + " * 36000);\n" +
					"\tpresentation.GetCurrentSlide().AddObject(" + varRegistry.get('image') + ");\n"
			}
			return text;
		},
		addChart				: function(chart){ //todo title
			let series = chart.getAllSeries();
			let seriesNames = [];
			let seriesData = [];
			let numformat	= [];
			for (let i = 0; i < series.length; i++)
			{
				let currSer = series[i];
				let name = currSer.asc_getSeriesName();
				let value = currSer.asc_getValuesArr();
				let format = currSer.getCatSourceNumFormat();
				seriesNames.push(name);
				seriesData.push(value);
				numformat.push(format);
			}

			let categories	= chart.getCatValues();
			let chartType	= private_ChartInternalTypeToBuilder(chart.getChartType());
			let width		= chart.GetWidth() * 36000.0;		//mm2emu
			let height		= chart.GetHeight() * 36000.0;		//mm2emu
			let style		= chart.getChartStyleIdx();
			let title		= chart.getChartTitle().getDocContent().GetText();
			title = title ? title.replace(/[\r\n\t]+/g, '') : "";

			let value = "\tlet " + varRegistry.inc('chart') + " = Api.CreateChart(\n"
			+ "\t\t\"" + chartType + "\",\n"
			+ "\t\t" + JSON.stringify(seriesData) + ",\n"
			+ "\t\t" + JSON.stringify(seriesNames) + ",\n"
			+ "\t\t" + JSON.stringify(categories) + ",\n"
			+ "\t\t" + width + ",\n"
			+ "\t\t" + height + ",\n"
			+ "\t\t" + style + "\n"
			+ "\t);\n"
			+ "\t" + varRegistry.get('chart') + ".SetTitle(\"" + title + "\", " + 14 + ");\n"
			+ "\t" + varRegistry.get('chart') + ".SetPosition("+ chart.x + " * 36000, " + chart.y +" * 36000);\n"
			+ "\tpresentation.GetCurrentSlide().AddObject(" + varRegistry.get('chart') + ");\n"
			return value;
		},
		addComment				: function(comment){
			return "\tpresentation.GetCurrentSlide().AddComment("
				+ comment.x + " * 36000, "
				+ comment.y + " * 36000, "
				+ "\"" + comment.Data.m_sText + "\", "
				+ "\"" + comment.Data.m_sUserName + "\", "
				+ "\"" + comment.Data.m_sUserId + "\""
			+ ")\n";

			// api set time
		},
		addHyperlink			: function(hp){
			// no api
			return ""
		},
		addParagraph			: function()
		{
			return "\tpresentation.InsertParagraphBreak();\n"
		},
		addShape				: function(shapeProps){
			let fill = shapeProps.fill.getRGBAColor();
			let border = shapeProps.border;
			let borderwidth = border.w / 36000;
			let borderColor = border.Fill.getRGBAColor();
			let varName = varRegistry.register(shapeProps.id, 'shape');

			return "\tlet " + varName + " = (function () {\n" +
					"\t\tlet fill = Api.CreateSolidFill(Api.CreateRGBColor(" + fill.R +", " + fill.G + ", " + fill.B + "));\n" +
					"\t\tlet stroke = Api.CreateStroke(" + borderwidth +" * 36000, Api.CreateSolidFill(Api.CreateRGBColor(" + borderColor.R +", " + borderColor.G + ", " + borderColor.B + ")));\n" +
					"\t\tlet shape = Api.CreateShape(\"" + shapeProps.type + "\", " + shapeProps.extX + " * 36000, " + shapeProps.extY + " * 36000, fill, stroke);\n" +
					"\t\tshape.SetPosition(" + shapeProps.pos.x + " * 36000 , " + shapeProps.pos.y + " * 36000 );\n" +
					"\t\tpresentation.GetCurrentSlide().AddObject(shape);\n" +
					"\t\tshape.Select(true);\n" +
					"\t\treturn shape\n" +
				"\t}());\n";
		},
		selectDrawing				: function(shapes)
		{
			let str = "";
			let shapesSort = [];
			function isDrawingOnSlide(oDrawing) {
				if(!oDrawing) return false;
				let oSlide = oDrawing.parent;
				if(!oSlide) return false;
				let aSpTree = oSlide instanceof window['AscCommonSlide'].Slide && oSlide.cSld && oSlide.cSld.spTree;
				if(!Array.isArray(aSpTree)) return false;
				for(let nSp = 0; nSp < aSpTree.length; ++nSp) {
					if(aSpTree[nSp] === oDrawing) {
						return true;
					}
				}
				return false;
			}

			shapes.forEach(function(element) {
				if (!shapesSort.includes(element) && isDrawingOnSlide(element))
					shapesSort.push(element);
			});

			if (shapesSort.length === 0)
				return "";

			shapesSort = shapesSort.filter(function(element) {
				return varRegistry.get(element.getObjectName()) !== undefined;
			});

			if (shapesSort.length === 0)
				return "";

			let ids = "";
			for (let i = 0; i < shapesSort.length; i++) {
				let varName = varRegistry.get(shapesSort[i].getObjectName());
				if (varName !== undefined)
					ids += varName + (i != (shapesSort.length - 1) && shapesSort.length > 1 ? ", " : "")
			}

			return str + "\t[" + ids + "]\n\t\t.forEach(function(drawing, index){drawing.Select(index === 0)});\n"

		},
		deselectDrawing			: function(oAdditional)
		{
			let varName = varRegistry.get(oAdditional.name);
			let isTable = oAdditional.isTable;

			if (isTable === true) {
				return "\tpresentation.GetCurrentTable().Unselect();\n";
			}

			if (varName === undefined) {
				return "\tApi.GetSelection().GetShapes().forEach(function(drawing){drawing.Unselect()});\n";
			}
			return "\t[" + varName + "]\n\t\t.forEach(function(drawing){drawing.Unselect()});\n"
		},
		setDrawingFill			: function(unifill)
		{
			//TODO add transparent when add api
			if (unifill.fill.type === Asc.c_oAscFill.FILL_TYPE_SOLID)
			{
				let color = unifill.fill.color.color.RGBA;
				//let transparent = unifill.transparent;

				return "\tApi.GetSelection().GetShapes()\n"
						+ "\t\t.forEach(shape => shape.Fill(Api.CreateSolidFill(Api.CreateRGBColor(" + color.R + ", " + color.G + ", " + color.B + "))));\n";
			}
			else if (unifill.fill instanceof AscFormat.CGradFill)
			{
				let colors = unifill.fill.colors;
				let strColor = "";

				for(let nColor = 0; nColor < colors.length; nColor++)
				{
					let gs = colors[nColor];
					let color = gs.color.color.RGBA;

					strColor += (nColor !== 0)
						? ",\n\t\t\t\t"
						: "\n\t\t\t\t";

					strColor += "Api.CreateGradientStop(Api.CreateRGBColor(" + color.R + ", " + color.G + ", " + color.B + "), " + gs.pos + ")";
				}

				let gradient = "";
				if (unifill.fill.lin) // leinear
					gradient = "\n\t\t\tApi.CreateLinearGradientFill([" + strColor + "\n\t\t\t], " + unifill.fill.lin.angle + ")";
				else if (unifill.fill.path) // radial
					gradient = "\n\t\t\tApi.CreateRadialGradientFill([" + strColor + "\n\t\t\t])";

				return "\tApi.GetSelection().GetShapes()\n"
						+ "\t\t.forEach(shape => shape.Fill(" + gradient + "));\n";

			}
			else if (unifill.fill instanceof AscFormat.CBlipFill)
			{
				let base64data = unifill.fill.getBase64RasterImageId()
				let blipFill = "\n\t\t\tApi.CreateBlipFill(\"" + base64data + "\", \"" + (unifill.fill.tile ? "tile" : "stretch") + "\")";
				return "\tApi.GetSelection().GetShapes()\n"
						+ "\t\t.forEach(shape => shape.Fill(" + blipFill + ")\n\t\t);\n";

			}
			else if (unifill.fill instanceof AscFormat.CPattFill)
			{
				let bgClr = unifill.fill.bgClr.color.RGBA;
				let fgClr = unifill.fill.fgClr.color.RGBA;
				let ftype = unifill.fill.ftype;
				let type = AscCommon.global_hatch_names[ftype];

				let blipFill = "\n\t\t\tApi.CreatePatternFill(\n\t\t\t\t\"" + type + "\",\n\t\t\t\t" + "Api.CreateRGBColor(" + bgClr.R + ", " + bgClr.G + ", " + bgClr.B + "),\n\t\t\t\tApi.CreateRGBColor(" + fgClr.R + ", " + fgClr.G + ", " + fgClr.B + ")";
				return "\tApi.GetSelection().GetShapes()\n"
						+ "\t\t.forEach(shape => shape.Fill(" + blipFill + ")\n\t\t\t)\n\t\t);\n";
			}
			else if (unifill.fill instanceof AscFormat.CNoFill)
			{
				return "\tApi.GetSelection().GetShapes()\n"
						+ "\t\t.forEach(shape => shape.Fill(Api.CreateNoFill()));\n";
			}
		},
		setDrawingLine			: function(line)
		{
			let strStrokeShape = "\t(function () {\n";
			let type = AscFormat.CLn.prototype.GetDashByCode(line.prstDash);
			let color = line.Fill.fill.color.color.RGBA;

			strStrokeShape += "\t\tlet stroke = Api.CreateStroke(\n" +
					"\t\t\t" + line.w / 12700.0 + " * 12700.0,\n" +
					"\t\t\t" + "Api.CreateSolidFill(Api.CreateRGBColor(" + color.R + ", " + color.G + ", " + color.B + ")),\n" +
					"\t\t\t\"" + type + "\"\n" +
				"\t\t);\n";

			strStrokeShape += "\t\tApi.GetSelection().GetShapes()\n"
				+ "\t\t\t.forEach(shape => shape.SetOutLine(stroke));\n";

			strStrokeShape += "\t}());\n";

			return strStrokeShape;
		},
		setDrawingRotation		: function(nRot)
		{
			return "\tApi.GetSelection().GetShapes().forEach(draw => {\n" +
				"\t\tdraw.SetRotation(" + nRot * 180 / Math.PI + ");\n" +
				"\t});\n"
		},
		setDrawingFlipH			: function(isFlip)
		{
			return "\tApi.GetSelection().GetShapes()\n"
				+ "\t\t.filter(item => item.GetClassType() === \"shape\")\n"
				+ "\t\t.forEach(shape => {shape.SetHorFlip(" + isFlip +")});\n";

		},
		setDrawingFlipV			: function(isFlip)
		{
			return "\tApi.GetSelection().GetShapes()\n"
				+ "\t\t.filter(item => item.GetClassType() === \"shape\")\n"
				+ "\t\t.forEach(shape => {shape.SetVertFlip(" + isFlip + ")});\n";

		},
		setDrawingWrapping		: function(props)
		{
			return "\tApi.GetSelection().GetShapes().forEach(draw => draw.SetWrappingStyle(\"" + getWrappingStyleName(props) + "\"))\n"
		},
		setDrawingPos			: function(pos)
		{
			return "\tApi.GetSelection().GetShapes().forEach(draw => {\n" +
				"\t\tdraw.SetPosition(" + pos.x + " * 36000.0, " + pos.y + " * 36000.0);\n" +
				"\t});\n"
		},
		setShapeX				: function(pos)
		{
			return "\tApi.GetSelection().GetShapes().forEach(draw => {\n" +
				"\t\tdraw.SetPosX(" + pos + " * 36000);\n" +
				"\t});\n"
		},
		setShapeY				: function(pos)
		{
			return "\tApi.GetSelection().GetShapes().forEach(draw => {\n" +
				"\t\tdraw.SetPosY(" + pos + " * 36000);\n" +
				"\t});\n"
		},
		setPositionH			: function(data)
		{
			let relative = "";
			switch(data.relativeFrom)
			{
				 case Asc.c_oAscRelativeFromH.Character:	relative = "character";		break;
				 case Asc.c_oAscRelativeFromH.Column:		relative = "column";		break;
				 case Asc.c_oAscRelativeFromH.LeftMargin:	relative = "leftMargin";	break;
				 case Asc.c_oAscRelativeFromH.Margin:		relative = "margin";		break;
				 case Asc.c_oAscRelativeFromH.RightMargin:	relative = "rightMargin";	break;
				 case Asc.c_oAscRelativeFromH.Page:			relative = "page";			break;
				 default:									relative = "page";
			}

			if (data.useAlign)
			{
				let useAlign = "";
				switch (data.value) {
					case Asc.c_oAscAlignH.Center:	useAlign = "center";	break;
					case Asc.c_oAscAlignH.Left:		useAlign = "left";		break;
					case Asc.c_oAscAlignH.Right:	useAlign = "right";		break;
					default:						useAlign = "center";
				}

				return "\tdoc.GetSelectedDrawings().forEach(draw => {" +
					"draw.SetHorAlign(\"" + relative + "\", \"" + useAlign + "\")" +
				"});\n"
			}

			return "\tdoc.GetSelectedDrawings().forEach(draw => {" +
				"draw.SetHorPosition(\"" + relative + "\", " + data.value + " * 36000.0, " + data.percent + ")" +
			"});\n"
		},
		setPositionV			: function(data)
		{
			let relative = "";
			switch(data.relativeFrom)
			{
				 case Asc.c_oAscRelativeFromH.Character:	relative = "character";		break;
				 case Asc.c_oAscRelativeFromH.Column:		relative = "column";		break;
				 case Asc.c_oAscRelativeFromH.LeftMargin:	relative = "leftMargin";	break;
				 case Asc.c_oAscRelativeFromH.Margin:		relative = "margin";		break;
				 case Asc.c_oAscRelativeFromH.RightMargin:	relative = "rightMargin";	break;
				 case Asc.c_oAscRelativeFromH.Page:			relative = "page";			break;
				 default:									relative = "page";
			}

			if (data.useAlign)
			{
				let useAlign = "";
				switch (data.value) {
					case Asc.c_oAscAlignV.Bottom:	useAlign = "bottom";	break;
					case Asc.c_oAscAlignV.Center:	useAlign = "center";	break;
					case Asc.c_oAscAlignV.Top:		useAlign = "top";		break;
					default:						useAlign = "center";
				}

				return "\tdoc.GetSelectedDrawings().forEach(draw => {" +
					"draw.SetVerAlign(\"" + relative + "\", \"" + useAlign + "\")" +
				"});\n"
			}

			return "\tdoc.GetSelectedDrawings().forEach(draw => {" +
				"draw.SetVerPosition(\"" + relative + "\", " + data.value + " * 36000.0, " + data.percent + ")" +
			"});\n"
		},
		setShapeSize			: function(oSize)
		{
			return "\tApi.GetSelection().GetShapes().forEach(draw => {\n" +
				"\t\tdraw.SetSize(" + oSize.width + " * 36000.0, " + oSize.height + " * 36000.0);\n" +
			"\t});\n"
		},
		setDrawingDistances		: function(oDistances)
		{
			return "\tdoc.GetSelectedDrawings().forEach(draw => {\n"
					+ "\t\t\tdraw.SetDistances(" + oDistances.Left + " * 36000.0, " + oDistances.Top + " * 36000.0, " + oDistances.Right + " * 36000.0, " + oDistances.Bottom + " * 36000.0)\n"
				+ "\t});\n"

		},
		setShapeInnerPadding	: function(oPadding)
		{
			let left =	(oPadding.set.Left !== null)	? oPadding.set.Left		: oPadding.original.Left;
			let right =	(oPadding.set.Right !== null)	? oPadding.set.Right	: oPadding.original.Right;
			let top =	(oPadding.set.Top !== null)		? oPadding.set.Top		: oPadding.original.Top;
			let bottom =(oPadding.set.Bottom !== null)	? oPadding.set.Bottom	: oPadding.original.Bottom;
			return "\tApi.GetSelection().GetShapes()\n"
				+ "\t\t.forEach(draw => {\n"
					+ "\t\t\tdraw.SetPaddings(" + left + " * 36000.0, " + top + " * 36000.0, " + right + " * 36000.0, " + bottom + " * 36000.0)\n"
				+ "\t\t});\n"

		},
		setShapeRelSizeH		: function(oSize)
		{
			let relative = "";
			switch(oSize.RelativeFrom)
			{
				 case Asc.c_oAscRelativeFromH.Character:	relative = "character";		break;
				 case Asc.c_oAscRelativeFromH.Column:		relative = "column";		break;
				 case Asc.c_oAscRelativeFromH.LeftMargin:	relative = "leftMargin";	break;
				 case Asc.c_oAscRelativeFromH.Margin:		relative = "margin";		break;
				 case Asc.c_oAscRelativeFromH.RightMargin:	relative = "rightMargin";	break;
				 case Asc.c_oAscRelativeFromH.Page:			relative = "page";			break;
				 default:									relative = "page";
			}
			return "\tdoc.GetSelectedDrawings().forEach(draw => {" +
				"draw.SetRelativeWidth(\"" + relative + "\", " + oSize.Value + ")" +
			"});\n"
		},
		setShapeRelSizeV		: function(oSize)
		{
			let relative = "";
			switch(oSize.RelativeFrom)
			{
				 case Asc.c_oAscRelativeFromV.BottomMargin:	relative = "bottomMargin";		break;
				 case Asc.c_oAscRelativeFromV.Line:			relative = "line";				break;
				 case Asc.c_oAscRelativeFromV.TopMargin:	relative = "topMargin";			break;
				 case Asc.c_oAscRelativeFromV.Margin:		relative = "margin";			break;
				 case Asc.c_oAscRelativeFromV.Paragraph:	relative = "paragraph";			break;
				 case Asc.c_oAscRelativeFromV.Page:			relative = "page";				break;
				 default:									relative = "page";
			}
			return "\tdoc.GetSelectedDrawings().forEach(draw => {" +
				"draw.SetRelativeHeight(\"" + relative + "\", " + oSize.Value + ")" +
			"});\n"
		},
		setGeometry			: function(type)
		{
			return "\tApi.GetSelection().GetShapes().forEach(draw => {\n" +
				"\t\tdraw.SetGeometry(Api.CreatePresetGeometry(\""+ type +"\"));\n" +
				"\t});\n"
		},
		setDrawingTitle		: function(title){
			return "\tApi.GetSelection().GetShapes()\n"
				+ "\t\t.forEach(function(shape) {\n"
				+ "\t\t\tshape.SetTitle(\"" + title + "\");\n"
				+ "\t\t});\n";
		},
		setDrawingDescription	: function(description){
			return "\tApi.GetSelection().GetShapes()\n"
				+ "\t\t.forEach(function(shape) {\n"
				+ "\t\t\tshape.SetDescription(\"" + description + "\");\n"
				+ "\t\t});\n";
		},
		setDrawingAspectRatio	: function(aspect){
			return "\tApi.GetSelection().GetShapes().forEach(function(draw) { draw.SetLockAspect(" + aspect + ") });\n";
		},
		setTableStyle		: function(oAdditional)
		{
			let reg = private_getTableVar(oAdditional.tableId);
			let style = oMacroRecorderContext.editor.getLogicDocument().globalTableStyles.GetName(oAdditional.data);
			if (!style) return "";
			let styleVar = varRegistry.inc('tableStyle');
			return reg.declaration
				+ "\tlet " + styleVar + " = presentation.GetStyle(\"" + style + "\");\n"
				+ "\tif (" + styleVar + ") " + reg.varName + ".SetStyle(" + styleVar + ");\n";
		},
		setTableResize		: function(oAdditional)
		{
			let reg = private_getTableVar(oAdditional.tableId);
			let width = AscCommon.MMToTwips(oAdditional.width, -1);
			let height = AscCommon.MMToTwips(oAdditional.height, -1);

			return reg.declaration
				+ "\t" + reg.varName + ".SetWidth(\"twips\", " + width + ");\n"
				//+ reg.varName + ".SetHeight(\"twips\", " + height + ");\n"   // no api for height
		},
		setTableBorders		: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let borders = oAdditional.data;
			let bottom = borders.Bottom;
			let top = borders.Top;
			let left = borders.Left;
			let right = borders.Right;
			let insideH = borders.InsideH;
			let insideV = borders.InsideV;
			let insideHStr = insideH
				? "\t" + reg.varName + ".SetTableBorderInsideH(\"" + (insideH.Value ? 'single' : 'none') + "\", " + insideH.Size * 72 * 8 / 25.4 + ", " + insideH.Space + ", " + insideH.Color.r + ", " + insideH.Color.g + ", " + insideH.Color.b + ");\n"
				: "";
			let insideVStr = insideV
				? "\t" + reg.varName + ".SetTableBorderInsideV(\"" + (insideV.Value ? 'single' : 'none') + "\", " + insideV.Size * 72 * 8 / 25.4 + ", " + insideV.Space + ", " + insideV.Color.r + ", " + insideV.Color.g + ", " + insideV.Color.b + ");\n"
				: "";
			return reg.declaration +
				"\t" + reg.varName + ".SetTableBorderBottom(\"" + (bottom.Value ? 'single' : 'none') + "\", " + bottom.Size * 72 * 8 / 25.4 + ", " + bottom.Space + ", " + bottom.Color.r + ", " + bottom.Color.g + ", " + bottom.Color.b + ");\n" +
				"\t" + reg.varName + ".SetTableBorderLeft(\"" + (left.Value ? 'single' : 'none') + "\", " + left.Size * 72 * 8 / 25.4 + ", " + left.Space + ", " + left.Color.r + ", " + left.Color.g + ", " + left.Color.b + ");\n" +
				"\t" + reg.varName + ".SetTableBorderRight(\"" + (right.Value ? 'single' : 'none') + "\", " + right.Size * 72 * 8 / 25.4 + ", " + right.Space + ", " + right.Color.r + ", " + right.Color.g + ", " + right.Color.b + ");\n" +
				"\t" + reg.varName + ".SetTableBorderTop(\"" + (top.Value ? 'single' : 'none') + "\", " + top.Size * 72 * 8 / 25.4 + ", " + top.Space + ", " + top.Color.r + ", " + top.Color.g + ", " + top.Color.b + ");\n" +
				insideHStr +
				insideVStr;
		},
		setCellBorders		: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let borders = oAdditional.data;
			let bottom = borders.Bottom;
			let top = borders.Top;
			let left = borders.Left;
			let right = borders.Right;

			return reg.declaration
				+ "\t" + reg.varName + ".GetSelectedCells().forEach(function(cell){\n"
				+ (bottom ? "\t\tcell.SetCellBorderBottom(\"" + (bottom.Value ? 'single' : 'none') + "\", " + bottom.Size + ", Api.CreateSolidFill(Api.RGB(" + bottom.Color.r + ", " + bottom.Color.g + ", " + bottom.Color.b + "))); \n" : "")
				+ (left ? "\t\tcell.SetCellBorderLeft(\"" + (left.Value ? 'single' : 'none') + "\", " + left.Size + ", Api.CreateSolidFill(Api.RGB(" + left.Color.r + ", " + left.Color.g + ", " + left.Color.b + "))); \n" : "")
				+ (right ? "\t\tcell.SetCellBorderRight(\"" + (right.Value ? 'single' : 'none') + "\", " + right.Size + ", Api.CreateSolidFill(Api.RGB(" + right.Color.r + ", " + right.Color.g + ", " + right.Color.b + "))); \n" : "")
				+ (top ? "\t\tcell.SetCellBorderTop(\"" + (top.Value ? 'single' : 'none') + "\", " + top.Size + ", Api.CreateSolidFill(Api.RGB(" + top.Color.r + ", " + top.Color.g + ", " + top.Color.b + "))); \n" : "")
				+ "\t});\n";
		},
		setTableDefaultMargin	: function(oAdditional){
			let d = oAdditional.data;
			if (!d || !d.Bottom || !d.Left || !d.Right || !d.Top)
				return "";
			let reg = private_getTableVar(oAdditional.tableId);
			let bottom = d.Bottom;
			let top = d.Top;
			let left = d.Left;
			let right = d.Right;
			return reg.declaration +
				(bottom ? "\t" + reg.varName + ".SetTableCellMarginBottom(" + AscCommon.MMToTwips(bottom, -1) + ");\n" : "") +
				(left ? "\t" + reg.varName + ".SetTableCellMarginLeft(" + AscCommon.MMToTwips(left, -1) + ");\n" : "") +
				(right ? "\t" + reg.varName + ".SetTableCellMarginRight(" + AscCommon.MMToTwips(right, -1) + ");\n" : "") +
				(top ? "\t" + reg.varName + ".SetTableCellMarginTop(" + AscCommon.MMToTwips(top, -1) + ");\n" : "");
		},
		setCellMargins		: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let d = oAdditional.data;
			let bottom = d.Bottom;
			let top = d.Top;
			let left = d.Left;
			let right = d.Right;
			return reg.declaration + "\t" + reg.varName + ".GetSelectedCells().forEach(function(cell){\n" +
				(bottom ? "\t\tcell.SetCellMarginBottom(" + AscCommon.MMToTwips(bottom, -1) + ");\n" : "") +
				(left ? "\t\tcell.SetCellMarginLeft(" + AscCommon.MMToTwips(left, -1) + ");\n" : "") +
				(right ? "\t\tcell.SetCellMarginRight(" + AscCommon.MMToTwips(right, -1) + ");\n" : "") +
				(top ? "\t\tcell.SetCellMarginTop(" + AscCommon.MMToTwips(top, -1) + ");\n" : "") +
				"\t});\n";
		},
		setTableBackground	: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let color = oAdditional.data.Color;
			return reg.declaration + "\t" + reg.varName + ".SetBackgroundColor(Api.RGB(" + color.r + ", " + color.g + ", " + color.b + "));\n";
		},
		setCellBackground	: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let color = oAdditional.data.Color;
			return reg.declaration + "\t" + reg.varName + ".GetSelectedCells().forEach(function(cell){\n" +
				"\t\tcell.SetBackgroundColor(Api.RGB(" + color.r + ", " + color.g + ", " + color.b + "));\n" +
				"\t});\n";
		},
		setCellWidth		: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let cellPrVar = varRegistry.inc('cellPr');
			return reg.declaration
				+ "\tlet " + cellPrVar + " = Api.CreateTableCellPr();\n"
				+ "\t" + cellPrVar + ".SetWidth(\"twips\", " + AscCommon.MMToTwips(oAdditional.data, -1) + ");\n"
				+ "\t" + reg.varName + ".GetSelectedColumnsCells().forEach(function(cell){\n"
				+ "\t\tcell.SetCellPr(" + cellPrVar + ");\n"
				+ "\t});\n";
		},
		setRowHeightTable	: function(oAdditional){
			let reg = private_getTableVar(oAdditional.tableId);
			let rowPrVar = varRegistry.inc('rowPr');
			return reg.declaration
				+ "\tlet " + rowPrVar + " = Api.CreateTableRowPr();\n"
				+ "\t" + rowPrVar + ".SetHeight(\"atLeast\", " + AscCommon.MMToTwips(oAdditional.data, -1) + ");\n"
				+ "\t" + reg.varName + ".GetSelectedRows().forEach(function(row){\n"
				+ "\t\trow.SetRowPr(" + rowPrVar + ");\n"
				+ "\t});\n";
		},
		moveCursorLeft			: function(arrData){
			let data = groupDataForCursor(arrData);
			let text = "";

			for (let i = 0; i < data.length; i++)
			{
				let currentChange = data[i];
				let nCount = currentChange.count;
				let pattern = currentChange.pattern;

				if (pattern.isRtl)
					text += "\tpresentation.MoveCursorRight(" + nCount + ", " + pattern.isAddSelect + ", " + pattern.isWord + ");\n";
				else
					text += "\tpresentation.MoveCursorLeft(" + nCount + ", " + pattern.isAddSelect + ", " + pattern.isWord + ");\n";
			}

			return text;
		},
		moveCursorRight			: function(arrData){
			let data = groupDataForCursor(arrData);
			let text = "";

			for (let i = 0; i < data.length; i++)
			{
				let currentChange	= data[i];
				let nCount			= currentChange.count;
				let pattern			= currentChange.pattern;

				if (pattern.isRtl)
					text += "\tpresentation.MoveCursorLeft(" + nCount + ", " + pattern.isAddSelect + ", " + pattern.isWord + ");\n";
				else
					text += "\tpresentation.MoveCursorRight(" + nCount + ", " + pattern.isAddSelect + ", " + pattern.isWord + ");\n";
			}

			return text;
		},
		moveCursorUp			: function(arrData){
			let data = groupDataForCursor(arrData, true);
			let text = "";

			for (let i = 0; i < data.length; i++)
			{
				let currentChange	= data[i];
				let nCount			= currentChange.count;
				let pattern			= currentChange.pattern;

				text += "\tpresentation.MoveCursorUp(" + nCount + ", " + pattern.isAddSelect + ", " + pattern.isWord + ");\n";
			}
			return text;
		},
		moveCursorDown			: function(arrData){
			let data = groupDataForCursor(arrData, true);
			let text = "";

			for (let i = 0; i < data.length; i++)
			{
				let currentChange	= data[i];
				let nCount			= currentChange.count;
				let pattern			= currentChange.pattern;

				text += "\tpresentation.MoveCursorDown(" + nCount + ", " + pattern.isAddSelect + ", " + pattern.isWord + ");\n";
			}
			return text;
		}
	};

	// alignTo no api
	// merge shapes no api
	// show from start/n-slide ... when add api

	const PEActionMap = {};
	PEActionMap[AscDFH.historydescription_Presentation_ParagraphAdd] 				= presActions.paragraphAdd;
	PEActionMap[AscDFH.historydescription_Presentation_PutTextPrBold]				= presActions.putTextPrBold;
	PEActionMap[AscDFH.historydescription_Document_SetTextBoldHotKey]				= presActions.putTextPrBold;
	PEActionMap[AscDFH.historydescription_Presentation_PutTextPrItalic]				= presActions.putTextPrItalic;
	PEActionMap[AscDFH.historydescription_Document_SetTextItalicHotKey]				= presActions.putTextPrItalic;
	PEActionMap[AscDFH.historydescription_Presentation_PutTextPrUnderline]			= presActions.putTextPrUnderline;
	PEActionMap[AscDFH.historydescription_Document_SetTextUnderlineHotKey]			= presActions.putTextPrUnderline;
	PEActionMap[AscDFH.historydescription_Presentation_PutTextPrStrikeout]			= presActions.putTextPrStrikeout;
	PEActionMap[AscDFH.historydescription_Document_SetTextStrikeoutHotKey]			= presActions.putTextPrStrikeout;
	PEActionMap[AscDFH.historydescription_Document_SetTextVertAlign]				= presActions.setTextVertAlign;
	PEActionMap[AscDFH.historydescription_Document_SetTextVertAlignHotKey3]			= presActions.setTextVertAlign;
	PEActionMap[AscDFH.historydescription_Document_SetTextVertAlignHotKey2]			= presActions.setTextVertAlign;
	PEActionMap[AscDFH.historydescription_Document_SetTextHighlight]				= presActions.setTextHighlight;
	PEActionMap[AscDFH.historydescription_Presentation_PutTextColor]				= presActions.putTextColor;
	PEActionMap[AscDFH.historydescription_Presentation_ParagraphClearFormatting]	= presActions.clearFormatting;
	PEActionMap[AscDFH.historydescription_Presentation_PutTextPrFontName]			= presActions.putTextPrFontName;
	PEActionMap[AscDFH.historydescription_Presentation_PutTextPrFontSize]			= presActions.putTextPrFontSize;
	PEActionMap[AscDFH.historydescription_Presentation_SetParagraphAlign]			= presActions.setParagraphAlign;
	PEActionMap[AscDFH.historydescription_Document_SetParagraphAlignHotKey]			= presActions.setParagraphAlign;
	PEActionMap[AscDFH.historydescription_Presentation_AddNextSlide]				= presActions.addNextSlide;
	PEActionMap[AscDFH.historydescription_Presentation_DeleteSlides]				= presActions.deleteSlides;
	PEActionMap[AscDFH.historydescription_Presentation_ChangeLayout]				= presActions.changeLayout;
	PEActionMap[AscDFH.historydescription_Presentation_SetVerticalAlign]			= presActions.setVerticalAlign;
	// PEActionMap[AscDFH.historydescription_Presentation_BringForward]				= presActions.bringForward;
	// PEActionMap[AscDFH.historydescription_Presentation_BringToFront]				= presActions.bringToFront;
	// PEActionMap[AscDFH.historydescription_Presentation_BringBackward]			= presActions.bringBackward;
	// PEActionMap[AscDFH.historydescription_Presentation_SendToBack]				= presActions.sendToBack;
	PEActionMap[AscDFH.historydescription_Presentation_CreateGroup]					= presActions.createGroup;
	PEActionMap[AscDFH.historydescription_Presentation_UnGroup]						= presActions.unGroup;
	//PEActionMap[AscDFH.historydescription_Presentation_PutTextPrIncreaseFontSize]	= presActions.putTextPrIncreaseFontSize;
	//PEActionMap[AscDFH.historydescription_Presentation_ParagraphIncDecFontSize]	= presActions.incDecFontSize;
	//PEActionMap[AscDFH.historydescription_Presentation_SetParagraphNumbering]		= presActions.setNumbering;
	PEActionMap[AscDFH.historydescription_Presentation_PutTextPrLineSpacing]		= presActions.putTextPrLineSpacing;
	//PEActionMap[AscDFH.historydescription_Spreadsheet_Remove]						= presActions.paragraphRemove; // stange
	PEActionMap[AscDFH.historydescription_Presentation_AddFlowTable]				= presActions.addFlowTable;
	PEActionMap[AscDFH.historydescription_Presentation_AddFlowImage]				= presActions.addFlowImage;
	PEActionMap[AscDFH.historydescription_Presentation_AddShape]					= presActions.addShape;
	PEActionMap[AscDFH.historydescription_Presentation_AddChart]					= presActions.addChart;
	PEActionMap[AscDFH.historydescription_Presentation_AddComment]					= presActions.addComment;
	PEActionMap[AscDFH.historydescription_Presentation_HyperlinkAdd]				= presActions.addHyperlink;
	PEActionMap[AscDFH.historydescription_Presentation_AddNewParagraph]				= presActions.addParagraph;
	PEActionMap["SetDrawingPos"]													= presActions.setDrawingPos;
	PEActionMap["SetShapeSize"]														= presActions.setShapeSize;
	PEActionMap["SetDrawingRotation"]												= presActions.setDrawingRotation;
	PEActionMap["SetDrawingFill"]													= presActions.setDrawingFill;
	PEActionMap["SetGeometry"]														= presActions.setGeometry;
	PEActionMap["SetDrawingLine"]													= presActions.setDrawingLine;
	PEActionMap["SetShapeX"]														= presActions.setShapeX;
	PEActionMap["SetShapeY"]														= presActions.setShapeY;
	PEActionMap['SetShapeInnerPadding']												= presActions.setShapeInnerPadding;
	PEActionMap['SetDrawingFlipH']													= presActions.setDrawingFlipH;
	PEActionMap['SetDrawingFlipV']													= presActions.setDrawingFlipV;
	PEActionMap["SelectDrawing"]													= presActions.selectDrawing;
	PEActionMap["DeselectDrawing"]													= presActions.deselectDrawing;
	PEActionMap["SetDrawingTitle"]													= presActions.setDrawingTitle;
	PEActionMap["SetDrawingDescription"]											= presActions.setDrawingDescription;
	PEActionMap["SetDrawingAspectRatio"]											= presActions.setDrawingAspectRatio;
	PEActionMap['remove']															= presActions.remove;
	// Table
	PEActionMap["SetTableBorders"]													= presActions.setTableBorders;
	PEActionMap["SetTableResize"]													= presActions.setTableResize;
	PEActionMap["SetTableStyle"]													= presActions.setTableStyle;
	PEActionMap["SetCellBorders"]													= presActions.setCellBorders;
	PEActionMap["SetTableBackground"]												= presActions.setTableBackground;
	PEActionMap["SetCellBackground"]												= presActions.setCellBackground;
	PEActionMap["SetTableDefaultMargin"]											= presActions.setTableDefaultMargin;
	PEActionMap["SetCellMargins"]													= presActions.setCellMargins;
	PEActionMap["SetRowHeightTable"]												= presActions.setRowHeightTable;
	PEActionMap["SetCellWidth"]														= presActions.setCellWidth;
	PEActionMap['moveCursorLeft']													= presActions.moveCursorLeft;
	PEActionMap['moveCursorRight']													= presActions.moveCursorRight;
	PEActionMap['moveCursorUp']														= presActions.moveCursorUp;
	PEActionMap['moveCursorDown']													= presActions.moveCursorDown;
	PEActionMap[AscDFH.historydescription_Document_AddLetter]						= presActions.addLetter;

	//--------------------------------------------------------export----------------------------------------------------
	AscCommon.MacroRecorder = MacroRecorder;
	
	MacroRecorder.prototype["start"]        = MacroRecorder.prototype.start;
	MacroRecorder.prototype["stop"]         = MacroRecorder.prototype.stop;
	MacroRecorder.prototype["cancel"]       = MacroRecorder.prototype.cancel;
	MacroRecorder.prototype["pause"]        = MacroRecorder.prototype.pause;
	MacroRecorder.prototype["resume"]       = MacroRecorder.prototype.resume;
	MacroRecorder.prototype["isInProgress"] = MacroRecorder.prototype.isInProgress;
	MacroRecorder.prototype["isPaused"]     = MacroRecorder.prototype.isPaused;
	
})(window);
