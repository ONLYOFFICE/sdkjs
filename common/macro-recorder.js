/*
 * (c) Copyright Ascensio System SIA 2010-2025
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

(function (window)
{
	let oMacroRecorderContext = {};

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
				if (_t.editor.editorId !== AscCommon.c_oEditorId.Word)
					return;

				_t.addStepData("remove", 1);
				// TODO when we have:
				// * Selection.Delete
				// * Selection.TypeBackspace
			}
			else if (e.KeyCode === 9) // Tab
			{
				if (_t.editor.editorId !== AscCommon.c_oEditorId.Word)
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
				if (_t.editor.editorId !== AscCommon.c_oEditorId.Word)
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
				if (_t.editor.editorId !== AscCommon.c_oEditorId.Word)
					return;

				_t.addStepData('moveCursorUp', [{
					isAddSelect:	e.IsShift(),
					isWord:			e.IsCtrl()
				}]);
			}
			else if (e.KeyCode === 39) // Right Arrow
			{
				if (_t.editor.editorId !== AscCommon.c_oEditorId.Word)
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
				if (_t.editor.editorId !== AscCommon.c_oEditorId.Word)
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
			CounterStore.reset();
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
			actionsMacros = WordActionsMacroList;
		else if (this.editor.editorId === AscCommon.c_oEditorId.Spreadsheet)
			actionsMacros = CellActionsMacroList;
		else if (this.editor.editorId === AscCommon.c_oEditorId.Presentation)
			actionsMacros = PresentationActionMacroList;

		return actionsMacros;
	}
	MacroRecorder.prototype.onFinalizeAction = function()
	{
		// todo implement onFinalizeAction
		this.getResultByType(this.prevChangeType, this.prevData);
		this.prevChangeType	= null;
		this.prevData		= undefined;
	};
	MacroRecorder.prototype.onAction = function(type, additional, isStart)
	{
		if (isStart === true)
			this.currentDescription = type;

		if (!this.isInProgress() || this.isPaused() || undefined === additional)
			return;

		if (isStart === false)
			type = this.currentDescription;

		//additional = this.proceedDataBefoeApply(additional);

		if (this.prevChangeType === type)
		{
			this.prevData = this.joinDataForMacros(this.prevData, additional);
		}
		else if (type !== AscDFH.historydescription_Document_AutoCorrectMath)
		{
			this.getResultByType(this.prevChangeType, this.prevData);
			this.prevChangeType = type;
			this.prevData = additional;
		}

		if (this.isFirstAction)
		{
			this.addDefualtVaribalesForEditor();
			this.isFirstAction = false;
		}
	};
	// for now duplicate, leter we delete onAction
	MacroRecorder.prototype.addStepData = function(type, additional)
	{
		if (!this.isInProgress() || this.isPaused() || undefined === additional)
			return;

		// for meta action
		if (type === 'remove')
		{
			this.getResultByType(type, additional);
			return;
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
			+ "";
	};
	MacroRecorder.prototype.proceedDefualtVariablesForPresentation = function()
	{
		this.result += "\tlet presentation = Api.GetPresentation();\n"
			+ "";
	};

	function CounterStoreFn()
	{
		this.store = {};
		this.get = function(name)
		{
			this.create(name);
			return name + (this.store[name] === 0 ? "" : this.store[name]);
		};
		this.inc = function(name) {
			if (this.create(name))
				return this.get(name);

			this.store[name] = this.store[name] + 1;
			return this.get(name);
		};
		this.create = function(name)
		{
			if (!this.store.hasOwnProperty(name)) {
				this.store[name] = 1;
				return true;
			}

			return false;
		};
		this.reset = function()
		{
			this.store = {};
		};
	}
	let CounterStore = new CounterStoreFn();

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

			return "\tdoc.GetRangeBySelect().GetAllParagraphs().forEach(para => para.SetSpacingLine(" + value + " * 240, \"" + type + "\"));\n"
		},
		// incIndentetLineSpacing	: function(){
		// 	// for now we don't have relative increaee/decrease for api
		// 	//"\tdoc.GetRangeBySelect().GetAllParagraphs().forEach(para => para.SetIndFirstLine());\n"
		// 	//paragraph.SetIndFirstLine(1440);
		// },
		setParagraphNumbering	: function(num){
			return "\tlet " + CounterStore.inc('numbering') + " = doc.CreateNumbering(\"" + num.Type + "\");\n"
				+ "\tdoc.GetRangeBySelect().GetAllParagraphs().forEach(para => {\n\t\tpara.SetNumbering(" + CounterStore.get('numbering') + ".GetLevel(0));\n\t\tpara.SetContextualSpacing(true)\n\t});\n"
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
			return "\t(function () {\n"
				+ "\t\tlet tableStyle = doc.GetStyle(\"" + (prop.style ? prop.style : "") + "\");\n"
				+ "\t\tlet table = Api.CreateTable(" + prop.col + ", " + prop.row + ");\n"
				+ "\t\tif (tableStyle)\n"
				+	"\t\t\ttable.SetStyle(tableStyle);\n"
				+ "\t\ttable.SetWidth(\"percent\", 100);\n"
				+ "\t\tdoc.Push(table);\n"
				+ "\t\ttable.GetRange(0, 0).Select();\n"
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
				text += "\tlet " + CounterStore.inc('image') + " = Api.CreateImage(\"" + image.src + "\", " + size.wEMU + ", " + size.hEMU + ");\n"
				text += "\tdoc.GetCurrentParagraph().AddDrawing("+ CounterStore.get('image') + ");\n"
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

			let value = "\tlet " + CounterStore.inc('chart') + " = Api.CreateChart(\n"
				+ "\t\t\"" + chartType + "\",\n"
				+ "\t\t" + JSON.stringify(seriesData) + ",\n"
				+ "\t\t" + JSON.stringify(seriesNames) + ",\n"
				+ "\t\t" + JSON.stringify(categories) + ",\n"
				+ "\t\t" + width + ",\n"
				+ "\t\t" + height + ",\n"
				+ "\t\t" + style + "\n"
				+ "\t);\n"
			//+ "\t" + CounterStore.get('chart') + ".SetTitle(\"" + title + "\", " + 14 + ");\n"
			+ "\tdoc.GetCurrentParagraph().AddDrawing(" + CounterStore.get('chart') + ");\n"
			return value;
		},
		addHyperlink			: function(hl){
			// create hyperlink text
			return "\tdoc.GetRangeBySelect().AddHyperlink(\"" + (hl.Value ? hl.Value : hl.Text) + "\", \"" + hl.ToolTip + "\");\n"
		},
		addShape				: function(shapeProps){
			// for now we don't have api for move shape in exact position
			let fill = shapeProps.fill.getRGBAColor();
			let border = shapeProps.border;
			let borderwidth = border.w / 36000;
			let borderColor = border.Fill.getRGBAColor();
			return "\t(function () {\n" +
					"\t\tlet fill = Api.CreateSolidFill(Api.CreateRGBColor("+ fill.R +", " + fill.G + ", " + fill.B + "));\n" +
					"\t\tlet stroke = Api.CreateStroke(" + borderwidth +"* 36000, Api.CreateSolidFill(Api.CreateRGBColor("+ borderColor.R +", " + borderColor.G + ", " + borderColor.B + ")));\n" +
					"\t\tlet shape = Api.CreateShape(\"" + shapeProps.type + "\", " + shapeProps.extX + " * 36000, " + shapeProps.extY + " * 36000, fill, stroke);\n" +
					//"\t\tshape.SetPosition(" + shapeProps.pos.x + " * 36000.0, " + shapeProps.pos.y + " * 36000.0)\n" +
					"\t\tdoc.GetCurrentParagraph().AddDrawing(shape);\n" +
				"\t}());\n";
		},
		removeHdr				: function(hdr){
			if (hdr.isHeader)
				return "\tdoc.GetFinalSection().RemoveHeader(\"default\");\n";
			else
				return "\tdoc.GetFinalSection().RemoveFooter(\"default\");\n";
		},
		addComment				: function(commentData){
			return "\tlet " + CounterStore.inc('comment') + " = doc.AddComment(\"" + commentData.m_sText + "\", \"" + commentData.m_sUserName + "\", \"" + commentData.m_sUserId + "\");\n"
				+ "\t" + CounterStore.get('comment') + ".SetTime(" + commentData.m_sTime + ")\n";
			// todo add time
		},
		addMath					: function(type){
			let paraMath = new AscCommonWord.ParaMath();
			paraMath.Root.Load_FromMenu(type);
			paraMath.Root.Correct_Content(true);

			return "\tdoc.AddMathEquation(\"" + paraMath.GetTextOfElement().GetText() + "\", \"unicode\");\n"
		},
		addMathHotkey			: function(obj){
			let type = 'unicode';
			if (obj.type === 1)
				type === "latex";
			else if (obj.type === 2)
				type === "mathml"

			return "\tdoc.AddMathEquation(\"" + obj.math + "\", \"" + type + "\");\n";
		},
		addBlockContentControl	: function(){
			return "\tdoc.Push(Api.CreateBlockLvlSdt());\n";
		},
		addInlineContentControl	: function(){
			return "\tdoc.GetCurrentParagraph().AddInlineLvlSdt((Api.CreateInlineLvlSdt()));\n";
		},
		addContentControlList	: function(props){
			if (props === true)
				return "\tdoc.AddComboBoxContentControl();\n";
			else if (props === false)
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
		}
	};

	const WordActionsMacroList = {};
	// home tab and general changes
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextBold]				= wordActions.setTextBold;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextBoldHotKey]			= wordActions.setTextBold;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextItalic]				= wordActions.setTextItalic;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextItalicHotKey]		= wordActions.setTextItalic;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextUnderline]			= wordActions.setTextUnderline;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextUnderlineHotKey]		= wordActions.setTextUnderline;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextStrikeout]			= wordActions.setTextStrikeout;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextStrikeoutHotKey]		= wordActions.setTextStrikeout;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextFontName]			= wordActions.setTextFontName;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextFontSize]			= wordActions.setTextFontSize;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextHighlightColor]		= wordActions.setTextHighlightNone;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextHighlightNone]		= wordActions.setTextHighlightColor;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextHighlight]			= wordActions.setTextHighlightColor;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextVertAlignHotKey2]	= wordActions.setTextVertAlign;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextVertAlignHotKey3]	= wordActions.setTextVertAlign;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextVertAlignHotKey]		= wordActions.setTextVertAlign;
	WordActionsMacroList[AscDFH.historydescription_Document_SetTextColor]				= wordActions.setTextColor;
	WordActionsMacroList[AscDFH.historydescription_Document_SetStyleHeading]			= wordActions.setStyleHeading;
	WordActionsMacroList[AscDFH.historydescription_Document_SetParagraphStyle]			= wordActions.setStyleHeading;
	WordActionsMacroList[AscDFH.historydescription_Document_Shortcut_ClearFormatting]	= wordActions.clearFormat;
	WordActionsMacroList[AscDFH.historydescription_Document_ClearFormatting]			= wordActions.clearFormat;
	WordActionsMacroList[AscDFH.historydescription_Cut]									= wordActions.cut;
	WordActionsMacroList[AscDFH.historydescription_Document_ChangeTextCase]				= wordActions.changeTextCase;
	WordActionsMacroList[AscDFH.historydescription_Document_AddLetter]					= wordActions.addLetter;
	WordActionsMacroList[AscDFH.historydescription_Document_SetParagraphAlign]			= wordActions.setAlign;
	WordActionsMacroList[AscDFH.historydescription_Document_SetParagraphAlignHotKey]	= wordActions.setAlign;
	WordActionsMacroList[AscDFH.historydescription_Document_SetParagraphShd]			= wordActions.setParagraphShd;
	WordActionsMacroList[AscDFH.historydescription_Document_SetParagraphLineSpacing]	= wordActions.setLineSpacing;
	//WordActionsMacroList[AscDFH.historydescription_Document_IncParagraphIndent]		= wordActions.incIndentetLineSpacing;
	//WordActionsMacroList[AscDFH.historydescription_Document_DecParagraphIndent]		= wordActions.decIndentetLineSpacing;
	//WordActionsMacroList[AscDFH.historydescription_Document_IncFontSize]				= wordActions.incFontSize;
	WordActionsMacroList[AscDFH.historydescription_Document_SetParagraphNumbering]		= wordActions.setParagraphNumbering;
	WordActionsMacroList[AscDFH.historydescription_Document_SetParagraphNumberingHotKey]= wordActions.setParagraphNumbering;
	WordActionsMacroList[AscDFH.historydescription_Document_AddParagraph]				= wordActions.addParagraph;
	//WordActionsMacroList[AscDFH.historydescription_Document_AddPageNumHotKey]			= wordActions.addPageNum;
	// WordActionsMacroList[AscDFH.historydescription_Document_FormatPasteHotKey]		= wordActions;
	// WordActionsMacroList[AscDFH.historydescription_Document_PasteHotKey]				= wordActions;
	// WordActionsMacroList[AscDFH.historydescription_Document_PasteSafariHotKey]		= wordActions;
	// WordActionsMacroList[AscDFH.historydescription_Document_CutHotKey]				= wordActions;
	WordActionsMacroList['moveCursorLeft']												= wordActions.moveCursorLeft;
	WordActionsMacroList['moveCursorRight']												= wordActions.moveCursorRight;
	WordActionsMacroList['moveCursorUp']												= wordActions.moveCursorUp;
	WordActionsMacroList['moveCursorDown']												= wordActions.moveCursorDown;
	WordActionsMacroList['remove']														= wordActions.remove;

	//WordActionsMacroList[AscDFH.historydescription_Document_BackSpaceButton]			= wordActions.backSpaceButton;
	// WordActionsMacroList[AscDFH.historydescription_Document_DeleteButton]			= wordActions.deleteButton;
	
	// input tab
	WordActionsMacroList[AscDFH.historydescription_Document_AddBlankPage]				= wordActions.addBlankPage;
	WordActionsMacroList[AscDFH.historydescription_Document_AddPageBreak]				= wordActions.addPageBreak;
	WordActionsMacroList[AscDFH.historydescription_Document_AddSectionBreak]			= wordActions.addSectionBreak;
	WordActionsMacroList[AscDFH.historydescription_Document_AddTable]					= wordActions.addTable;
	WordActionsMacroList[AscDFH.historydescription_Document_ApplyTablePr]				= wordActions.applyTablePr;
	WordActionsMacroList[AscDFH.historydescription_Document_AddChart]					= wordActions.addChart;
	WordActionsMacroList[AscDFH.historydescription_Document_AddImages]					= wordActions.addImage;
	WordActionsMacroList[AscDFH.historydescription_Document_AddHyperlink]				= wordActions.addHyperlink;
	WordActionsMacroList[AscDFH.historydescription_Document_AddNewShape]				= wordActions.addShape;
	WordActionsMacroList[AscDFH.historydescription_Document_RemoveHdrFtr]				= wordActions.removeHdr;
	WordActionsMacroList[AscDFH.historydescription_Document_AddComment]					= wordActions.addComment;
	//WordActionsMacroList[AscDFH.AscDFH.historydescription_Document_AddTextArt]		= wordActions.addTextArt;
	//WordActionsMacroList[AscDFH.AscDFH.historydescription_Document_AddDropCap]		= wordActions.addDropCap;
	//WordActionsMacroList[AscDFH.AscDFH.historydescription_Document_AddDateTimeField]	= wordActions.addDateTimeField;
	WordActionsMacroList[AscDFH.historydescription_Document_AddMath]					= wordActions.addMath;
	WordActionsMacroList[AscDFH.historydescription_Document_AddMathHotKey]				= wordActions.addMathHotkey;
	//WordActionsMacroList[AscDFH.historydescription_Document_AddTextWithProperties]	= wordActions.addMathHotkey;
	WordActionsMacroList[AscDFH.historydescription_Document_AddBlockLevelContentControl]= wordActions.addBlockContentControl;
	WordActionsMacroList[AscDFH.historydescription_Document_AddInlineLevelContentControl]= wordActions.addInlineContentControl;
	WordActionsMacroList[AscDFH.historydescription_Document_AddContentControlList]		= wordActions.addContentControlList;
	WordActionsMacroList[AscDFH.historydescription_Document_AddContentControlCheckBox]	= wordActions.addContentControlCheckBox;
	WordActionsMacroList[AscDFH.historydescription_Document_AddContentControlDatePicker]= wordActions.addContentControlDatePicker;
	WordActionsMacroList[AscDFH.historydescription_Document_AddContentControlPicture]	= wordActions.addContentControlPicture;

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
			return "\tlet " + CounterStore.inc('format') + " = Api.Format(worksheet.GetActiveCell().GetValue(), \'" + format + "\')\n"
			+ "\tworksheet.GetActiveCell().SetValue(" + CounterStore.get('format') + ");\n";
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
			let range = "\tlet " + CounterStore.inc('range') + " = Api.GetSelection().GetAddress(true, true);\n";

			if (obj.type === Asc.c_oAscSortOptions.Ascending)
				range += "\tApi.GetSelection().SetSort(" + CounterStore.get('range') + " , \"xlAscending\", undefined, undefined, undefined, undefined, \"xlYes\");\n";
			else if (obj.type === Asc.c_oAscSortOptions.Descending)
				range += "\tApi.GetSelection().SetSort(" + CounterStore.get('range') + " , \"xlDescending\", undefined, undefined, undefined, undefined, \"xlYes\");\n";

			return range;
		},
		setCellEmpty			: function(){ return "\tApi.GetSelection().Clear();\n";},
		setNumberFormat			: function(format){ return "\tApi.GetSelection().SetNumberFormat(\"" + format + "\");\n";},
		setCellPaste			: function(){return "\tApi.GetSelection().Paste();\n";},
		addShape				: function(shapeProps){
			let fill = shapeProps.fill.getRGBAColor();
			let border = shapeProps.border;
			let borderwidth = border.w / 36000;
			let borderColor = border.Fill.getRGBAColor();
			let from = shapeProps.base.from;
			return "\t(function () {\n" +
					"\t\tlet fill = Api.CreateSolidFill(Api.CreateRGBColor("+ fill.R +", " + fill.G + ", " + fill.B + "));\n" +
					"\t\tlet stroke = Api.CreateStroke(" + borderwidth +"* 36000, Api.CreateSolidFill(Api.CreateRGBColor("+ borderColor.R +", " + borderColor.G + ", " + borderColor.B + ")));\n" +
					"\t\tworksheet.AddShape(\"" + shapeProps.type + "\", " + shapeProps.extX + " * 36000, " + shapeProps.extY + " * 36000, fill, stroke, " + from.col + ", " + from.colOff * 36000 + ", " + from.row + ", " + from.rowOff * 36000 + ");\n" +
				"\t}());\n";
		},
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

			return "\tlet range = worksheet.GetRangeByNumber(" + row + ", " + col + ");\n" +
				"\tlet comment = range.AddComment(\""+ comment.sText + "\");\n" +
				"\tcomment.SetAuthorName(\"" + comment.sUserName +"\");\n" +
				"\tcomment.SetUserId(\"" + comment.sUserId +"\");\n" +
				"\tcomment.SetTime(" + time +");\n"
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
		}
	};
	const CellActionsMacroList = {};
	//CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellIncreaseFontSize]	= cellActions.setCellIncreaseFontSize,
	//CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellDecreaseFontSize]	= cellActions.setCellDecreaseFontSize,
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellFontSize]				= cellActions.setCellFontSize;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellFontName]				= cellActions.setCellFontName;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellBold]					= cellActions.setCellBold;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellItalic]				= cellActions.setCellItalic;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellUnderline]			= cellActions.setCellUnderline;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellStrikeout]			= cellActions.setCellStrikeout;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellSubscript]			= cellActions.setCellSubscript;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellSuperscript]			= cellActions.setCellSuperscript;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellReadingOrder]			= cellActions.setCellReadingOrder;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellAlign]				= cellActions.setCellAlign;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellVertAlign]			= cellActions.setCellVerticalAlign;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellTextColor]			= cellActions.setCellTextColor;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellBackgroundColor]	    = cellActions.setCellBackgroundColor;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellWrap]				    = cellActions.setCellWrap;
	//CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellShrinkToFit]		= cellActions.setCellShrinkToFit;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellBorder]				= cellActions.setCellBorder;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellValue]				= cellActions.setCellValue;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellAngle]				= cellActions.setCellAngle;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellMerge]				= cellActions.setCellMerge;
	//CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellStyle]				= cellActions.setCellStyle;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellChangeTextCase]		= cellActions.setCellChangeTextCase;
	//CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellChangeFontSize]		= cellActions.setCellChangeFontSize;
	//CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellHyperlinkAdd]		= cellActions.setCellHyperlinkAdd;
	//CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellHyperlinkModify]	= cellActions.setCellHyperlinkModify;
	//CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellHyperlinkRemove]	= cellActions.setCellHyperlinkRemove;
	//CellActionsMacroList[AscDFH.historydescription_Cut]									= cellActions.cut;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellFormat]				= cellActions.setCellFormat;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellSort]					= cellActions.setCellSort;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellEmpty]				= cellActions.setCellEmpty;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellChangeDigNum]			= cellActions.setNumberFormat;
	//CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellPaste]				= cellActions.setCellPaste;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_AddShape]					= cellActions.addShape;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_AddChart]					= cellActions.addChart;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_AddComment]					= cellActions.addComment;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellHyperlink]			= cellActions.addHyperlink;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_AddImageUrls]				= cellActions.addImageUrls;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_AddAutoFilter]				= cellActions.addAutoFilter;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_RemoveAutoFilter]			= cellActions.removeAutoFilter;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SelectRange]					= cellActions.selectRange;
	CellActionsMacroList[AscDFH.historydescription_Spreadsheet_SetCellFormula]				= cellActions.setCellFormula;

	const presActions = {
		setParagraphAlign		: function(align){
			switch (align) {
				case AscCommon.align_Left:		align = 'left';		break;
				case AscCommon.align_Right:		align = 'right';	break;
				case AscCommon.align_Justify:	align = 'justify';	break;
				case AscCommon.align_Center:	align = 'center';	break;
				default:						align = 'center';
			}
			return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.SetJc(\"" + align + "\"));\n\t})\n";
		},
		paragraphAdd			: function(additional){
			if (!additional.length)
				return "";

			let text = "";
			for (let nChar = 0; nChar < additional.length; nChar++)
				text += String.fromCodePoint(additional[nChar]);

			return "\t(function () {\n"
				+ "\t\tlet shapes = Api.GetSelection().GetShapes();\n"
				+ "\t\tif (shapes.length) {\n"
				+	"\t\t\tlet content = shapes[0].GetDocContent();\n"
				+	"\t\t\tlet len = content.GetElementsCount();\n"
				+	"\t\t\tcontent.GetElement(len - 1).AddText(\"" + text + "\")\n"
				+ "\t\t}\n"
				+ "\t}());\n";
		},
		putTextPrBold			: function(bold){return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.SetBold(" + bold + "));\n\t})\n"},
		putTextPrItalic			: function(italic){return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.SetItalic(" + italic + "));\n\t})\n"},
		putTextPrUnderline		: function(underline){return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.SetUnderline(" + underline + "));\n\t})\n"},
		putTextPrStrikeout		: function(strikeout){return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.SetStrikeout(" + strikeout + "));\n\t})\n"},
		putTextPrFontName		: function(fontName){return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.SetFontName(" + fontName + "));\n\t})\n"},
		putTextPrFontSize		: function(fontsize){return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.SetFontSize(" + fontsize + "));\n\t})\n"},
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

			return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.SetVertAlign(\"" + textOfVertAlign + "\"));\n\t})\n"
		},
		addNextSlide			: function(data){
			if (data === undefined)
			{
				return "\tpresentation.AddSlide(Api.CreateSlide());\n"
			}
			else
			{
				return "\tlet " + CounterStore.inc('slide') +" = Api.CreateSlide();\n"
					+ "\tlet " + CounterStore.inc('master') +" = presentation.GetMaster(0);\n"
					+ "\tlet " + CounterStore.inc('layout') +" = " + CounterStore.get('master') + ".GetLayout(" + data + ");\n"
					+ "\t" + CounterStore.get('slide') + ".ApplyLayout(" + CounterStore.get('layout') + ");\n"
					+ "\tpresentation.AddSlide(" + CounterStore.get('slide') + ");\n";
			}
		},
		deleteSlides			: function(index){
			return "\tlet " + CounterStore.inc('slide') +" = presentation.GetSlideByIndex(" + index + ");\n"
				+ "\tif (" + CounterStore.get('slide') + ") " + CounterStore.get('slide') + ".Delete();\n";
		},
		changeLayout			: function(changeObj) {
			return "\t[" + changeObj.slides.toString() + "].forEach(index => {\n"
				+ "\t\tlet " + CounterStore.inc('slide') +" = presentation.GetSlideByIndex(index);\n"
				+ "\t\tlet " + CounterStore.inc('master') +" = presentation.GetMaster(0);\n"
				+ "\t\tlet " + CounterStore.inc('layout') +" = " + CounterStore.get('master') + ".GetLayout(" + changeObj.layout + ");\n"
				+ "\t\tif (" + CounterStore.get('slide') + ") " + CounterStore.get('slide') + ".ApplyLayout(" + CounterStore.get('layout') + ");\n"
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
			return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.SetHighlight(\"" + highlightColor + "\"));\n\t})\n"
		},
		putTextColor			: function(color){
			return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.SetColor(" + color.r + ", " + color.g + ", " + color.b + "));\n\t})\n"
		},
		clearFormatting			: function(isClear){return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.ClearFormating(" + isClear + "));\n\t})\n";},
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
			return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.SetSpacingLine(" + value + " * 240, \"" + type + "\"));\n\t})\n"
		},
		paragraphRemove			: function(args){
			return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.GetDocContent().GetContent().forEach(para => para.RemoveAllElements());\n\t})\n"
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

			return "\tApi.GetSelection().GetShapes().forEach(shape => {\n\t\tshape.SetVerticalTextAlign(\"" + typeOfVertAlign + "\");\n\t})\n";
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
			return "\tApi.GetSelection().GetShapes().forEach(shape => {shape.Ungroup()});\n"
		},
		addFlowTable			: function(table){
			return "\tconst table = Api.CreateTable(" + table.col + ", " + table.row + ");\n" +
				"\tpresentation.GetCurrentSlide().AddObject(table);\n";
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

				text += "\tlet " + CounterStore.inc('image') + " = Api.CreateImage(\"" + curImageUrl + "\", " + width + " * 36000, " + height + " * 36000);\n" +
					"\t" + CounterStore.get('image') + ".SetPosition(" + posX + " * 36000, " + posY + " * 36000);\n" +
					"\tpresentation.GetCurrentSlide().AddObject(" + CounterStore.get('image') + ");\n"
			}
			return text;
		},
		addShape				: function(shapeProps){
			let fill = shapeProps.fill.getRGBAColor();
			let border = shapeProps.border;
			let borderwidth = border.w / 36000;
			let borderColor = border.Fill.getRGBAColor();
			return "\t(function () {\n" +
					"\t\tlet fill = Api.CreateSolidFill(Api.CreateRGBColor("+ fill.R +", " + fill.G + ", " + fill.B + "));\n" +
					"\t\tlet stroke = Api.CreateStroke(" + borderwidth +"* 36000, Api.CreateSolidFill(Api.CreateRGBColor("+ borderColor.R +", " + borderColor.G + ", " + borderColor.B + ")));\n" +
					"\t\tlet shape = Api.CreateShape(\"" + shapeProps.type + "\", " + shapeProps.extX + " * 36000, " + shapeProps.extY + " * 36000, fill, stroke);\n" +
					"\t\tshape.SetPosition(" + shapeProps.pos.x + " * 36000 , " + shapeProps.pos.y + " * 36000 );\n" +
					"\t\tpresentation.GetCurrentSlide().AddObject(shape);\n" +
				"\t}());\n";
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

			let value = "\tlet " + CounterStore.inc('chart') + " = Api.CreateChart(\n"
			+ "\t\t\"" + chartType + "\",\n"
			+ "\t\t" + JSON.stringify(seriesData) + ",\n"
			+ "\t\t" + JSON.stringify(seriesNames) + ",\n"
			+ "\t\t" + JSON.stringify(categories) + ",\n"
			+ "\t\t" + width + ",\n"
			+ "\t\t" + height + ",\n"
			+ "\t\t" + style + "\n"
			+ "\t);\n"
			+ "\t" + CounterStore.get('chart') + ".SetTitle(\"" + title + "\", " + 14 + ");\n"
			+ "\t" + CounterStore.get('chart') + ".SetPosition("+ chart.x + " * 36000, " + chart.y +" * 36000);\n"
			+ "\tpresentation.GetCurrentSlide().AddObject(" + CounterStore.get('chart') + ");\n"
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
			return "\t(function () {\n"
			+ "\t\tlet shapes = Api.GetSelection().GetShapes();\n"
			+ "\t\tif (shapes.length)\n"
			+	"\t\t\tshapes[0].GetDocContent().Push(Api.CreateParagraph());\n"
			+ "\t}());\n";
		}
	};

	// alignTo no api
	// merge shapes no api
	// show from start/n-slide ... when add api

	const PresentationActionMacroList = {};
	PresentationActionMacroList[AscDFH.historydescription_Presentation_ParagraphAdd] 				= presActions.paragraphAdd;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_PutTextPrBold]				= presActions.putTextPrBold;
	PresentationActionMacroList[AscDFH.historydescription_Document_SetTextBoldHotKey]				= presActions.putTextPrBold;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_PutTextPrItalic]				= presActions.putTextPrItalic;
	PresentationActionMacroList[AscDFH.historydescription_Document_SetTextItalicHotKey]				= presActions.putTextPrItalic;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_PutTextPrUnderline]			= presActions.putTextPrUnderline;
	PresentationActionMacroList[AscDFH.historydescription_Document_SetTextUnderlineHotKey]			= presActions.putTextPrUnderline;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_PutTextPrStrikeout]			= presActions.putTextPrStrikeout;
	PresentationActionMacroList[AscDFH.historydescription_Document_SetTextStrikeoutHotKey]			= presActions.putTextPrStrikeout;
	PresentationActionMacroList[AscDFH.historydescription_Document_SetTextVertAlign]				= presActions.setTextVertAlign;
	PresentationActionMacroList[AscDFH.historydescription_Document_SetTextVertAlignHotKey3]			= presActions.setTextVertAlign;
	PresentationActionMacroList[AscDFH.historydescription_Document_SetTextVertAlignHotKey2]			= presActions.setTextVertAlign;
	PresentationActionMacroList[AscDFH.historydescription_Document_SetTextHighlight]				= presActions.setTextHighlight;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_PutTextColor]				= presActions.putTextColor;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_ParagraphClearFormatting]	= presActions.clearFormatting;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_PutTextPrFontName]			= presActions.putTextPrFontName;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_PutTextPrFontSize]			= presActions.putTextPrFontSize;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_SetParagraphAlign]			= presActions.setParagraphAlign;
	PresentationActionMacroList[AscDFH.historydescription_Document_SetParagraphAlignHotKey]			= presActions.setParagraphAlign;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_AddNextSlide]				= presActions.addNextSlide;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_DeleteSlides]				= presActions.deleteSlides;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_ChangeLayout]				= presActions.changeLayout;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_SetVerticalAlign]			= presActions.setVerticalAlign;
	// PresentationActionMacroList[AscDFH.historydescription_Presentation_BringForward]				= presActions.bringForward;
	// PresentationActionMacroList[AscDFH.historydescription_Presentation_BringToFront]				= presActions.bringToFront;
	// PresentationActionMacroList[AscDFH.historydescription_Presentation_BringBackward]			= presActions.bringBackward;
	// PresentationActionMacroList[AscDFH.historydescription_Presentation_SendToBack]				= presActions.sendToBack;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_CreateGroup]					= presActions.createGroup;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_UnGroup]						= presActions.unGroup;
	//PresentationActionMacroList[AscDFH.historydescription_Presentation_PutTextPrIncreaseFontSize]	= presActions.putTextPrIncreaseFontSize;
	//PresentationActionMacroList[AscDFH.historydescription_Presentation_ParagraphIncDecFontSize]	= presActions.incDecFontSize;
	//PresentationActionMacroList[AscDFH.historydescription_Presentation_SetParagraphNumbering]		= presActions.setNumbering;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_PutTextPrLineSpacing]		= presActions.putTextPrLineSpacing;
	//PresentationActionMacroList[AscDFH.historydescription_Spreadsheet_Remove]						= presActions.paragraphRemove; // stange
	PresentationActionMacroList[AscDFH.historydescription_Presentation_AddFlowTable]				= presActions.addFlowTable;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_AddFlowImage]				= presActions.addFlowImage;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_AddShape]					= presActions.addShape;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_AddChart]					= presActions.addChart;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_AddComment]					= presActions.addComment;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_HyperlinkAdd]				= presActions.addHyperlink;
	PresentationActionMacroList[AscDFH.historydescription_Presentation_AddNewParagraph]				= presActions.addParagraph;

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
