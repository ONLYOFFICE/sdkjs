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

(function(window)
{
	const EVENT_TIMEOUT = 1000;
	
	/**
	 * @param editor
	 * @constructor
	 */
	function TextAnnotatorEventManager(editor)
	{
		this.editor = editor;
		
		this.logicDocument = null;
		this.textAnnotator = null;
		
		this.sendPara = {};
		this.waitPara = {};
		
		this.curParagraph = null;
		this.curRanges    = null;
		
		this.plugins = {};
		this.startedPlugins = {}; // По ключу-гуиду плагину храним параграфы, которые нужно обработать после загрузки плагина
	}
	TextAnnotatorEventManager.prototype.init = function()
	{
		if (this.logicDocument)
			return;
		
		this.logicDocument = this.editor.private_GetLogicDocument();
		this.textAnnotator = this.logicDocument.CustomTextAnnotator;
	};
	/**
	 * @param {AscWord.Paragraph} paragraph
	 */
	TextAnnotatorEventManager.prototype.onChangeParagraph = function(paragraph)
	{
		this.init();
		
		let paraId = paragraph.GetId();
		
		this.sendPara[paraId] = this.logicDocument.GetRecalcId();
		
		// Не посылаем сразу сообщение, чтобы не посылать их на каждое действие при быстром наборе
		if (this.waitPara[paraId])
			clearTimeout(this.waitPara[paraId]);
		
		let _t = this;
		this.waitPara[paraId] = setTimeout(function(){
			_t.send(paragraph);
			_t.waitPara[paraId] = null;
		}, EVENT_TIMEOUT);
	};
	TextAnnotatorEventManager.prototype.send = function(paragraph)
	{
		if (!this.hasPluginListeners())
			return;
		
		// TODO: Надо проверить, нужно ли вообще ивент отправлять
		
		// Чтобы не было моргания при быстром изменении параграфа, мы не чистим метки сразу при изменении.
		// Поэтому, отправляем в плагин новые позиции меток, чтобы плагин сам решал обновить всю разметку или
		// просто удалить какие-то метки
		
		let obj = this.textAnnotator.getEventObject(paragraph);
		let paraId = paragraph.GetId();
		let recalcId = this.sendPara[paraId];
		
		let text = obj.text;
		
		for (let pluginGuid in this.startedPlugins)
		{
			this.startedPlugins[pluginGuid].RemoveParagraph(paraId);
		}
		
		let objByGuid = {};
		for (let handlerId in obj.ranges)
		{
			let _ranges = obj.ranges[handlerId];
			
			let guid = this.getGuid(handlerId);
			if (!objByGuid[guid])
			{
				objByGuid[guid] = {
					"paragraphId" : paraId,
					"recalcId"    : recalcId,
					"text"        : text,
					"annotations" : []
				};
			}
			
			let handlerName = this.getHandlerName(handlerId);
			
			for (let rangeId in _ranges)
			{
				let start = _ranges[rangeId].start;
				let end   = _ranges[rangeId].end;
				
				if (undefined === start || undefined === end)
					continue;
				
				let range = {
					"id" : rangeId,
					"start" : start,
					"length" : end - start
				};
				if (handlerName)
					range["name"] = handlerName;
				
				objByGuid[guid]["annotations"].push(range);
			}
		}
		
		let excludeGuids = {};
		for (let guid in objByGuid)
		{
			let guids = {};
			guids[guid] = true;
			window.g_asc_plugins.onPluginEvent2("onParagraphText", objByGuid[guid], guids);
			
			excludeGuids[guid] = 1;
		}
		
		let _obj = {
			"paragraphId" : paraId,
			"recalcId"    : recalcId,
			"text"        : text
		};
		
		window.g_asc_plugins.onPluginEvent2("onParagraphText", _obj, null, false, false, excludeGuids);
	};
	TextAnnotatorEventManager.prototype.onResponse = function(obj)
	{
		this.init();
		
		if (!obj)
			return;
		
		let name = obj["name"];
		if (name)
		{
			if (this.plugins[obj["guid"]])
				this.plugins[obj["guid"]][name] = true;
			
			obj["guid"] += "AnnotationName:" + name;
		}
		
		switch (obj["type"])
		{
			case "highlightText":
			{
				let guid     = obj["guid"];
				let paraId   = obj["paragraphId"];
				let recalcId = obj["recalcId"];
				let ranges   = obj["ranges"];
				
				if (undefined === guid
					|| this.sendPara[paraId] !== recalcId
					|| !ranges
					|| !Array.isArray(ranges))
					return;
				
				let _ranges = [];
				for (let i = 0; i < ranges.length; ++i)
				{
					let _r = this.parseHighlightTextRange(ranges[i]);
					if (_r)
						_ranges.push(_r);
				}
				
				this.logicDocument.CustomTextAnnotator.highlightTextResponse(guid, paraId, _ranges);
				break;
			}
		}
	};
	TextAnnotatorEventManager.prototype.parseHighlightTextRange = function(obj)
	{
		if (!obj || undefined === obj["start"] || undefined === obj["length"] || undefined === obj["id"])
			return null;
		
		return {
			start : obj["start"],
			length : obj["length"],
			id : obj["id"]
		};
	};
	TextAnnotatorEventManager.prototype.selectRange = function(obj)
	{
		if (!obj || !obj["paragraphId"] || !obj["guid"] || !obj["rangeId"])
			return;
		
		let handlerId = this.getHandlerId(obj);
		this.textAnnotator.getMarks().selectRange(obj["paragraphId"], handlerId, obj["rangeId"]);
	};
	TextAnnotatorEventManager.prototype.removeRange = function(obj)
	{
		if (!obj || !obj["paragraphId"] || !obj["guid"] || (!obj["rangeId"] && true !== obj["all"]))
			return;
		
		let handlerId = this.getHandlerId(obj);
		
		if (true === obj["all"])
			this.textAnnotator.getMarks().removeAllRanges(obj["paragraphId"], handlerId);
		else
			this.textAnnotator.getMarks().removeRange(obj["paragraphId"], handlerId, obj["rangeId"]);
	};
	TextAnnotatorEventManager.prototype.onCurrentRanges = function(paragraph, ranges)
	{
		let prevRanges = this.curRanges;
		let prevPara   = this.curParagraph;
		
		let currPara   = paragraph;
		let currRanges = ranges;
		
		let changePara = currPara !== prevPara;
		
		for (let handlerId in prevRanges)
		{
			let noHandler = !currRanges[handlerId];
			for (let rangeId in prevRanges[handlerId])
			{
				if (changePara || noHandler || !currRanges[handlerId][rangeId])
				{
					let obj = {
						"paragraphId" : prevPara.GetId(),
						"rangeId"     : rangeId
					};
					this.addNameFromHandlerId(handlerId, obj);
					let guids = {}; guids[this.getGuid(handlerId)] = true;
					window.g_asc_plugins.onPluginEvent2("onBlurAnnotation", obj, guids);
				}
			}
		}
		
		for (let handlerId in currRanges)
		{
			let noHandler = !prevRanges[handlerId];
			for (let rangeId in currRanges[handlerId])
			{
				if (changePara || noHandler || !prevRanges[handlerId][rangeId])
				{
					let obj = {
						"paragraphId" : currPara.GetId(),
						"rangeId"     : rangeId
					};
					this.addNameFromHandlerId(handlerId, obj);
					let guids = {}; guids[this.getGuid(handlerId)] = true;
					window.g_asc_plugins.onPluginEvent2("onFocusAnnotation", obj, guids);
				}
			}
		}
		
		this.curParagraph = currPara;
		this.curRanges    = currRanges;
	};
	TextAnnotatorEventManager.prototype.onClick = function(paragraph, ranges)
	{
		let paraId = paragraph.GetId();
		for (let handlerId in ranges)
		{
			let _ranges = [];
			let obj = {
				"paragraphId" : paraId,
				"ranges"      : _ranges
			};
			this.addNameFromHandlerId(handlerId, obj);
			for (let rangeId in ranges[handlerId])
			{
				_ranges.push(rangeId);
			}
			let guids = {}; guids[this.getGuid(handlerId)] = true;
			window.g_asc_plugins.onPluginEvent2("onClickAnnotation", obj, guids);
		}
	};
	TextAnnotatorEventManager.prototype.getHandlerId = function(obj)
	{
		return (obj["name"] ? obj["guid"] + "AnnotationName:" + obj["name"] : obj["guid"]);
	};
	TextAnnotatorEventManager.prototype.addNameFromHandlerId = function(handlerId, obj)
	{
		let name = this.getHandlerName(handlerId);
		if (name)
			obj["name"] = name;
	};
	TextAnnotatorEventManager.prototype.getHandlerName = function(handlerId)
	{
		let pos = handlerId.indexOf("AnnotationName:");
		if (-1 !== pos)
			return handlerId.substr(pos + 15);
		
		return null;
	};
	TextAnnotatorEventManager.prototype.getGuid = function(handlerId)
	{
		let pos = handlerId.indexOf("AnnotationName:");
		if (-1 !== pos)
			return handlerId.substr(0, pos);
		else
			return handlerId;
	};
	TextAnnotatorEventManager.prototype.addPluginListener = function(guid)
	{
		this.plugins[guid] = {};
		
		this.startedPlugins[guid] = new PluginStartupNotifier(this.textAnnotator, this, guid, this.sendPara);
		this.startedPlugins[guid].Begin();
	};
	TextAnnotatorEventManager.prototype.removePluginListener = function(guid)
	{
		if (!this.plugins[guid])
			return;
		
		for (let name in this.plugins[guid])
		{
			let handleId = this.getHandlerId({"name" : name, "guid" : guid});
			this.textAnnotator.getMarks().removeAllMarks(handleId);
		}
		
		delete this.plugins[guid];
	};
	TextAnnotatorEventManager.prototype.hasPluginListeners = function()
	{
		for (let guid in this.plugins)
		{
			return true;
		}
		
		return false;
	};
	TextAnnotatorEventManager.prototype.onEndPluginStartup = function(guid)
	{
		delete this.startedPlugins[guid];
	};
	TextAnnotatorEventManager.prototype.getRecalcId = function(paraId)
	{
		return this.sendPara[paraId];
	};
	
	/**
	 * @param {AscWord.CustomTextAnnotator} textAnnotator
	 * @param {TextAnnotatorEventManager} eventManager
	 * @param {string} guid
	 * @param {object} paragraphs
	 * @constructor
	 * @extends AscCommon.CActionOnTimerBase
	 */
	function PluginStartupNotifier(textAnnotator, eventManager, guid, paragraphs)
	{
		AscCommon.CActionOnTimerBase.call(this);
		
		this.guid          = guid;
		this.textAnnotator = textAnnotator;
		this.eventManager  = eventManager;
		this.paragraphs    = {};
		
		for (let paraId in paragraphs)
		{
			let paragraph = AscCommon.g_oTableId.GetById(paraId);
			if (paragraph)
				this.paragraphs[paraId] = paragraph;
		}
	}
	PluginStartupNotifier.prototype = Object.create(AscCommon.CActionOnTimerBase.prototype);
	PluginStartupNotifier.prototype.constructor = PluginStartupNotifier;
	PluginStartupNotifier.prototype.OnEnd = function()
	{
		this.eventManager.onEndPluginStartup(this.guid);
	};
	PluginStartupNotifier.prototype.IsContinue = function()
	{
		for (let paraId in this.paragraphs)
		{
			return true;
		}
		
		return false;
	};
	PluginStartupNotifier.prototype.DoAction = function()
	{
		let paragraph = this.Pop();
		
		let obj = this.textAnnotator.getEventObject(paragraph);
		let paraId = paragraph.GetId();
		let recalcId = this.eventManager.getRecalcId(paraId);
		
		let text = obj.text;
		
		let _obj = {
			"paragraphId" : paraId,
			"recalcId"    : recalcId,
			"text"        : text
		};
		
		let guids = {}; guids[this.guid] = true;
		window.g_asc_plugins.onPluginEvent2("onParagraphText", _obj, guids);
	};
	PluginStartupNotifier.prototype.Pop = function()
	{
		for (let paraId in this.paragraphs)
		{
			let paragraph = this.paragraphs[paraId];
			delete this.paragraphs[paraId];
			return paragraph;
		}
		
		return null;
	};
	PluginStartupNotifier.prototype.RemoveParagraph = function(paraId)
	{
		delete this.paragraphs[paraId];
	};
	
	//-------------------------------------------------------------export-----------------------------------------------
	AscCommon.TextAnnotatorEventManager = TextAnnotatorEventManager;
})(window);
