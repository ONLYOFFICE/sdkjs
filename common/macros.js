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

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined)
{
	/** @constructor */
	function CDocumentMacros()
	{
		this.Id = "_macrosGlobalId";//AscCommon.g_oIdCounter.Get_NewId();

		this.Lock = new AscCommon.CLock();

		this.Data = "";

		AscCommon.g_oTableId.Add(this, this.Id);
	}
	CDocumentMacros.prototype.SetData = function(sData)
	{
		AscCommon.History.Add(new CChangesDocumentMacrosData(this, this.Data, sData));
		this.Data = sData;
	};
	CDocumentMacros.prototype.GetData = function()
	{
		return this.Data;
	};
	CDocumentMacros.prototype.Get_Id = function()
	{
		return this.Id;
	};
	CDocumentMacros.prototype.CheckLock = function()
	{
		this.Lock.Check(this.Id);
	};
	CDocumentMacros.prototype.Write_ToBinary2 = function(Writer)
	{
		Writer.WriteLong(AscDFH.historyitem_type_DocumentMacros);

		// String2 : Id
		// String2 : Data

		Writer.WriteString2("" + this.Id);
		Writer.WriteString2(this.Data);
	};
	CDocumentMacros.prototype.Read_FromBinary2 = function(Reader)
	{
		// String2 : Id
		// String2 : Data

		this.Id   = Reader.GetString2();
		this.Data = Reader.GetString2();
	};

	CDocumentMacros.prototype.Refresh_RecalcData = function()
	{
	};

	CDocumentMacros.prototype.runAuto = function()
	{
		try
		{
			var obj = JSON.parse(this.Data);
			if (!obj["macrosArray"])
				return;
			for (var i = 0; i < obj["macrosArray"].length; i++)
			{
				if (true !== obj["macrosArray"][i]["autostart"])
					continue;

				AscCommon.safePluginEval(obj["macrosArray"][i]["value"]);
			}
		}
		catch (err)
		{
			console.error(err);
		}
	};
	CDocumentMacros.prototype.run = function(sGuid)
	{
		try
		{
			var obj = JSON.parse(this.Data);
			if (!obj["macrosArray"])
				return;
			for (var i = 0; i < obj["macrosArray"].length; i++)
			{
				if (sGuid === obj["macrosArray"][i]["guid"])
				{
					AscCommon.safePluginEval(obj["macrosArray"][i]["value"]);
					break;
				}
			}
		}
		catch (err)
		{
			console.error(err);
		}
	};
	CDocumentMacros.prototype.getGuidByName = function(sName)
	{
		try
		{
			var obj = JSON.parse(this.Data);
			if (!obj["macrosArray"])
				return;
			for (var i = 0; i < obj["macrosArray"].length; i++)
			{
				if (sName === obj["macrosArray"][i]["name"])
				{
					return obj["macrosArray"][i]["guid"];
				}
			}
		}
		catch (err)
		{
		}
		return "";
	};
	CDocumentMacros.prototype.getNameByGuid = function(sGuid)
	{
		try
		{
			var obj = JSON.parse(this.Data);
			if (!obj["macrosArray"])
				return;
			for (var i = 0; i < obj["macrosArray"].length; i++)
			{
				if (sGuid === obj["macrosArray"][i]["guid"])
				{
					return obj["macrosArray"][i]["name"];
				}
			}
		}
		catch (err)
		{
		}
		return "";
	};
	CDocumentMacros.prototype.getAllNames = function()
	{
		try
		{
			var obj = JSON.parse(this.Data);
			if (!obj["macrosArray"])
				return [];
			var oNamesMap = {};
			var aNames = [];
			for (var i = 0; i < obj["macrosArray"].length; i++)
			{
				var sName = obj["macrosArray"][i]["name"];
				if(!oNamesMap[sName])
				{
					oNamesMap[sName] = true;
					aNames.push(sName);
				}
			}
			return aNames;
		}
		catch (err)
		{
		}
		return [];
	};
    CDocumentMacros.prototype.isExistAuto = function()
    {
        try
        {
            var obj = JSON.parse(this.Data);
            if (!obj["macrosArray"])
                return;
            for (var i = 0; i < obj["macrosArray"].length; i++)
            {
                if (true === obj["macrosArray"][i]["autostart"])
                    return true;
            }
        }
        catch (err)
        {
        }
        return false;
    };

	function VbaProject() {
		this.filename = null;
		this.vbaXml = null;
		this.vbaDataXml = null;
	}
	VbaProject.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteString2(0, this.filename);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.WriteRecord2(0, this.vbaXml, function(val){
			s.WriteString2(val);
		});
		s.WriteRecord2(1, this.vbaDataXml, function(val){
			s.WriteString2(val);
		});
	};
	VbaProject.prototype.fromStream = function (s) {
		var _len = s.GetULong();
		var _start_pos = s.cur;
		var _end_pos = _len + _start_pos;
		var _at;
// attributes
		s.GetUChar();
		while (true)
		{
			_at = s.GetUChar();
			if (_at === AscCommon.g_nodeAttributeEnd)
				break;
			switch (_at)
			{
				case 0: {
					this.filename = s.GetString2();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
//members
		var _type;
		while (true)
		{
			if (s.cur >= _end_pos)
				break;
			_type = s.GetUChar();
			switch (_type)
			{
				case 0:
				{
					s.GetULong();//len
					this.vbaXml = s.GetString2();
					break;
				}
				case 1:
				{
					s.GetULong();//len
					this.vbaDataXml = s.GetString2();
					break;
				}
				default:
				{
					s.SkipRecord();
					break;
				}
			}
		}
		s.Seek2(_end_pos);
	};

	AscDFH.changesFactory[AscDFH.historyitem_DocumentMacros_Data]     = CChangesDocumentMacrosData;
	AscDFH.changesRelationMap[AscDFH.historyitem_DocumentMacros_Data] = [AscDFH.historyitem_DocumentMacros_Data];

	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseStringProperty}
	 */
	function CChangesDocumentMacrosData(Class, Old, New)
	{
		AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New);
	}
	CChangesDocumentMacrosData.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
	CChangesDocumentMacrosData.prototype.constructor = CChangesDocumentMacrosData;
	CChangesDocumentMacrosData.prototype.Type = AscDFH.historyitem_DocumentMacros_Data;
	CChangesDocumentMacrosData.prototype.private_SetValue = function(Value)
	{
		this.Class.Data = Value;
	};

	function customXMLHttpRequest() {
		this._headers = [];
		var t = this;

		this.open = function(method, url, async, user, password) {
			this._url = url;
			this._method = method;
			this._async = async;
			this._user = user;
			this._password = password;
		};

		this.setRequestHeader = function(name, value) {
			this._headers.push({name: name, value: value});
		};

		this.send = function(body) {
			setTimeout(function() {
				window.g_asc_plugins.api.asc_getUserPermissionToMakeRequestFromMacros(t._url, sendRequest);		
				function sendRequest (permission) {
					if (permission) {
						var xhr = new XMLHttpRequest();

						if (t.timeout)
							xhr.timeout = t.timeout;

						if (t.responseType)
							xhr.responseType = t.responseType;

						if ( t.hasOwnProperty('withCredentials') )
							xhr.withCredentials = t.withCredentials;

						xhr.open(t._method, t._url, t._async, t._user, t._password);

						t._headers.forEach(function(el) {
							xhr.setRequestHeader(el.name, el.value);
						});

						xhr.onload = function() {
							t.status = xhr.status;
							t.statusText = xhr.statusText;
							t.response = xhr.response;
							t.responseText = xhr.responseText;
							t.responseURL = xhr.responseURL;
							t.responseXML = xhr.responseXML;
							t.onload &&	t.onload();
						};

						xhr.onprogress = function(event) {
							t.onprogress && t.onprogress(event);
						};

						xhr.onreadystatechange = function() {
							t.readyState = this.readyState;
							t.onreadystatechange && t.onreadystatechange();
						};

						xhr.onerror = function(error) {
							t.onerror && t.onerror(error || "User doesn't allow this request.");
						};

						xhr.ontimeout = function(event) {
							t.ontimeout && t.ontimeout(event);
						};

						xhr.onloadstart = function(event) {
							t.onloadstart && t.onloadstart(event);
						};

						xhr.onloadend = function(event) {
							t.onloadend && t.onloadend(event);
						};

						xhr.onabort = function(event) {
							t.onabort && t.onabort(event);
						};

						if (typeof t.upload == 'object') {
							xhr.upload.onabort = function(event) {
								t.upload.onabort && t.upload.onabort(event);
							};

							xhr.upload.onerror = function(event) {
								t.upload.onerror && t.upload.onerror(event);
							};

							xhr.upload.onload = function(event) {
								t.upload.onload && t.upload.onload(event);
							};

							xhr.upload.onloadend = function(event) {
								t.upload.onloadend && t.upload.onloadend(event);
							};

							xhr.upload.onloadstart = function(event) {
								t.upload.onloadstart && t.upload.onloadstart(event);
							};

							xhr.upload.onprogress = function(event) {
								t.upload.onprogress && t.upload.onprogress(event);
							};

							xhr.upload.ontimeout = function(event) {
								t.upload.ontimeout && t.upload.ontimeout(event);
							};
						}

						t.getResponseHeader = function(name) {
							return xhr.getResponseHeader(name);
						};

						t.getAllResponseHeaders = function() {
							return xhr.getAllResponseHeaders();
						};

						t.abort = function() {
							xhr.abort();
						}

						xhr.send(body || null);

					} else if (t.onerror)  {
						t.onerror("User doesn't allow this request.");
					}
				}
			});
		};
	}

	window['AscCommon'] = window['AscCommon'] || {};
	window["AscCommon"].CDocumentMacros = CDocumentMacros;
	window['AscCommon'].VbaProject = VbaProject;

	var _safe_eval_closure = new Function("Function", "Api", "window", "AscDesktopEditor", "alert", "document", "XMLHttpRequest", "self", "globalThis", "setTimeout", "setInterval", "value", "return eval(\"\\\"use strict\\\";\\r\\n\" + value)");
	function _safePluginEval(value) {
		let protoFunc = Object.getPrototypeOf(function(){});
		// for minimization we use eval!!!
		let protoFuncGen = null;
		try {
			protoFuncGen = eval("Object.getPrototypeOf(function*(){})");
		} catch (err) {
			protoFuncGen = null;
		}
		const normalConstructor = protoFunc.constructor;
		const generatorNext = protoFuncGen ? protoFuncGen.prototype.next : null;
		protoFunc.constructor = function(){};
		if (protoFuncGen)
			protoFuncGen.prototype.next = function(){};

		const timeout = function(cb, delay) {
			var args = Array.prototype.slice.call(arguments, 2);
			return setTimeout(function() {
				cb.apply({}, args);
			}, delay);
		};
		const interval = function(cb, delay) {
			var args = Array.prototype.slice.call(arguments, 2);
			return setInterval(function() {
				cb.apply({}, args);
			}, delay);
		};
		const Api = window.g_asc_plugins.api;
		// clear this field on each run
		delete Api.parsedJSDoc;
		// check if we add a custom function into this macros we will parse a jsdoc
		if (value.includes('AddCustomFunction') && value.includes('@customfunction')) {
			// calculate how any times the function of adding will be called
			const countOfAdding = (value.match(/\.AddCustomFunction\(/g) || []).length;
			// parse JSDOC and put it to Api
			Api.parsedJSDoc = AscCommon.parseJSDoc(value);
			// remove extra parsed JSDOC
			if (Api.parsedJSDoc.length > countOfAdding)
				Api.parsedJSDoc.length = countOfAdding;
		}
		const result = _safe_eval_closure.call(null, {}, Api, {}, {}, function(){}, {}, customXMLHttpRequest, {}, {}, timeout, interval, value);
		protoFunc.constructor = normalConstructor;
		if (protoFuncGen)
			protoFuncGen.prototype.next = generatorNext;
		return result;
	};
	window['AscCommon'].safePluginEval = function(value){
		let result;
		let isCheckSymbols = AscFonts.IsCheckSymbols;
		AscFonts.IsCheckSymbols = true;
		try
		{
			result = _safePluginEval(value);
		}
		catch (err)
		{
			result = undefined;
			console.error(err);
		}
		AscFonts.IsCheckSymbols = isCheckSymbols;
		return result;
	};

})(window);
