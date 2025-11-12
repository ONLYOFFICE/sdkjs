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

window['AscCommon'] = window['AscCommon'] || {};

window['AscCommon']['spellcheckGetLanguages']=function(){return {"10241":{"hyphen":false,"name":"ar"},"1025":{"hyphen":false,"name":"ar"},"1026":{"hyphen":true,"name":"bg_BG"},"10266":{"hyphen":true,"name":"sr_Cyrl_RS"},"1027":{"hyphen":true,"name":"ca_ES"},"1029":{"hyphen":true,"name":"cs_CZ"},"1030":{"hyphen":true,"name":"da_DK"},"1031":{"hyphen":true,"name":"de_DE"},"1032":{"hyphen":true,"name":"el_GR"},"1033":{"hyphen":true,"name":"en_US"},"1036":{"hyphen":true,"name":"fr_FR"},"1038":{"hyphen":true,"name":"hu_HU"},"1040":{"hyphen":true,"name":"it_IT"},"1042":{"hyphen":false,"name":"ko_KR"},"1043":{"hyphen":true,"name":"nl_NL"},"1044":{"hyphen":true,"name":"nb_NO"},"1045":{"hyphen":true,"name":"pl_PL"},"1046":{"hyphen":true,"name":"pt_BR"},"1048":{"hyphen":true,"name":"ro_RO"},"1049":{"hyphen":true,"name":"ru_RU"},"1050":{"hyphen":true,"name":"hr_HR"},"1051":{"hyphen":true,"name":"sk_SK"},"1053":{"hyphen":true,"name":"sv_SE"},"1055":{"hyphen":false,"name":"tr_TR"},"1057":{"hyphen":true,"name":"id_ID"},"1058":{"hyphen":true,"name":"uk_UA"},"1060":{"hyphen":true,"name":"sl_SI"},"1062":{"hyphen":true,"name":"lv_LV"},"1063":{"hyphen":false,"name":"lt_LT"},"1066":{"hyphen":false,"name":"vi_VN"},"1068":{"hyphen":false,"name":"az_Latn_AZ"},"1069":{"hyphen":false,"name":"eu_ES"},"1087":{"hyphen":false,"name":"kk_KZ"},"1091":{"hyphen":false,"name":"uz_Latn_UZ"},"1104":{"hyphen":true,"name":"mn_MN"},"1110":{"hyphen":true,"name":"gl_ES"},"11265":{"hyphen":false,"name":"ar"},"1134":{"hyphen":false,"name":"lb_LU"},"1154":{"hyphen":false,"name":"oc_FR"},"12289":{"hyphen":false,"name":"ar"},"13313":{"hyphen":false,"name":"ar"},"14337":{"hyphen":false,"name":"ar"},"15361":{"hyphen":false,"name":"ar"},"16385":{"hyphen":false,"name":"ar"},"2049":{"hyphen":false,"name":"ar"},"2051":{"hyphen":false,"name":"ca_ES_valencia"},"2055":{"hyphen":true,"name":"de_CH"},"2057":{"hyphen":true,"name":"en_GB"},"2067":{"hyphen":true,"name":"nl_NL"},"2068":{"hyphen":true,"name":"nn_NO"},"2070":{"hyphen":true,"name":"pt_PT"},"2115":{"hyphen":false,"name":"uz_Cyrl_UZ"},"3073":{"hyphen":false,"name":"ar"},"3079":{"hyphen":true,"name":"de_AT"},"3081":{"hyphen":true,"name":"en_AU"},"3082":{"hyphen":true,"name":"es_ES"},"4097":{"hyphen":false,"name":"ar"},"4105":{"hyphen":false,"name":"en_CA"},"5121":{"hyphen":false,"name":"ar"},"6145":{"hyphen":false,"name":"ar"},"7169":{"hyphen":false,"name":"ar"},"7177":{"hyphen":false,"name":"en_ZA"},"8193":{"hyphen":false,"name":"ar"},"9217":{"hyphen":false,"name":"ar"},"9242":{"hyphen":true,"name":"sr_Latn_RS"}}};
window['AscCommon'].spellcheckGetLanguages = window['AscCommon']['spellcheckGetLanguages'];

function CSpellchecker(settings)
{
	this.useWasm = false;
	var webAsmObj = window["WebAssembly"];
	if (typeof webAsmObj === "object")
	{
		if (typeof webAsmObj["Memory"] === "function")
		{
			if ((typeof webAsmObj["instantiateStreaming"] === "function") || (typeof webAsmObj["instantiate"] === "function"))
				this.useWasm = true;
		}
	}

	this.enginePath = "./spell/";
	if (settings && settings.enginePath)
	{
		this.enginePath = settings.enginePath;
		if (this.enginePath.substring(this.enginePath.length - 1) != "/")
			this.enginePath += "/";
	}

	var dictionariesPath = "./../dictionaries";
	if (settings && settings.dictionariesPath)
	{
		dictionariesPath = settings.dictionariesPath;
		if (dictionariesPath.substring(dictionariesPath.length - 1) == "/")
			dictionariesPath = dictionariesPath.substr(0, dictionariesPath.length - 1);
	}

	this.isUseSharedWorker = !!window.SharedWorker;
	if (this.isUseSharedWorker && (false === settings.useShared))
		this.isUseSharedWorker = false;

	// disable for WKWebView
	if (this.isUseSharedWorker && (undefined !== window["webkit"]))
		this.isUseSharedWorker = false;

	this.worker = null;

	this.languages = AscCommon.spellcheckGetLanguages();

	this.stop = function()
	{
		if (!this.worker)
			return;

		try
		{
			if (this.worker.port)
				this.worker.port.close();
			else if (this.worker.terminate)
				this.worker.terminate();
		}
		catch (err)
		{
		}

		this.worker = null;
	};

	this.restartCallback = function() { console.log("restart"); }

	this.restart = function()
	{
		this.stop();

		var worker_src = this.useWasm ? "spell.js" : "spell_ie.js";
		worker_src = this.enginePath + worker_src;

		if (this.isUseSharedWorker)
		{
			try
			{
				// may be security errors
				this.worker = new SharedWorker(worker_src, "onlyoffice-spellchecker");
			}
			catch (err)
			{
				this.isUseSharedWorker = false;
				return this.restart();
			}

			this.worker.creator = this;
			this.worker.onerror = function() {
				var creator = this.creator;
				creator.worker = new Worker(worker_src);
				creator._start(creator.worker);
			};
			this._start(this.worker.port);
		}
		else
		{
			this.worker = new Worker(worker_src);

			var _t = this;

			// для "обычного воркера" - обрабатываем ошибку, чтобы он не влиял на работу редактора
			// и если ошибка из wasm модуля - то просто попробуем js версию - и рестартанем
			this.worker.onerror = function(e) {
				if (e.preventDefault)
					e.preventDefault();
				if (e.stopPropagation)
					e.stopPropagation();
				
				if (_t.useWasm)
				{
					_t.useWasm = false;
					_t.restart();
					_t.restartCallback && _t.restartCallback();
				}
			};

			this._start(this.worker);
		}
	};

	this.oncommand = function(message) { console.log(message); };

	this.checkDictionary = function(lang) {
		return (undefined !== this.languages["" + lang]) ? true : false;
	};
	this.getLanguages = function() {
		var ret = [];
		for (var lang in this.languages)
			ret.push(lang);
		return ret;
	};

	this._start = function(_port)
	{
		var _worker = this;

		_port.onmessage = function(message) {
			_worker.oncommand && _worker.oncommand(message.data);
		};
		let langs = {};
		for (let i in this.languages)
		{
			if (!this.languages.hasOwnProperty(i))
				continue;
			langs[i] = this.languages[i]["name"];
		}
		_port.postMessage({ "type" : "init", "dictionaries_path" : dictionariesPath, "languages" : langs });

		this.command = function(message)
		{
			_port && _port.postMessage(message);
		};
	};

	this.restart();
}
