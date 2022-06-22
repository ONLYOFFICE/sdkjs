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

(function(window, undefined) {

	function CHashWorker(message, callback)
	{
		this.message = message;
		this.callback = callback;

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

		this.enginePath = "./../../../../sdkjs/common/hash/hash/";
		this.worker = null;

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

		this.restart = function()
		{
			this.stop();

			var worker_src = this.useWasm ? "engine.js" : "engine_ie.js";
			worker_src = this.enginePath + worker_src;

			this.worker = new Worker(worker_src);

			var _t = this;

			// обрабатываем ошибку, чтобы он не влиял на работу редактора
			// и если ошибка из wasm модуля - то просто попробуем js версию - и рестартанем
			this.worker.onerror = function(e) {

				e && e.preventDefault && e.preventDefault();
				e && e.stopPropagation && e.stopPropagation();

				if (_t.useWasm)
				{
					_t.useWasm = false;
					_t.restart();
				}
			};

			this._start(this.worker);
		};

		this.oncommand = function(message)
		{
			if (message && message["hashValue"] !== undefined)
			{
				if (this.callback)
					this.callback.call(window, message["hashValue"]);
				this.stop();
			}
		};

		this._start = function(_port)
		{
			var _worker = this;

			_port.onmessage = function(message) {
				_worker.oncommand && _worker.oncommand(message.data);
			};
			_port.postMessage({ "type" : "hash", "value" : this.message });
		};

		this.restart();
	}

	var currentHashWorker = null;

	window['AscCommon'] = window['AscCommon'] || {};

	window['AscCommon'].calculateProtectHash = function(args, callback) {
		var sendedData = [];
		for (var i = 0, len = args.length; i < len; i++)
		{
			sendedData.push({
				"password" : args[i].password,
				"salt" : args[i].salt,
				"spinCount" : args[i].spinCount,
				"alg" : args[i].alg
			});
		}

		currentHashWorker = new CHashWorker(sendedData, callback);
	};

	window['AscCommon'].HashAlgs = {
		MD2       : 0,
		MD4       : 1,
		MD5       : 2,
		RMD160    : 3,
		SHA1      : 4,
		SHA256    : 5,
		SHA384    : 6,
		SHA512    : 7,
		WHIRLPOOL : 8
	};

})(window);
