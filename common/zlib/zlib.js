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

(function(window, undefined){

	var useWasm = false;
	var webAsmObj = window["WebAssembly"];
	if (typeof webAsmObj === "object")
	{
		if (typeof webAsmObj["Memory"] === "function")
		{
			if ((typeof webAsmObj["instantiateStreaming"] === "function") || (typeof webAsmObj["instantiate"] === "function"))
				useWasm = true;
		}
	}

	var enginePath = AscCommon.RELATIVE_SDKJS_PATH + "common/zlib/engine/";

	var scriptSrc = useWasm ? "zlib.js" : "zlib_ie.js";
	scriptSrc = enginePath + scriptSrc;

	window["AscCommon"].zlib_load = function(onSuccess, onError)
	{
		window["ZLibModule_onLoad"] = function()
		{
			function ZLib()
			{
				/** @suppress {checkVars} */
				this.engine = window["NATIVE_EDITOR_ENJINE"] ? CreateEmbedObject("CZipEmbed") : new AscCommon["CZLibEngineJS"]();
				this.files = [];
			}
			/**
			 * Open archive from bytes
			 * @param {Uint8Array | ArrayBuffer} buf
			 * @returns {boolean} success or not
			 */
			ZLib.prototype.open = function(buf)
			{
				if (window["NATIVE_EDITOR_ENJINE"])
				{
					this.files = this.engine["open"]((undefined !== buf.byteLength) ? new Uint8Array(buf) : buf);
					if (!this.files)
						this.files = [];
				}
				else
				{
					if (this.engine.open(buf))
						this.files = this.engine["getPaths"]();
				}
				return (this.files.length > 0) ? true : false;
			};
			/**
			 * Create new archive
			 * @returns {boolean} success or not
			 */
			ZLib.prototype.create = function()
			{
				return this.engine["create"]();
			};
			/**
			 * Save archive from current files
			 * @returns {Uint8Array | null} zip-archive bytes, or null if error
			 */
			ZLib.prototype.save = function()
			{
				return this.engine["save"]();
			};
			/**
			 * Get uncomressed file from archive
			 * @param {string} path
			 * @returns {Uint8Array | null} bytes of uncompressed data, or null if error
			 */
			ZLib.prototype.getFile = function(path)
			{
				return this.engine["getFile"](path);
			};
			/**
			 * Add uncomressed file to archive
			 * @param {string} path
			 * @param {Uint8Array | ArrayBuffer} new file in archive
			 * @returns {boolean} success or not
			 */
			ZLib.prototype.addFile = function(path, data)
			{
				return this.engine["addFile"](path, (undefined !== data.byteLength) ? new Uint8Array(data) : data);
			};
			/**
			 * Remove file from archive
			 * @param {string} path
			 * @returns {boolean} success or not
			 */
			ZLib.prototype.removeFile = function(path)
			{
				return this.engine["removeFile"](path);
			};
			/**
			 * Close & remove all used memory in archive
			 * @returns {undefined}
			 */
			ZLib.prototype.close = function()
			{
				return this.engine["close"]();
			};
			/**
			 * Get image blob for browser
			 * @returns {Blob}
			 */
			ZLib.prototype.getImageBlob = function(path)
			{
				return this.engine["getImageBlob"](path);
			};
			/**
			 * Get all file paths in archive
			 * @returns {Array}
			 */
			ZLib.prototype.getPaths = function()
			{
				return this.engine["getPaths"]();
			};
			/**
			 * Get image file raw data. this memory was copied and detach from archive.
			 * @returns {Uint8Array}
			 */
			ZLib.prototype.getImageBuffer = function(path)
			{
				return this.engine["getImageBuffer"](path);
			};

			AscCommon.ZLib = ZLib;
			AscCommon.ZLib.prototype.isModuleInit = true;

			window.nativeZlibEngine = new ZLib();
			onSuccess();
		};

		AscCommon.loadScript(scriptSrc, function(){}, onError);

		if (window["NATIVE_EDITOR_ENJINE"] === true)
			window["ZLibModule_onLoad"]();
	};

})(window, undefined);
