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

(function (undefined) {
	function CInsertDocumentManager(api) {
		this.api = api;
		this.convertedFiles = [];
	}
	CInsertDocumentManager.prototype.closeConvertedFiles = function () {
			for (let i = 0; i < this.convertedFiles.length; i++) {
				this.convertedFiles[i]["close"]();
			}
	};
	CInsertDocumentManager.prototype.insertDocuments_local = function (files) {
		const oThis = this;
		const api = this.api;
		const fPromises = [];
		for (let i = 0; i < files.length; i += 1) {
			const fPromise = function () {
				return new Promise(function (resolve) {
					window["AscDesktopEditor"]["convertFile"](files[i], Asc.c_oAscFileType.CANVAS_WORD, function (file) {
						if (file) {
							oThis.convertedFiles.push(file);
							const stream = file["get"]();
							const imageMap = file["getImages"]();
							if (stream) {
								resolve({stream: new Uint8Array(stream), imageMap: imageMap});
							} else {
								resolve(null);
							}
						} else {
							resolve(null);
						}
					});
				});
			};
			fPromises.push(fPromise);
		}
		const promiseIterator = new AscCommon.CPromiseGetterIterator(fPromises);
		promiseIterator.forAllSuccessValues(function (streamInfos) {
			const filterStreams = streamInfos.filter(function (streamInfo) {
				return !!streamInfo;
			});
			if (filterStreams.length === streamInfos.length) {
				oThis.insertDocuments(streamInfos);
			} else {
				api.sendEvent("asc_onError", Asc.c_oAscError.ID.UplDocumentExt, Asc.c_oAscError.Level.NoCritical);
				oThis.endLongAction();
			}
		});
	};
	CInsertDocumentManager.prototype.startLongAction = function () {
		this.api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
	};
	CInsertDocumentManager.prototype.endLongAction = function () {
		this.closeConvertedFiles();
		this.api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
	};
	CInsertDocumentManager.prototype.insertTextFromFile_local = function () {
		const oThis = this;
		window["AscDesktopEditor"]["OpenFilenameDialog"]("word", true, function (files) {
			oThis.startLongAction();
			if (!Array.isArray(files)) {
				files = [files];
			}

			const preparingFiles = files.filter(function (file) {
				return !!file;
			});

			oThis.insertDocuments_local(preparingFiles);
		});
	};

	CInsertDocumentManager.prototype.checkLocked = function () {
		const logicDocument = this.getLogicDocument();
		return logicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, logicDocument.IsFormFieldEditing());
	};
	CInsertDocumentManager.prototype.startAction = function () {
		const logicDocument = this.getLogicDocument();
		logicDocument.StartAction(AscDFH.historydescription_Document_InsertTextFromFile);
	};
	CInsertDocumentManager.prototype.finalizeAction = function () {
		const logicDocument = this.getLogicDocument();

		logicDocument.Recalculate();
		logicDocument.UpdateSelection();
		logicDocument.UpdateInterface();
		logicDocument.UpdateRulers();
		logicDocument.UpdateTracks();
		logicDocument.FinalizeAction();
	};

	CInsertDocumentManager.prototype.checkSelectionBeforePaste = function () {
		const logicDocument = this.getLogicDocument();
		if (logicDocument.IsTableCellSelection()) {
			logicDocument.MoveCursorLeft(false, false);
		}
	};

	CInsertDocumentManager.prototype.insertDocuments = function (streamInfos) {
		if (this.checkLocked() || !streamInfos.length) {
			this.endLongAction();
			return;
		}

		this.startAction();

		this.checkSelectionBeforePaste();

		const fPromises = [];
		const oThis = this;
		const api = this.api;
		const insertDocumentUrlsData = {imageMap: null, documents: [], convertCallback: function (_api, url) {}, endCallback: function (_api) {}};
		for (let i = 0; i < streamInfos.length; i++) {
			const stream = streamInfos[i].stream;
			const imageMap = streamInfos[i].imageMap;
			fPromises.push(function () {
				api.insertDocumentUrlsData = insertDocumentUrlsData;
				insertDocumentUrlsData.imageMap = imageMap;
				return new Promise(function (resolve) {
					oThis.pasteData(stream, resolve);
				});
			});
		}

		const promiseFunctionIterator = new AscCommon.CPromiseGetterIterator(fPromises);
		promiseFunctionIterator.forAllSuccessValues(function () {
			oThis.finalizeAction();
			api.endInsertDocumentUrls();
			oThis.endLongAction();
		});
	};
	CInsertDocumentManager.prototype.convertDocuments = function (resultDocuments, isUseDirectUrlError) {
		if (!resultDocuments.length) {
			this.endLongAction();
			return;
		}

		const oThis = this;
		const streamInfos = [];
		this.api._ConvertDocuments(resultDocuments.slice(), !!isUseDirectUrlError, function (stream, imageMap) {
			streamInfos.push({stream: stream, imageMap: imageMap});
		}, function (api) {
			if (streamInfos.length === resultDocuments.length) {
				oThis.insertDocuments(streamInfos);
			} else {
				oThis.endLongAction();
			}
		});
	};
	CInsertDocumentManager.prototype.insertTextFromFile = function () {
		const oThis = this;
		const api = oThis.api;

		if (api.isLocalMode()) {
			this.insertTextFromFile_local();
			return;
		}


		AscCommon.ShowDocumentFileDialog(function (error, files) {
			if (Asc.c_oAscError.ID.No !== error) {
				api.sendEvent("asc_onError", error, Asc.c_oAscError.Level.NoCritical);
				return;
			}
			oThis.startLongAction();
			const promises = [];
			for (let i = 0; i < files.length; i += 1) {
				promises.push(new Promise(function (resolve, reject) {
					const format = AscCommon.GetFileExtension(files[i].name);
					const reader = new FileReader();
					reader.onload = function () {
						const fileData = {data: new Uint8Array(reader.result), format: format};
						resolve(fileData);
					};
					reader.onerror = function () {
						reject(null);
					};
					reader.readAsArrayBuffer(files[i]);
				}));
			}
			Promise.all(promises).then(oThis.convertDocuments.bind(oThis)).catch(function () {
				api.sendEvent("asc_onError", Asc.c_oAscError.ID.Unknown, Asc.c_oAscError.Level.NoCritical);
				oThis.endLongAction();
			});
		}, true);
	};

	CInsertDocumentManager.prototype.insertTextFromUrl = function (url, token) {
		this.startLongAction();
		if (this.api.isLocalMode()) {
			this.insertTextFromUrl_local(url, token);
			return;
		}
		this.convertDocuments([{url: url, format: "docx", token: token}], true);

	};
	CInsertDocumentManager.prototype.insertTextFromUrl_local = function (url, token) {
		const oThis = this;
		const api = this.api;
		const urls = [url];
		window["AscDesktopEditor"]["DownloadFiles"](urls, [], function (files) {
			const preparingFiles = [];
			for (let i = 0; i < urls.length; i += 1) {
				const tmp = files[urls[i]];
				if (tmp) {
					preparingFiles.push(tmp);
				} else {
					api.sendEvent("asc_onError", Asc.c_oAscError.ID.DirectUrl, Asc.c_oAscError.Level.NoCritical);
					oThis.endLongAction();
					return;
				}
			}
			oThis.insertDocuments_local(preparingFiles);
		});
	};

	CInsertDocumentManager.prototype.getLogicDocument = function () {
		return this.api.WordControl.m_oLogicDocument;
	};

	CInsertDocumentManager.prototype.pasteData = function (stream, resolve) {
		this.api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Internal, stream, undefined, undefined, undefined, function () {resolve();}, false, function () {resolve();});
	};

	AscCommonWord.CInsertDocumentManager = CInsertDocumentManager;
})();
