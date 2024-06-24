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

(function (undefined) {
	function CInsertDocumentManager(api) {
		this.api = api;
	}

	CInsertDocumentManager.prototype.insertDocuments_local = function (files) {
		const oThis = this;
		const api = this.api;
		const fPromises = [];

		for (let i = 0; i < files.length; i += 1) {
			const fPromise = function () {
				return new Promise(function (resolve) {
					window["AscDesktopEditor"]["convertFile"](files[i], Asc.c_oAscFileType.CANVAS_WORD, function (file) {
						if (file) {
							const stream = new Uint8Array(file["get"]());
							file["close"]();
							resolve(stream);
						} else {
							resolve(null);
						}
					});
				});
			};
			fPromises.push(fPromise);
		}
		const promiseIterator = new AscCommon.CPromiseGetterIterator(fPromises);
		promiseIterator.forAllSuccessValues(function (streams) {
			const filterStreams = streams.filter(function (stream) {
				return !!stream;
			});
			if (filterStreams.length === streams.length) {
				oThis.insertDocuments(streams);
			} else {
				api.sendEvent("asc_onError", Asc.c_oAscError.ID.ConvertationOpenError, Asc.c_oAscError.Level.NoCritical);
			}
			oThis.endLongAction();
		});
	};
	CInsertDocumentManager.prototype.startLongAction = function () {
		this.api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
	};
	CInsertDocumentManager.prototype.endLongAction = function () {
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

	CInsertDocumentManager.prototype.insertDocuments = function (streams) {
		if (this.checkLocked() || !streams.length) {
			return;
		}

		this.startAction();

		this.checkSelectionBeforePaste();

		for (let i = 0; i < streams.length; i++) {
			this.pasteData(streams[i]);
		}

		this.finalizeAction();
	};
	CInsertDocumentManager.prototype.convertDocuments = function (resultDocuments, isUseDirectUrlError) {
		if (!resultDocuments.length) {
			this.endLongAction();
			return;
		}

		const oThis = this;
		const streams = [];
		this.api._ConvertDocuments(resultDocuments.slice(), !!isUseDirectUrlError, function (stream) {
			streams.push(stream);
		}, function (api) {
			if (streams.length === resultDocuments.length) {
				oThis.insertDocuments(streams);
			}
			oThis.endLongAction();
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

	CInsertDocumentManager.prototype.pasteData = function (stream) {
		this.api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Internal, stream, undefined, undefined, true, function () {}, false);
	};

	AscCommonWord.CInsertDocumentManager = CInsertDocumentManager;
})();
