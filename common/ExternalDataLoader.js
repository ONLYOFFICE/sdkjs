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

(function (window) {
	function CExternalDataLoader(arrExternalReference, oApi, fCallback) {
		this.externalReferences = arrExternalReference || [];
		this.api = oApi;
		this.isLocalDesktop = window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]();
		this.fCallback = fCallback;
	}

	CExternalDataLoader.prototype.updateExternalData = function () {
		if (this.externalReferences) {
			if (this.isLocalDesktop) {
				this.resolveUpdateData();
			} else {
				this.api.sendEvent("asc_onUpdateExternalReference", this.externalReferences, this.tryForceSave.bind(this));
			}
		}
	};
	CExternalDataLoader.prototype.tryForceSave = function (arrData) {
		arrData = arrData || [];
		const oThis = this;
		const nTimeout = 10000;
		const arrPromises = [];
		let isForceSavePossible = true;
		for (let i = 0; i < arrData.length; i++) {
			const oData = arrData[i];
			const sToken = oData['token'];
			const sKey = oData['key'];

			//check updated file on server. compare keys. if file not updated - check only collaborative editing
			let curEr = oThis.externalReferences[i].externalReference;
			let curErKey = curEr.getKey();
			oData.notChangedFile = sKey === curEr.getKey();
			if (curErKey == null) {
				curEr.setKey(sKey);
			}

			if (!sKey) {
				//if don't have key, then don't have force save
				isForceSavePossible = false;
				break;
			}

			const oForceUpdatePromise = new Promise(function (fResolve) {
				oThis.api.saveRelativeFromChanges(sKey, sToken, nTimeout, function (bTimeout, oResult) {
					//oResult -> {code, inProgress, url}

					//errors
					//oResult.code -> c_oAscServerCommandErrors
					//c_oAscServerCommandErrors.NoError - everything all right
					//c_oAscServerCommandErrors.NotModified - no file, no changes
					//c_oAscServerCommandErrors.Token - wrong token
					//inProgress && с c_oAscServerCommandErrors.NoError -> the function was called several times and did'nt wait answer

					//url - file link

					if (!bTimeout && oResult["code"] === AscCommon.c_oAscServerCommandErrors.NoError) {
						arrData[i]["directUrl"] = oResult["url"];
						arrData[i]["url"] = oResult["url"];
						arrData[i]["fileType"] = "xlsx";
						arrData[i]["token"] = null;
						arrData[i].notChangedFile = false;
					}
					fResolve();
				});
			});
			arrPromises.push(oForceUpdatePromise);
		}
		if (!isForceSavePossible) {
			oThis.resolveUpdateData(arrData);
		} else {
			Promise.all(arrPromises).then(function () {
				oThis.resolveUpdateData(arrData);
			});
		}
	};
	CExternalDataLoader.prototype.resolveUpdateData = function (arrData) {
		arrData = arrData || [];
		const nLength = Math.max(arrData.length, this.externalReferences.length);
		const arrFPromiseGetters = [];
		for (let i = 0; i < nLength; i += 1) {
			if (this.isLocalDesktop || (arrData[i] && !arrData[i]["error"] && !arrData[i].notChangedFile)) {
				const oPromiseGetter = new CExternalDataPromiseGetter(this.api, this.getExternalReference(i), arrData[i]);
				arrFPromiseGetters.push(oPromiseGetter.getPromise.bind(oPromiseGetter));
			}
		}
		this.doUpdate(arrFPromiseGetters);
	};
	CExternalDataLoader.prototype.doUpdate = function (arrFPromiseGetters) {
		const oThis = this;
		const oPromiseGetterIterator = new AscCommon.CPromiseGetterIterator(arrFPromiseGetters);
		oPromiseGetterIterator.forAllSuccessValues(function (arrValues) {
			oThis.fCallback(arrValues);
		});
	};
	CExternalDataLoader.prototype.getExternalReference = function (nId) {
		if (this.externalReferences[nId]) {
			return this.externalReferences[nId].asc_getPath();
		}
	};

	function CExternalDataPromiseGetter(oApi, sExternalReference, oData) {
		this.externalReference = sExternalReference;
		this.data = oData;
		this.api = oApi;
		this.fileUrl = this.getFileUrl();
	}

	CExternalDataPromiseGetter.prototype.isLocalDesktop = function () {
		return window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]();
	};
	CExternalDataPromiseGetter.prototype.resolveStream = function (arrStream, fResolve) {
		fResolve({stream: arrStream, externalReferenceId: this.externalReference, data: this.data});
	};
	CExternalDataPromiseGetter.prototype.getLocalDesktopPromise = function () {
		const oThis = this;
		return new Promise(function (resolve) {
			if (oThis.fileUrl) {
				window["AscDesktopEditor"]["convertFile"](oThis.fileUrl, 0x2002, function (_file) {
					let arrStream = null;
					if (_file) {
						arrStream = new Uint8Array(_file["get"]());
						_file["close"]();
					}
					oThis.resolveStream(arrStream, resolve);
				});
			} else {
				oThis.resolveStream(null, resolve);
			}
		});
	};

	CExternalDataPromiseGetter.prototype.getLocalFileLink = function () {
		let res = this.externalReference;
		// todo should we make .replace to the local link here?
		if (res) {
			res = res.replace(/^file:\/\/\//, '');
			res = res.replace(/^file:\/\//, '');
		}
		return res;
	};

	CExternalDataPromiseGetter.prototype.isExternalLink = function () {
		const p = /^(?:http:|https:)/;
		return this.externalReference.match(p);
	};

	CExternalDataPromiseGetter.prototype.getFileUrl = function () {
		if (this.isLocalDesktop() && !this.isExternalLink()) {
			return this.getLocalFileLink();
		} else if (this.data && !this.data["error"]) {
			return this.data["url"];
		}
		return this.externalReference;
	};
	CExternalDataPromiseGetter.prototype.isSupportOOXML = function () {
		return this.api["asc_isSupportFeature"]("ooxml");
	};
	CExternalDataPromiseGetter.prototype.getPromise = function () {
		const oThis = this;
		if (this.isLocalDesktop()) {
			return this.getLocalDesktopPromise();
		} else if (!window["NATIVE_EDITOR_ENJINE"]) {
			return this.getSDKPromise();
		} else {
			return new Promise(function (fResolve) {
				oThis.resolveStream(null, fResolve);
			});
		}
	};

	CExternalDataPromiseGetter.prototype.loadFileContentFromUrl = function (sFileUrl, resolve) {
		const oThis = this;
		AscCommon.loadFileContent(sFileUrl, function (httpRequest) {
			let arrStream = null;
			if (httpRequest) {
				arrStream = AscCommon.initStreamFromResponse(httpRequest);
			}
			oThis.resolveStream(arrStream, resolve);
		}, "arraybuffer");
	};

	CExternalDataPromiseGetter.prototype.getSDKPromise = function () {
		const oThis = this;
		return new Promise(function (fResolve) {
			const isSupportOOXML = oThis.isSupportOOXML();
			const nOutputFormat = isSupportOOXML ? Asc.c_oAscFileType.XLSX : Asc.c_oAscFileType.XLSY;
			const sFileUrl = oThis.getFileUrl();
			const sFileType = oThis.data["fileType"];
			const sToken = oThis.data["token"];
			const sDirectUrl = oThis.data["directUrl"];
			const bIsXLSX = sFileType === "xlsx";

			if ((sFileUrl && !bIsXLSX) || !isSupportOOXML) {
				let bLoad = false;
				oThis.api.getConvertedXLSXFileFromUrl(sFileUrl, sFileType, sToken, nOutputFormat,
					function (sFileUrlAfterConvert) {
						if (sFileUrlAfterConvert) {
							oThis.loadFileContentFromUrl(sFileUrlAfterConvert, fResolve);
							bLoad = true;
						} else if (!bLoad) {
							oThis.resolveStream(null, fResolve);
						}
					});
			} else if (sDirectUrl || sFileUrl) {
				oThis.api._downloadOriginalFile(sDirectUrl, sFileUrl, sFileType, sToken, function (arrStream) {
					oThis.resolveStream(arrStream, fResolve);
				});
			} else {
				oThis.resolveStream(null, fResolve);
			}
		});
	};

	AscCommon.CExternalDataLoader = CExternalDataLoader;
})(window);

