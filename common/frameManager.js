/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
	function InitClassWithoutType(CMain, CBase)
	{
		CMain.prototype = Object.create(CBase.prototype);
		CMain.prototype.constructor = CMain;
	}

	const FrameEditorTypes = {
		JustBlock: 0,
		OleEditor: 1,
		ChartEditor: 2
	};

	function CFrameManagerBase(api)
	{
		this.api = api;
		this.isInitFrameManager = false;
	}
	CFrameManagerBase.prototype.clear = function () {};
	CFrameManagerBase.prototype.getWorkbookBinary = function () {};
	CFrameManagerBase.prototype.getAllImageIds = function () {};
	CFrameManagerBase.prototype.getImagesForHistory = function () {};
	CFrameManagerBase.prototype.obtain = function (oInfo) {};
	CFrameManagerBase.prototype.setGeneralDocumentUrls = function (oPr) {};
	CFrameManagerBase.prototype.getGeneralImageUrl = function (sImageId) {};
	CFrameManagerBase.prototype.openWorkbookData = function (sStream, oInfo) {};
	CFrameManagerBase.prototype.updateGeneralDiagramCache = function (aRanges) {};
	CFrameManagerBase.prototype.sendLoadImages = function (arrImages, token, bNotShowError) {};
	CFrameManagerBase.prototype.sendFromFrameToGeneralEditor = function (oSendObject) {};
	CFrameManagerBase.prototype.getAscSettings               = function (oSendObject) {};
	CFrameManagerBase.prototype.startLoadOleEditor = function () {};
	CFrameManagerBase.prototype.endLoadOleEditor = function () {};
	CFrameManagerBase.prototype.startLoadChartEditor = function () {};
	CFrameManagerBase.prototype.sendUpdateDiagram = function () {};
	CFrameManagerBase.prototype.endLoadChartEditor = function () {};
	CFrameManagerBase.prototype.isSaveZip = function () {};
	CFrameManagerBase.prototype.applyCloseCallbacks = function () {};
	CFrameManagerBase.prototype.addCloseCallback = function (fCallback) {};
	CFrameManagerBase.prototype.openLocalDesktopFileLink = function (sLocalFileLink) {};

	CFrameManagerBase.prototype.preObtain = function (oInfo) {
		this.obtain(oInfo);
	};
	CFrameManagerBase.prototype.isDiagramEditor = function ()
	{
		return false;
	};
	CFrameManagerBase.prototype.isOleEditor = function ()
	{
		return false;
	};
	CFrameManagerBase.prototype.isGeneralEditor = function ()
	{
		return false;
	};
	CFrameManagerBase.prototype.isFrameEditor = function () {
		return false;
	};

	CFrameManagerBase.prototype.getDecodedArray = function (stream)
	{
			return new Uint8Array(stream);
	};

	CFrameManagerBase.prototype.getEncodedArray = function (arrStream)
	{
		if (!arrStream) {
			return;
		}
		return Array.from(arrStream);
	};


	function CMainEditorFrameManager(api)
	{
		CFrameManagerBase.call(this, api);
		this.isInitFrameManager = true;
		this.isLoadingOleEditor = false;
		this.isLoadingChartEditor = false;
		this.chartData = null;
		this.closeCallbacks = [];
	}
	InitClassWithoutType(CMainEditorFrameManager, CFrameManagerBase);
	CMainEditorFrameManager.prototype.saveChartData = function (chartSpace) {
		if (!this.chartData) {
			const chart = chartSpace.chart;
			const chartData = chartSpace.chartData;
			this.chartData = {chart: chart, chartData: chartData};
		}
	}
	CMainEditorFrameManager.prototype.loadChartData = function (chartSpace) {
		if (this.chartData) {
			AscFormat.ExecuteNoHistory(function () {
				chartSpace.setChart(this.chartData.chart);
				chartSpace.setChartData(this.chartData.chartData);
			}, this, []);
			this.chartData = null;
		}
	}
	CMainEditorFrameManager.prototype.isGeneralEditor = function ()
	{
		return true;
	};
	CMainEditorFrameManager.prototype.isSaveZip = function ()
	{
		return this.api.isOpenOOXInBrowser;
	};
	CMainEditorFrameManager.prototype.startLoadOleEditor = function ()
	{
		this.isLoadingOleEditor = true;
	};
	CMainEditorFrameManager.prototype.endLoadOleEditor = function ()
	{
		this.isLoadingOleEditor = false;
	};

	CMainEditorFrameManager.prototype.startLoadChartEditor = function ()
	{
		this.isLoadingChartEditor = true;
	};
	CMainEditorFrameManager.prototype.endLoadChartEditor = function ()
	{
		this.isLoadingChartEditor = false;
	};
	CMainEditorFrameManager.prototype.applyCloseCallbacks = function () {
		while (this.closeCallbacks.length) {
			const fCallback = this.closeCallbacks.pop();
			fCallback();
		}
	};
	CMainEditorFrameManager.prototype.addCloseCallback = function (fCallback) {
		if (fCallback) {
			this.closeCallbacks.push(fCallback);
		}
	};
	CMainEditorFrameManager.prototype.openLocalDesktopFileLink = function(sLocalFileLink) {
		const api = this.api;
		window["AscDesktopEditor"]["openExternalReference"](sLocalFileLink, function(error) {
			let internalError = Asc.c_oAscError.ID.No;
			switch (error) {
				case 0: internalError = Asc.c_oAscError.ID.ConvertationOpenError; break;
				default: break;
			}

			if (Asc.c_oAscError.ID.No !== internalError) {
				api.sendEvent("asc_onError", internalError, Asc.c_oAscError.Level.NoCritical);
			}
		});
	};


	function CCellFrameManager(api)
	{
		CFrameManagerBase.call(this, api);
		this.generalDocumentUrls = {};
		this.isInitFrameManager = true;
		this.isSaveZipWorkbook = false;
	}
	InitClassWithoutType(CCellFrameManager, CFrameManagerBase);

	CCellFrameManager.prototype.isSaveZip = function ()
	{
		return this.isSaveZipWorkbook;
	};
	CCellFrameManager.prototype.updateOpenOnClient = function ()
	{
		this.sendFromFrameToGeneralEditor(new CFrameUpdateIsOpenOnClient());
	};
	CCellFrameManager.prototype.getWorkbookBinary = function ()
	{
		const oThis = this;
		return new Promise(function (resolve)
		{
			if (oThis.isSaveZip())
			{
				oThis.api.saveDocumentToZip(oThis.api.wb.model, AscCommon.c_oEditorId.Spreadsheet, function (data)
				{
					resolve(Array.from(data));
				});
			}
			else
			{
				const bOldFlag = window["IsEmbedImagesInInternalFormat"];
				window["IsEmbedImagesInInternalFormat"] = true;
				const oBinaryFileWriter = new AscCommonExcel.BinaryFileWriter(oThis.api.wbModel);
				const arrBinaryData = oBinaryFileWriter.Write(true, false, true);
				window["IsEmbedImagesInInternalFormat"] = bOldFlag;
				resolve(Array.from(arrBinaryData));
			}
		});

	};
	CCellFrameManager.prototype.getAllImageIds = function ()
	{
		const arrRasterImageIds = [];
		const arrWorksheetLength = this.api.wbModel.aWorksheets.length;
		for (let i = 0; i < arrWorksheetLength; i += 1)
		{
			const oWorksheet = this.api.wbModel.aWorksheets[i];
			const arrDrawings = oWorksheet.Drawings;
			if (arrDrawings)
			{
				for (let j = 0; j < arrDrawings.length; j += 1)
				{
					const oDrawing = arrDrawings[j];
					oDrawing.graphicObject.getAllRasterImages(arrRasterImageIds);
				}
			}
		}
		return arrRasterImageIds;
	};

	CCellFrameManager.prototype.getImagesForHistory = function ()
	{
		if (this.api.isOpenOOXInBrowser)
		{
			return [];
		}
		const arrRasterImageIds = this.getAllImageIds();
		const urlsForAddToHistory = [];
		for (let i = 0; i < arrRasterImageIds.length; i += 1)
		{
			const url = AscCommon.g_oDocumentUrls.mediaPrefix + arrRasterImageIds[i];
			if (!(this.generalDocumentUrls[url] && this.generalDocumentUrls[url] === AscCommon.g_oDocumentUrls.getUrls()[url]))
			{
				urlsForAddToHistory.push(arrRasterImageIds[i]);
			}
		}

		return urlsForAddToHistory;
	};

	CCellFrameManager.prototype.obtain = function (oInfo)
	{
		this.updateOpenOnClient();
		this.isInitFrameManager = false;
		let sStream = oInfo["binary"];
		if (Array.isArray(sStream))
		{
			sStream = new Uint8Array(sStream);
		}
		this.setGeneralDocumentUrls(oInfo["documentImageUrls"] || {});
		this.openWorkbookData(sStream, oInfo);
	};
	CCellFrameManager.prototype.setGeneralDocumentUrls = function (oPr)
	{
		this.generalDocumentUrls = oPr;
	};
	CCellFrameManager.prototype.getGeneralImageUrl = function (sImageId)
	{
		if (!this.api.isOpenOOXInBrowser)
		{
			return this.generalDocumentUrls[sImageId];
		}
	};
	CCellFrameManager.prototype.initBlob = function(blobUrl2Data, url2BlobUrl) {
		AscCommon.g_oDocumentBlobUrls.blobUrl2Data = blobUrl2Data || {};
		AscCommon.g_oDocumentBlobUrls.url2BlobUrl = url2BlobUrl || {};
		if (url2BlobUrl) {
			for (let sPath in url2BlobUrl) {
				const urls = {};
				const sLocalPath = sPath.slice(6);
				urls[sLocalPath] = url2BlobUrl[sPath]
				this.generalDocumentUrls[sPath] = url2BlobUrl[sPath];
			}
		}
	};
	CCellFrameManager.prototype.openWorkbookData = function (sStream, oInfo)
	{
		const oFile = new AscCommon.OpenFileResult();
		oFile.bSerFormat = AscCommon.checkStreamSignature(sStream, AscCommon.c_oSerFormat.Signature);
		oFile.data = sStream;
		oFile.isDecodedData = true;
		this.api.asc_CloseFile();
		AscCommon.g_oDocumentUrls.documentUrl = oInfo["documentUrl"];
		this.initBlob(oInfo["blobUrl2Data"], oInfo["url2BlobUrl"]);
		this.api.openDocument(oFile);
	}
	CCellFrameManager.prototype.isGeneralEditor = function ()
	{
		return false;
	};
	CCellFrameManager.prototype.isFrameEditor = function ()
	{
		return true;
	};

	CCellFrameManager.prototype.sendLoadImages = function (arrImages, token, bNotShowError)
	{
		this.sendFromFrameToGeneralEditor(new CFrameImageData(arrImages, token, bNotShowError));
	}

	CCellFrameManager.prototype.sendFromFrameToGeneralEditor = function (oSendObject)
	{
		this.api.sendFromFrameToGeneralEditor(oSendObject);
	};
	CCellFrameManager.prototype.openLocalDesktopFileLink = function(sLocalFileLink) {
		this.sendFromFrameToGeneralEditor(new CFrameOpenLocalDesktopFileLink(sLocalFileLink));
	};

	function COleCellFrameManager(api)
	{
		CCellFrameManager.call(this, api);
		this.imageWidthCoefficient = null;
		this.imageHeightCoefficient = null;
		this.isFromSheetEditor = false;
		this.isCreatingOleObject = false;
	}

	InitClassWithoutType(COleCellFrameManager, CCellFrameManager);
	COleCellFrameManager.prototype.clear = function ()
	{
		this.imageWidthCoefficient = null;
		this.imageHeightCoefficient = null;
		this.isFromSheetEditor = false;
		this.isCreatingOleObject = false;
	}
	COleCellFrameManager.prototype.getBase64Image = function ()
	{
		return this.api.wb.getImageFromTableOleObject();
	};
	COleCellFrameManager.prototype.getImageWidthCoefficient = function ()
	{
		return this.imageWidthCoefficient || undefined;
	}
	COleCellFrameManager.prototype.getImageHeightCoefficient = function ()
	{
		return this.imageHeightCoefficient || undefined;
	}
	COleCellFrameManager.prototype.getBinary = function ()
	{
		const oThis = this;
		return new Promise(function (resolve)
		{
			oThis.getWorkbookBinary().then(function (workbookBinary)
			{
				resolve({
					'binary'               : workbookBinary,
					'base64Image'          : oThis.getBase64Image(),
					'imagesForAddToHistory': oThis.getImagesForHistory(),
					'widthCoefficient'     : oThis.getImageWidthCoefficient(),
					'heightCoefficient'    : oThis.getImageHeightCoefficient()
				});
			});
		});
	};
	COleCellFrameManager.prototype.calculateImageSaveCoefficients = function (nImageHeight, nImageWidth)
	{
		if (!nImageHeight || !nImageWidth)
			return;

		const saveImageCoefficients = this.api.getScaleCoefficientsForOleTableImage(nImageWidth, nImageHeight)
		this.imageWidthCoefficient = saveImageCoefficients.widthCoefficient;
		this.imageHeightCoefficient = saveImageCoefficients.heightCoefficient;
	}
	COleCellFrameManager.prototype.obtain = function (oInfo)
	{
		this.clear();
		if (!oInfo)
		{
			oInfo = {"binary": AscCommon.getEmpty()};
			this.isCreatingOleObject = true;
		}
		this.isFromSheetEditor = oInfo["isFromSheetEditor"];
		this.calculateImageSaveCoefficients(oInfo["imageHeight"], oInfo["imageWidth"]);
		CCellFrameManager.prototype.obtain.call(this, oInfo);
	}
	COleCellFrameManager.prototype.isFromSheetEditor = function ()
	{
		return this.isFromSheetEditor;
	}

	COleCellFrameManager.prototype.setAfterLoadCallback = function ()
	{

		const oApi = this.api;
		const oThis = this;
		oApi.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Open);
		// на случай, если изображение поставили на загрузку, закрыли редактор, и потом опять открыли
		oApi.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadImage);

		this.sendFromFrameToGeneralEditor({
			"type": AscCommon.c_oAscFrameDataType.OpenFrame,
			"information": {"editorType": FrameEditorTypes.OleEditor}
		});
		oApi.fAfterLoad = function ()
		{
			if (oThis.isCreatingOleObject)
			{
				AscFormat.ExecuteNoHistory(function ()
				{
					const oFirstWorksheet = oThis.api.wbModel.getWorksheet(0);
					if (oFirstWorksheet)
					{
						oFirstWorksheet.sName = '';
						const sName = oThis.api.wbModel.getUniqueSheetNameFrom(AscCommon.translateManager.getValue(AscCommonExcel.g_sNewSheetNamePattern), false);
						oFirstWorksheet.setName(sName);
						oThis.api.sheetsChanged();
					}
				}, oThis, []);
			}
			oThis.api.wb.scrollToOleSize();
			// добавляем первый поинт после загрузки, чтобы в локальную историю добавился либо стандартный oleSize, либо заданный пользователем
			const oleSize = oApi.wb.getOleSize();
			oleSize.addPointToLocalHistory();

			oApi.wb.onFrameEditorReady();
			oApi.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Open);
			delete oApi.fAfterLoad;
		}
	};
	COleCellFrameManager.prototype.openWorkbookData = function (sStream, oInfo)
	{
		this.setAfterLoadCallback();
		CCellFrameManager.prototype.openWorkbookData.call(this, sStream, oInfo);
	}

	COleCellFrameManager.prototype.isOleEditor = function ()
	{
		return true;
	};


	function CDiagramCellFrameManager(oApi)
	{
		CCellFrameManager.call(this, oApi);
		this.arrAfterLoadCallbacks = [];
		this.mainDiagram = null;
	}
	InitClassWithoutType(CDiagramCellFrameManager, CCellFrameManager);

	CDiagramCellFrameManager.prototype.clear = function ()
	{
		this.arrAfterLoadCallbacks = [];
		this.mainDiagram = null;
	};
	CDiagramCellFrameManager.prototype.isDiagramEditor = function ()
	{
		return true;
	};
	CDiagramCellFrameManager.prototype.updateProtectChart = function ()
	{
		let oAscLink;
		if (this.mainDiagram.externalReference)
		{
			oAscLink = this.mainDiagram.externalReference.getAscLink();
		}
		if (oAscLink && this.api.canEdit())
		{
			this.api.asc_addRestriction(Asc.c_oAscRestrictionType.View);
		}
		else if (!oAscLink && !this.api.canEdit())
		{
			this.api.asc_removeRestriction(Asc.c_oAscRestrictionType.View);
		}
		this.api.sendEvent('asc_onShowProtectedChartPopup', oAscLink);
	};
	CDiagramCellFrameManager.prototype.repairDiagramXLSX = function ()
	{
		const oThis = this;
		AscFormat.ExecuteNoHistory(function ()
		{
			AscCommonExcel.executeInR1C1Mode(false,
				function ()
				{
					oThis.fillWorkbookFromDiagramCache();
				});
		}, this, []);
	}
	CDiagramCellFrameManager.prototype.fillWorkbookFromDiagramCache = function ()
	{
		const mapWorksheets = this.mainDiagram.getWorksheetsFromCache(this.api.wbModel, true);
		const wsView = this.api.wb.getWorksheet();
		const oSheetData = mapWorksheets[wsView.model.sName];
		if (oSheetData) {
			wsView._updateRange(new Asc.Range(oSheetData.minC, oSheetData.minR, oSheetData.maxC , oSheetData.maxR));
		}
		wsView.draw();
	}
	CDiagramCellFrameManager.prototype.setMainDiagram = function (oInfo)
	{
			const asc_chart_binary = new Asc.asc_CChartBinary();
			asc_chart_binary.asc_setBinary(oInfo["binary"]);
			asc_chart_binary.asc_setIsChartEx(oInfo["IsChartEx"]);
			const oModel = this.api.wb.getWorksheet().model;
			const oNewChartSpace = asc_chart_binary.getChartSpace(oModel);
			oNewChartSpace.setWorksheet(oModel);

			oNewChartSpace.convertToFrameChart();
			this.mainDiagram = oNewChartSpace;
	}
	CDiagramCellFrameManager.prototype.preObtain = function (oInfo)
	{
		this.clear();
		this.setAfterLoadCallback(oInfo);
		const obtainInfo = {"documentImageUrls": oInfo["documentImageUrls"], "documentUrl": oInfo["documentUrl"], "blobUrl2Data": oInfo["blobUrl2Data"], "url2BlobUrl": oInfo["url2BlobUrl"]};
		if (oInfo["workbookBinary"])
		{
			obtainInfo["binary"] = oInfo["workbookBinary"];
			this.obtain(obtainInfo);
		}
		else
		{
			obtainInfo["binary"] = AscCommon.getEmpty();
			this.obtainWithRepair(obtainInfo);
		}
	}
	CDiagramCellFrameManager.prototype.obtainWithRepair = function (oInfo)
	{
		this.setRepairAfterLoadCallback();
		this.obtain(oInfo);
	}
	CDiagramCellFrameManager.prototype.setRepairAfterLoadCallback = function ()
	{
		this.arrAfterLoadCallbacks.push(this.repairDiagramXLSX.bind(this));
	}
	CDiagramCellFrameManager.prototype.selectMainDiagram = function ()
	{
		const oWb = this.api.wb;
		const oController = this.api.getGraphicController();
		const oWs = oWb.getWorksheet();
		if (oWs)
		{
			oController.resetSelection();
			oWs.objectRender.selectDrawingObjectRange(this.mainDiagram)
			oController.selectObject(this.mainDiagram, this.api.asc_getActiveWorksheetIndex());
			oWs.setSelectionShape(true);
			oWs.draw();
		}
	};
	CDiagramCellFrameManager.prototype.setAfterLoadCallback = function (oInfo)
	{
		const oApi = this.api;
		const oThis = this;
		oApi.fAfterLoad = function ()
		{
			oThis.setMainDiagram(oInfo);
			for (let i = 0; i < oThis.arrAfterLoadCallbacks.length; i += 1)
			{
				oThis.arrAfterLoadCallbacks[i]();
			}
			oThis.isCacheEqual = oThis.mainDiagram.isEqualCacheAndWorkbookData();
			if (!oThis.isCacheEqual)
			{
				oThis.mainDiagram.recalculateReferences();
				oThis.sendUpdateDiagram();
			}
			oThis.selectMainDiagram();
			oThis.api.wb.onFrameEditorReady();
			oThis.updateProtectChart();
			oApi.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Open);
			delete oApi.fAfterLoad;
		}
	}

	CDiagramCellFrameManager.prototype.getBinary = function ()
	{
		const oThis = this;
		return new Promise(function (resolve)
		{
			const noHistory = !AscCommon.History.Can_Undo() && oThis.isCacheEqual;
			if (noHistory)
			{
				const oRet = {};
				oRet['noHistory'] = true;
				resolve(oRet);
			}
			else
			{
				const oDiagramBinary = new Asc.asc_CChartBinary(oThis.mainDiagram);
				oThis.getWorkbookBinary().then(function (workbookBinary)
				{
					oDiagramBinary["workbookBinary"] = workbookBinary;
					oDiagramBinary["imagesForAddToHistory"] = oThis.getImagesForHistory();
					resolve(oDiagramBinary);
				});
			}
		});
	}

	CDiagramCellFrameManager.prototype.updateGeneralDiagramCache = function (aRanges)
	{
		if (!this.mainDiagram) {
			return;
		}
		const aRefsToChange = [];
		this.mainDiagram.collectIntersectionRefs(aRanges, aRefsToChange);
		for (let i = 0; i < aRefsToChange.length; i += 1)
		{
			aRefsToChange[i].updateCacheAndCat();
		}
		if (aRefsToChange.length)
		{
			this.sendUpdateDiagram();
		}
	};
	CDiagramCellFrameManager.prototype.sendUpdateDiagram = function ()
	{
		this.sendFromFrameToGeneralEditor(new CFrameUpdateDiagramData(this.mainDiagram.chart));
	};

	CDiagramCellFrameManager.prototype.getAscSettings = function ()
	{
		if (!this.mainDiagram) {
			return;
		}
		const oProps = this.mainDiagram.getAscSettings();
		const oThis = this;

		oProps.setFUpdateGeneralChart(function (bSelectFrameChartRange)
		{
			oThis.sendUpdateDiagram();
			if (bSelectFrameChartRange)
			{
				oThis.selectMainDiagram();
			}
		});
		return oProps;
	};

	function CFrameUpdateDiagramData(oDiagram)
	{
		const oBinary = new Asc.asc_CChartBinary(oDiagram);
		const oData = {"binary": oBinary};
		CFrameData.call(this, AscCommon.c_oAscFrameDataType.UpdateDiagramInGeneral, oData);
	}


	function CFrameData(type, information)
	{
		this["information"] = information;
		this["type"] = type;
	}

	function CFrameImageData(arrImages, token, bNotShowError)
	{
		const oData = {
			"images"       : arrImages,
			"token"        : token,
			"bNotShowError": bNotShowError
		}
		CFrameData.call(this, AscCommon.c_oAscFrameDataType.SendImageUrls, oData);
	}

	function CBinaryLoaderData() {
		this.stream = null;
		this.blobs = null;
		this.closeCallbacks = [];
		this.isLoadedStream = true;
		this.documentUrl = null;
	};
	CBinaryLoaderData.prototype.addCloseCallback = function (fCallback) {
		if (fCallback) {
			this.closeCallbacks.push(fCallback);
		}
	};

	function CFrameBinaryLoader()
	{
		this.api = Asc.editor || editor;
		this.binary = null;
		this.frameEditorType = AscCommon.FrameEditorTypes.JustBlock;
		const isLocalDesktop = window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]();
		this.isOpenOnClient = this.api["asc_isSupportFeature"]("ooxml") && !isLocalDesktop;
		this.eventOnOpenFrame = "";

		this.XLSXBase64 = null;
		this.blobs = null;
		this.documentUrl = null;
	}
	CFrameBinaryLoader.prototype.fillDataFromXLSTZip = function(zipData, data) {
		let jsZlib = new AscCommon.ZLib();
		if (jsZlib.open(zipData))
		{
			if (jsZlib.files && jsZlib.files.length) {
				const oImgBlobs = this.getBlobCache(jsZlib);
				const arrStream = jsZlib.getFile(jsZlib.files[0]);
				data.stream = arrStream ? Array.from(arrStream) : null;
				data.blobs = oImgBlobs;
				data.addCloseCallback(this.getCloseCallbackFromBlobCache(oImgBlobs));
			}
			jsZlib.close();
		}
	};
	CFrameBinaryLoader.prototype.getNestedPromise = function (binaryId)
	{
		const oThis = this;
		let oData = new CBinaryLoaderData();
		const binaryData = AscCommon.g_oBinaryCacheManager.getBinary(binaryId);
		if (binaryId && binaryData === null) {
			oData.isLoadedStream = false;
		}
		return new Promise(function (resolve)
		{
			if (binaryData && binaryData.length)
			{
				if (oThis.isOpenOnClient)
				{
					oData.stream = CFrameManagerBase.prototype.getEncodedArray.call(oThis, binaryData);
					resolve(oData);
				}
				else
				{
					if (AscCommon.checkOOXMLSignature(binaryData))
					{
						oThis.api.getConvertedXLSXFileFromUrl({data: binaryData}, Asc.c_oAscFileType.XLSY, function (arrBinaryData) {
							if (arrBinaryData)
							{
								oThis.fillDataFromXLSTZip(arrBinaryData, oData);
							}
							resolve(oData);
						});
					}
					else
					{
						oData.stream = CFrameManagerBase.prototype.getEncodedArray.call(oThis, binaryData);
						resolve(oData);
					}
				}
			}
			else
			{
				resolve(oData);
			}
		});
	};

	CFrameBinaryLoader.prototype.resolvePromise = function (oData) {
		if (oData.isLoadedStream && this.isTruthStream(oData.stream)) {
			this.setFromData(oData);
			this.loadFrame();
		} else {
			for (let i = 0; i < oData.closeCallbacks.length; i += 1) {
				oData.closeCallbacks[i]();
			}
		}
		this.endLoadWorksheet();
	};

	CFrameBinaryLoader.prototype.isTruthStream = function (arrStream)
	{
		return (arrStream && (arrStream.length !== 0)) || !this.isExternal();
	}
	CFrameBinaryLoader.prototype.isExternal = function() {
		return false;
	};
	CFrameBinaryLoader.prototype.startLoadWorksheet = function ()
	{
		this.api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
	};
	CFrameBinaryLoader.prototype.endLoadWorksheet = function ()
	{
		this.api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
	};

	CFrameBinaryLoader.prototype.tryOpen = function ()
	{
		this.startLoadWorksheet();
		this.resolve();
	};

	CFrameBinaryLoader.prototype.getFrameBinary = function ()
	{
	};

	CFrameBinaryLoader.prototype.loadFrame = function ()
	{
		this.api.asc_onOpenFrameEditor(this.frameEditorType);
		if(!window['IS_NATIVE_EDITOR'] && this.api.WordControl)
		{
			this.api.WordControl.onMouseUpMainSimple();
		}
		this.api.frameManager.startLoadChartEditor();
		this.api.sendEvent(this.eventOnOpenFrame, this.getFrameBinary());
	};

	CFrameBinaryLoader.prototype.setXLSX = function (sStream)
	{
		if (sStream && sStream.length)
		{
			if (typeof sStream === 'string')
			{
				this.XLSXBase64 = sStream;
			}
			else
			{
				this.XLSXBase64 = Array.from(sStream);
			}
		}
		else
		{
			this.XLSXBase64 = null;
		}
	};
	CFrameBinaryLoader.prototype.setFromData = function (oData) {
		this.setXLSX(oData.stream);
		this.setBlobs(oData.blobs);
		this.setDocumentUrl(oData.documentUrl);
		this.setCloseCallbacks(oData.closeCallbacks);
	}
	CFrameBinaryLoader.prototype.setBlobs = function(oBlobs) {
		this.blobs = oBlobs;
	};
	CFrameBinaryLoader.prototype.setDocumentUrl = function(sDocumentUrl) {
		this.documentUrl = sDocumentUrl;
	};
	CFrameBinaryLoader.prototype.setCloseCallbacks = function (arrCallbacks) {
		for (let i = 0; i < arrCallbacks.length; i += 1) {
			this.api.frameManager.addCloseCallback(arrCallbacks[i]);
		}
	};
	CFrameBinaryLoader.prototype.getBlobCache = function(openZip) {
		const oRes = {"url2BlobUrl": {}, "blobUrl2Data":{}};
		for (let i = 0; i < openZip.files.length; i += 1) {
			const sBlobUrl = AscCommon.g_oDocumentBlobUrls.getBlobUrl(openZip.files[i], openZip, oRes);
			if (oRes["blobUrl2Data"][sBlobUrl]) {
				if (oRes["blobUrl2Data"][sBlobUrl]["data"]) {
					oRes["blobUrl2Data"][sBlobUrl]["data"] = Array.from(oRes["blobUrl2Data"][sBlobUrl]["data"]);
				} else if (oRes["blobUrl2Data"][sBlobUrl]["dataBlob"]) {
					oRes["blobUrl2Data"][sBlobUrl]["dataBlob"] = Array.from(oRes["blobUrl2Data"][sBlobUrl]["dataBlob"]);
				}
			}
		}
		return oRes;
	};
	CFrameBinaryLoader.prototype.getCloseCallbackFromBlobCache = function (oCache) {
		return function () {
			if (window.URL.revokeObjectURL) {
				const oOldUrls = oCache["url2BlobUrl"];
				for (let sPath in oOldUrls) {
					const sUrl = oOldUrls[sPath];
					window.URL.revokeObjectURL(sUrl);
				}
			}
		}
	};

	function CFrameOleBinaryLoader(ole, fCallback)
	{
		CFrameBinaryLoader.call(this);
		this.oleObject = ole;
		this.frameEditorType = AscCommon.FrameEditorTypes.OleEditor;
		this.fCallback = fCallback || this.resolvePromise.bind(this);
		this.eventOnOpenFrame = 'asc_doubleClickOnTableOleObject';
	}
	InitClassWithoutType(CFrameOleBinaryLoader, CFrameBinaryLoader);
	CFrameOleBinaryLoader.prototype.resolve = function ()
	{
		const oPromise = this.getNestedPromise(this.oleObject.m_sBinaryId);
		oPromise.then(this.fCallback);
	};
	CFrameOleBinaryLoader.prototype.getFrameBinary = function ()
	{
			const oApi = Asc.editor || editor;
			if (!oApi.isOpenedFrameEditor) {
				oApi.frameManager.startLoadOleEditor();
				oApi.asc_onOpenFrameEditor(this.frameEditorType);
				const oController = oApi.getGraphicController();
				if (oController) {
					AscFormat.ExecuteNoHistory(function () {
						oController.checkSelectedObjectsAndCallback(function () {}, [], false);
					}, this, []);
				}
			}
			const sData = this.XLSXBase64;
			const nImageWidth = this.oleObject.extX * AscCommon.g_dKoef_mm_to_pix;
			const nImageHeight = this.oleObject.extY * AscCommon.g_dKoef_mm_to_pix;
			const documentImageUrls = AscCommon.g_oDocumentUrls.urls;

			return {
				"binary": sData,
				"isFromSheetEditor": !!this.oleObject.worksheet,
				"imageWidth": nImageWidth,
				"imageHeight": nImageHeight,
				"documentImageUrls": documentImageUrls,
				"documentUrl": AscCommon.g_oDocumentUrls.documentUrl
			};
	};



	function CFrameDiagramBinaryLoader(oChart, fCallback)
	{
		CFrameBinaryLoader.call(this);
		this.chart = oChart;
		this.frameEditorType = AscCommon.FrameEditorTypes.ChartEditor;
		this.fCallback = fCallback || this.resolvePromise.bind(this);
		this.eventOnOpenFrame = 'asc_doubleClickOnChart';
	}
	InitClassWithoutType(CFrameDiagramBinaryLoader, CFrameBinaryLoader);

	CFrameDiagramBinaryLoader.prototype.getFrameBinary = function ()
	{
		const oBinaryChart = new Asc.asc_CChartBinary(this.chart);
		oBinaryChart.setWorkbookBinary(this.XLSXBase64);
		oBinaryChart.setBlobsInfo(this.blobs);
		if (this.documentUrl) {
			oBinaryChart.setDocumentUrl(this.documentUrl);
		}
		return oBinaryChart;
	};

	CFrameDiagramBinaryLoader.prototype.resolveFromArray = function (arrValues)
	{
		const isLocalDesktop = window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]();
		let arrStream;
		const oData = new CBinaryLoaderData();
		if (arrValues) {
			if (arrValues.length === 1) {
				const oExternalData = arrValues[0];
				arrStream = oExternalData.stream;
				oData.documentUrl = oExternalData.documentUrl;
				oData.addCloseCallback(oExternalData.closeCallback);
				if (!this.isOpenOnClient && !isLocalDesktop)
				{
					//xlst
					this.fillDataFromXLSTZip(arrStream, oData);
				} else if (arrStream) {
					oData.stream = Array.from(arrStream);
				}
			} else {
				for (let i = 0; i < arrValues.length; i += 1) {
					oData.addCloseCallback(arrValues[i].closeCallback);
				}
			}
		}
		this.fCallback(oData);
	}

	CFrameDiagramBinaryLoader.prototype.isExternal = function ()
	{
		return this.chart.isExternal();
	}
	CFrameDiagramBinaryLoader.prototype.startLoadWorksheet = function ()
	{
		this.api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
	};
	CFrameDiagramBinaryLoader.prototype.endLoadWorksheet = function ()
	{
		this.api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
	};

	CFrameDiagramBinaryLoader.prototype.loadExternal = function ()
	{
		const oExternalDataChartManager = new AscCommon.CExternalDataLoader([this.getAscLink()], this.api, this.resolveFromArray.bind(this));
		oExternalDataChartManager.props = {forceUpdate: true};
		oExternalDataChartManager.updateExternalData();
	}
	CFrameDiagramBinaryLoader.prototype.resolve = function ()
	{
		if (this.isExternal())
		{
			this.loadExternal();
		}
		else
		{
			const oPromise = this.getNestedPromise(this.chart.XLSXId);
			oPromise.then(this.fCallback);
		}
	}

	CFrameDiagramBinaryLoader.prototype.getAscLink = function ()
	{
		const oExternalReference = this.chart.getExternalReference();
		if (oExternalReference)
		{
			return oExternalReference.getAscLink();
		}
	}

	function CFrameUpdateIsOpenOnClient(bIsOpenOnClient) {
		CFrameData.call(this, AscCommon.c_oAscFrameDataType.UpdateIsOpenOnClient, {"isOpenOnClient": bIsOpenOnClient});
	}
	function CFrameOpenLocalDesktopFileLink(sLocalFileLink) {
		CFrameData.call(this, AscCommon.c_oAscFrameDataType.OpenLocalDesktopFileLink, {"localFileLink": sLocalFileLink});
	}

	const CLoadBinaryData_Loading = 1;
	const CLoadBinaryData_Complete = 2;
	const CLoadBinaryData_Error = 3;
	function CLoadBinaryData(hash, binary) {
		this.hash = hash;
		this.binary = binary;
		this.loadState = CLoadBinaryData_Loading;
		this.loadEventHandlers = [];
		this.data = null;
	}
	CLoadBinaryData.prototype.addLoadEventHandler = function(fResolve, fReject) {
		if (this.loadState === CLoadBinaryData_Complete) {
			fResolve(this);
		} else if (this.loadState === CLoadBinaryData_Error) {
			fReject();
		} else {
			this.loadEventHandlers.push({resolve: fResolve, reject: fReject});
		}
	};
	CLoadBinaryData.prototype.setLoadState = function(pr) {
		this.loadState = pr;
		this.onChangeLoadState();
	};
	CLoadBinaryData.prototype.onChangeLoadState = function() {
		if (this.loadState === CLoadBinaryData_Complete) {
			while (this.loadEventHandlers.length) {
				const resolvers = this.loadEventHandlers.pop();
				resolvers.resolve(this);
			}
		} else if (this.loadState === CLoadBinaryData_Error) {
			while (this.loadEventHandlers.length) {
				const resolvers = this.loadEventHandlers.pop();
				resolvers.reject();
			}
		}
	};
	function CBinaryCacheManager(api) {
		this.cache = {};
		this.collaborativeCache = {};
		this.api = api;
	}
	CBinaryCacheManager.prototype.loadBinaries = function(urls) {
		const oThis = this;
		return new Promise(function(resolve) {
			let len = urls.length;
			let res = Array(len);
			if (len === 0) {
				resolve(res);
			}
			for (let i = 0; i < urls.length; i++) {
				const sFullUrl = AscCommon.getFullImageSrc2(urls[i]);
				oThis.api._downloadOriginalFile(sFullUrl, undefined, undefined, undefined, function(binary) {
					res[i] = null;
					if (binary) {
						const hash = oThis.getHash(binary);
						const loadedData = new CLoadBinaryData(hash, binary);
						loadedData.setLoadState(CLoadBinaryData_Complete);
						loadedData.data = {path: urls[i], url: sFullUrl};
						oThis.collaborativeCache[hash] = loadedData;
						res[i] = loadedData;
						oThis.cache[hash] = binary;
					}
					len -= 1;
					if (len === 0) {
						resolve(res);
					}
				});
			}
		});


	};
	CBinaryCacheManager.prototype.getHash = function(binary) {
		const sha256 = AscCommon.Digest.sha256(binary, 0, binary.length);
		return AscCommon.Hex.encode(sha256);
	};
	CBinaryCacheManager.prototype.getBinary = function(hash) {
		if (this.cache[hash]) {
			return this.cache[hash];
		}
		return null;
	};
	CBinaryCacheManager.prototype.loadBinaryToServer = function(xlsxBinary) {
		const oThis = this;
		const hash = this.getHash(xlsxBinary);
		return new Promise(function(resolve, reject) {
			if (AscCommon.History.IsOn()) {
				if (!oThis.collaborativeCache[hash]) {
					oThis.collaborativeCache[hash] = new CLoadBinaryData(hash, xlsxBinary);
					const dataUrl = oThis.getDataURLFromBinary(xlsxBinary);
					AscCommon.sendImgUrls(oThis.api, [dataUrl], function(data) {
						if (data && data[0] && data[0].url !== "error") {
							oThis.collaborativeCache[hash].setLoadState(CLoadBinaryData_Complete);
							oThis.collaborativeCache[hash].data = data[0];
						} else {
							oThis.collaborativeCache[hash].setLoadState(CLoadBinaryData_Error);
							delete oThis.collaborativeCache[hash];
						}
					});
				}
				const oLoadedData = oThis.collaborativeCache[hash];
				if (oLoadedData) {
					oLoadedData.addLoadEventHandler(resolve, reject);
				} else {
					reject();
				}
			} else {
				resolve(new CLoadBinaryData(hash, xlsxBinary));
			}
		});

	};
	CBinaryCacheManager.prototype.addBinary = function(binary) {
		binary = binary instanceof Uint8Array ? binary : new Uint8Array(binary);
		const oThis = this;
		return this.getFormatBinary(binary).then(function(arrXLSXBinary) {
			return oThis.loadBinaryToServer(arrXLSXBinary);
		}).then(function(oLoadedData) {
			if (oLoadedData) {
				oThis.cache[oLoadedData.hash] = oLoadedData.binary;
				return oLoadedData;
			}
			return null;
		});
	};
	CBinaryCacheManager.prototype.addLocalBinary = function(binary) {
		const hash = this.getHash(binary);
		this.cache[hash] = binary;
		return hash;
	};
	CBinaryCacheManager.prototype.getBase64EncodedData = function(binary) {
		return ("XLSY;v2;" + binary.length + ";" + AscCommon.Base64.encode(binary)).toUtf8();
	}
	CBinaryCacheManager.prototype.getFormatBinary = function(binary) {
		const oThis = this;
		return new Promise(function(resolve) {
			if (AscCommon.checkOOXMLSignature(binary)) {
				resolve(binary);
			} else {
				const xlsyBinary = oThis.getBase64EncodedData(binary);
				oThis.api.getConvertedXLSXFileFromUrl({data: xlsyBinary}, Asc.c_oAscFileType.XLSX, function (arrBinaryData) {
					if (arrBinaryData) {
						resolve(arrBinaryData);
					} else {
						resolve(null);
					}
				});
			}
		});
	};
	CBinaryCacheManager.prototype.getDataURLFromBinary = function(arrXLSXBinary) {
		return "data:" + AscCommon.openXml.Types.package.contentType + ";base64," + AscCommon.Base64.encode(arrXLSXBinary, 0, arrXLSXBinary.length);
	};

	window["AscCommon"].CBinaryCacheManager = CBinaryCacheManager;
	window["AscCommon"].CDiagramCellFrameManager = CDiagramCellFrameManager;
	window["AscCommon"].CMainEditorFrameManager  = CMainEditorFrameManager;
	window["AscCommon"].COleCellFrameManager = COleCellFrameManager;
	window["AscCommon"].CFrameDiagramBinaryLoader = CFrameDiagramBinaryLoader;
	window["AscCommon"].CFrameOleBinaryLoader = CFrameOleBinaryLoader;
	window["AscCommon"].CFrameUpdateIsOpenOnClient = CFrameUpdateIsOpenOnClient;
	window["AscCommon"].FrameEditorTypes = FrameEditorTypes;
})(window);
