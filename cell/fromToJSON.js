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
(function(window, undefined){

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

	function private_EMU2MM(EMU)
	{
		return EMU / 36000.0;
	}
	function private_MM2EMU(MM)
	{
		return MM * 36000.0;
	}

	var DocumentPageSize = new function() {
        this.oSizes = [
            {id:Asc.EPageSize.pagesizeLetterPaper, w_mm: 215.9, h_mm: 279.4},
            {id:Asc.EPageSize.pagesizeLetterSmall, w_mm: 215.9, h_mm: 279.4},
            {id:Asc.EPageSize.pagesizeTabloidPaper, w_mm: 279.4, h_mm: 431.8},
            {id:Asc.EPageSize.pagesizeLedgerPaper, w_mm: 431.8, h_mm: 279.4},
            {id:Asc.EPageSize.pagesizeLegalPaper, w_mm: 215.9, h_mm: 355.6},
            {id:Asc.EPageSize.pagesizeStatementPaper, w_mm: 495.3, h_mm: 215.9},
            {id:Asc.EPageSize.pagesizeExecutivePaper, w_mm: 184.2, h_mm: 266.7},
            {id:Asc.EPageSize.pagesizeA3Paper, w_mm: 297, h_mm: 420},
            {id:Asc.EPageSize.pagesizeA4Paper, w_mm: 210, h_mm: 297},
            {id:Asc.EPageSize.pagesizeA4SmallPaper, w_mm: 210, h_mm: 297},
            {id:Asc.EPageSize.pagesizeA5Paper, w_mm: 148, h_mm: 210},
            {id:Asc.EPageSize.pagesizeB4Paper, w_mm: 250, h_mm: 353},
            {id:Asc.EPageSize.pagesizeB5Paper, w_mm: 176, h_mm: 250},
            {id:Asc.EPageSize.pagesizeFolioPaper, w_mm: 215.9, h_mm: 330.2},
            {id:Asc.EPageSize.pagesizeQuartoPaper, w_mm: 215, h_mm: 275},
            {id:Asc.EPageSize.pagesizeStandardPaper1, w_mm: 254, h_mm: 355.6},
            {id:Asc.EPageSize.pagesizeStandardPaper2, w_mm: 279.4, h_mm: 431.8},
            {id:Asc.EPageSize.pagesizeNotePaper, w_mm: 215.9, h_mm: 279.4},
            {id:Asc.EPageSize.pagesize9Envelope, w_mm: 98.4, h_mm: 225.4},
            {id:Asc.EPageSize.pagesize10Envelope, w_mm: 104.8, h_mm: 241.3},
            {id:Asc.EPageSize.pagesize11Envelope, w_mm: 114.3, h_mm: 263.5},
            {id:Asc.EPageSize.pagesize12Envelope, w_mm: 120.7, h_mm: 279.4},
            {id:Asc.EPageSize.pagesize14Envelope, w_mm: 127, h_mm: 292.1},
            {id:Asc.EPageSize.pagesizeCPaper, w_mm: 431.8, h_mm: 558.8},
            {id:Asc.EPageSize.pagesizeDPaper, w_mm: 558.8, h_mm: 863.6},
            {id:Asc.EPageSize.pagesizeEPaper, w_mm: 863.6, h_mm: 1117.6},
            {id:Asc.EPageSize.pagesizeDLEnvelope, w_mm: 110, h_mm: 220},
            {id:Asc.EPageSize.pagesizeC5Envelope, w_mm: 162, h_mm: 229},
            {id:Asc.EPageSize.pagesizeC3Envelope, w_mm: 324, h_mm: 458},
            {id:Asc.EPageSize.pagesizeC4Envelope, w_mm: 229, h_mm: 324},
            {id:Asc.EPageSize.pagesizeC6Envelope, w_mm: 114, h_mm: 162},
            {id:Asc.EPageSize.pagesizeC65Envelope, w_mm: 114, h_mm: 229},
            {id:Asc.EPageSize.pagesizeB4Envelope, w_mm: 250, h_mm: 353},
            {id:Asc.EPageSize.pagesizeB5Envelope, w_mm: 176, h_mm: 250},
            {id:Asc.EPageSize.pagesizeB6Envelope, w_mm: 176, h_mm: 125},
            {id:Asc.EPageSize.pagesizeItalyEnvelope, w_mm: 110, h_mm: 230},
            {id:Asc.EPageSize.pagesizeMonarchEnvelope, w_mm: 98.4, h_mm: 190.5},
            {id:Asc.EPageSize.pagesize6_3_4Envelope, w_mm: 92.1, h_mm: 165.1},
            {id:Asc.EPageSize.pagesizeUSStandardFanfold, w_mm: 377.8, h_mm: 279.4},
            {id:Asc.EPageSize.pagesizeGermanStandardFanfold, w_mm: 215.9, h_mm: 304.8},
            {id:Asc.EPageSize.pagesizeGermanLegalFanfold, w_mm: 215.9, h_mm: 330.2},
            {id:Asc.EPageSize.pagesizeISOB4, w_mm: 250, h_mm: 353},
            {id:Asc.EPageSize.pagesizeJapaneseDoublePostcard, w_mm: 200, h_mm: 148},
            {id:Asc.EPageSize.pagesizeStandardPaper3, w_mm: 228.6, h_mm: 279.4},
            {id:Asc.EPageSize.pagesizeStandardPaper4, w_mm: 254, h_mm: 279.4},
            {id:Asc.EPageSize.pagesizeStandardPaper5, w_mm: 381, h_mm: 279.4},
            {id:Asc.EPageSize.pagesizeInviteEnvelope, w_mm: 220, h_mm: 220},
            {id:Asc.EPageSize.pagesizeLetterExtraPaper, w_mm: 235.6, h_mm: 304.8},
            {id:Asc.EPageSize.pagesizeLegalExtraPaper, w_mm: 235.6, h_mm: 381},
            {id:Asc.EPageSize.pagesizeTabloidExtraPaper, w_mm: 296.9, h_mm: 457.2},
            {id:Asc.EPageSize.pagesizeA4ExtraPaper, w_mm: 236, h_mm: 322},
            {id:Asc.EPageSize.pagesizeLetterTransversePaper, w_mm: 210.2, h_mm: 279.4},
            {id:Asc.EPageSize.pagesizeA4TransversePaper, w_mm: 210, h_mm: 297},
            {id:Asc.EPageSize.pagesizeLetterExtraTransversePaper, w_mm: 235.6, h_mm: 304.8},
            {id:Asc.EPageSize.pagesizeSuperA_SuperA_A4Paper, w_mm: 227, h_mm: 356},
            {id:Asc.EPageSize.pagesizeSuperB_SuperB_A3Paper, w_mm: 305, h_mm: 487},
            {id:Asc.EPageSize.pagesizeLetterPlusPaper, w_mm: 215.9, h_mm: 12.69},
            {id:Asc.EPageSize.pagesizeA4PlusPaper, w_mm: 210, h_mm: 330},
            {id:Asc.EPageSize.pagesizeA5TransversePaper, w_mm: 148, h_mm: 210},
            {id:Asc.EPageSize.pagesizeJISB5TransversePaper, w_mm: 182, h_mm: 257},
            {id:Asc.EPageSize.pagesizeA3ExtraPaper, w_mm: 322, h_mm: 445},
            {id:Asc.EPageSize.pagesizeA5ExtraPaper, w_mm: 174, h_mm: 235},
            {id:Asc.EPageSize.pagesizeISOB5ExtraPaper, w_mm: 201, h_mm: 276},
            {id:Asc.EPageSize.pagesizeA2Paper, w_mm: 420, h_mm: 594},
            {id:Asc.EPageSize.pagesizeA3TransversePaper, w_mm: 297, h_mm: 420},
            {id:Asc.EPageSize.pagesizeA3ExtraTransversePaper, w_mm: 322, h_mm: 445},
            {id:Asc.EPageSize.pagesizeEnvelopeChoukei3, w_mm: 120, h_mm: 235},
            {id:Asc.EPageSize.pagesizeROC16K, w_mm: 196.8, h_mm: 273}
        ];
        this.getSizeByWH = function(widthMm, heightMm)
        {
            for( var index in this.oSizes)
            {
                var item = this.oSizes[index];
                if(widthMm == item.w_mm && heightMm == item.h_mm)
                    return item;
            }
            return this.oSizes[8];//A4
        };
        this.getSizeById = function(id)
        {
            for( var index in this.oSizes)
            {
                var item = this.oSizes[index];
                if(id == item.id)
                    return item;
            }
            return this.oSizes[8];//A4
        };
    };
	
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// End of private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    var WriterToJSON   = window['AscJsonConverter'].WriterToJSON;
	var ReaderFromJSON = window['AscJsonConverter'].ReaderFromJSON;
	
	/**
	 * Converts the specified worksheets objects into the JSON.
	 * @typeofeditors ["CSE"]
	 * @param {Worksheet} oWorksheet
	 * @return {object} 
	 */
	WriterToJSON.prototype.SerWorksheet = function(oWorksheet)
	{
		var aCols = this.SerCols(oWorksheet);

		var aDrawings = [];
		for (var nDrawing = 0; nDrawing < oWorksheet.Drawings.length; nDrawing++)
			aDrawings.push(this.SerDrawingExcell(oWorksheet.Drawings[nDrawing]));

		var aHyperlinks = [];
		var aWorksheetLinks = oWorksheet.hyperlinkManager.getAll();

		for (var nHyperlink = 0; nHyperlink < aWorksheetLinks.length; nHyperlink++)
			aHyperlinks.push(this.SerHyperlinkExcel(aWorksheetLinks[nHyperlink]));

		var aMergeCells = this.SerMergeCells(oWorksheet.mergeManager.getAll());
		return {
			// worksheet props
			"name":                  oWorksheet.sName,
			"id":                    oWorksheet.Id,
			"hidden":                oWorksheet.bHidden ? "hidden" : "visible",

			// members
			"cols":                  aCols.length > 0 ? aCols : undefined,
			"sheetFormatPr":         oWorksheet.oSheetFormatPr != null ? this.SerSheetFormatPr(oWorksheet.oSheetFormatPr) : undefined,
			"pageMargins":           oWorksheet.PagePrintOptions != null ? this.SerPageMargins(oWorksheet.PagePrintOptions.pageMargins) : undefined,
			"pageSetup":             oWorksheet.PagePrintOptions != null ? this.SerPageSetupExcel(oWorksheet.PagePrintOptions.pageSetup) : undefined,
			"printOptions":          oWorksheet.PagePrintOptions != null ? this.SerPrintOptionsExcel(oWorksheet.PagePrintOptions) : undefined,
			"hiperlinks":            aHyperlinks.length > 0 ? aHyperlinks : undefined,
			"mergeCells":            aMergeCells.length > 0 ? aMergeCells : undefined,
			"sheetData":             this.SerSheetData(oWorksheet),
			"drawings":              aDrawings.length > 0 ? aDrawings : undefined,  
			"autoFilter":            oWorksheet.AutoFilter != null ? this.SerAutoFilter(oWorksheet.AutoFilter) : undefined,
			"sortState":             oWorksheet.sortState != null ? this.SerSortState(oWorksheet.sortState) : undefined,
			"tableParts":            oWorksheet.TableParts.length > 0 ? this.SerTableParts(oWorksheet.TableParts) : undefined,
			"comments":              oWorksheet.aComments.length > 0 ? this.SerComments(oWorksheet.aComments) : undefined,
			"conditionalFormatting": oWorksheet.aConditionalFormattingRules.length > 0 ? this.SerCondFormatting(oWorksheet.aConditionalFormattingRules) : undefined,
			"sheetViews":            oWorksheet.sheetViews.length > 0 ? this.SerSheetViews(oWorksheet.sheetViews, oWorksheet) : undefined,
			"sheetPr":               oWorksheet.sheetPr != null ? this.SerSheetPr(oWorksheet.sheetPr) : undefined,
			"sparklineGroup":        oWorksheet.aSparklineGroups.length > 0 ? this.SerSparklineGroups(oWorksheet.aSparklineGroups) : undefined,
			"headerFooter":          oWorksheet.headerFooter != null ? this.SerHdrFtrExcell(oWorksheet.headerFooter) : undefined, /// всегда лежит объект CHeaderFooterData
			"dataValidations":       oWorksheet.dataValidations != null ? this.SerDataValidations(oWorksheet.dataValidations) : undefined,
			"pivotTables":           oWorksheet.pivotTables.length > 0 ? this.SerPivotTables(oWorksheet.pivotTables) : undefined,
			"slicers":               oWorksheet.aSlicers.length > 0 ? this.SerSlicers(oWorksheet.aSlicers) : undefined,
			"namedSheetViews":       oWorksheet.aNamedSheetViews.length > 0 ? this.SerNamedSheetViews(oWorksheet.aNamedSheetViews) : undefined,
			"sheetProtection":       oWorksheet.sheetProtection != null ? this.SerSheetProtection(oWorksheet.sheetProtection) : undefined,
			"protectedRanges":       oWorksheet.aProtectedRanges.length > 0 ? this.SerProtectedRanges(oWorksheet.aProtectedRanges) : undefined,
			"type":                  "worksheet"
		}
	};
	WriterToJSON.prototype.SerWorksheets = function(nStart, nEnd)
	{
		let aSheets = this.api.wbModel.aWorksheets;

		if (this.Workbook == null)
			this.Workbook = this.api.wbModel;

		if (this.InitSaveManager == null)
			this.InitSaveManager = new AscCommonExcel.InitSaveManager(this.Workbook);

		// init styles for write
		if (this.stylesForWrite == null)
		{
			this.stylesForWrite = new AscCommonExcel.StylesForWrite();
			this.InitSaveManager._prepeareStyles(this.stylesForWrite);
		}
		// pivot caches
		let oPivotCaches = {};
		let nCachesCount = this.Workbook.preparePivotForSerialization(oPivotCaches);

		// slicer cache
		let oSlicerCaches = this.InitSaveManager.getSlicersCache();
		let aSlicerCachesExt = this.InitSaveManager.getSlicersCache(true);

		let aSerSheets = [];
		for (let Index = nStart; Index <= nEnd; Index++)
			aSerSheets.push(this.SerWorksheet(aSheets[Index]));

		return {
			"sheets":          aSerSheets,

			// styles
			"styles":          this.SerExcelStylesForWrite(this.stylesForWrite),
			"pivotCaches":     nCachesCount > 0 ? this.SerPivotCaches(oPivotCaches) : undefined,
			"slicerCaches":    oSlicerCaches != null ? this.SerSlicerCaches(oSlicerCaches) : undefined,
			"slicerCachesExt": aSlicerCachesExt != null ? this.SerSlicerCaches(aSlicerCachesExt) : undefined,

			"type": "sheets"
		}
	};
	WriterToJSON.prototype.SerDrawingExcell = function(oDrawingExcel)
	{
		var nTypeToWrite = oDrawingExcel.Type;
		if(oDrawingExcel.graphicObject.getObjectType() === AscDFH.historyitem_type_OleObject)
		{
			nTypeToWrite = AscCommon.c_oAscCellAnchorType.cellanchorTwoCell;
		}

		var sEditAs, oFrom, oTo, oExt, oPos, oClientData;
		switch(nTypeToWrite)
		{
			case AscCommon.c_oAscCellAnchorType.cellanchorTwoCell:
			{
				sEditAs = ToXML_ST_EditAs(oDrawingExcel.editAs);
				oFrom = this.SerFromTo(oDrawingExcel.from);
				oTo = this.SerFromTo(oDrawingExcel.to);
				break;
			}
			case AscCommon.c_oAscCellAnchorType.cellanchorOneCell:
			{
				oFrom = this.SerFromTo(oDrawingExcel.from);
				oExt = {
					"cx": private_MM2EMU(oDrawingExcel.ext.cx),
					"cy": private_MM2EMU(oDrawingExcel.ext.cy)
				}
				break;
			}
			case AscCommon.c_oAscCellAnchorType.cellanchorAbsolute:
			{
				oPos = {
					"x": oDrawingExcel.Pos.X,
					"y": oDrawingExcel.Pos.Y
				};
				oExt = {
					"cx": private_MM2EMU(oDrawingExcel.ext.cx),
					"cy": private_MM2EMU(oDrawingExcel.ext.cy)
				}
				break;
			}
		}

		if (oDrawingExcel.clientData)
			oClientData = this.SerClientData(oDrawingExcel.graphicObject.clientData);

		return {
			"clientData": oClientData,
			"pos":        oPos,
			"ext":        oExt,
			"from":       oFrom,
			"to":         oTo,
			"graphic":    this.SerGraphicObject(oDrawingExcel.graphicObject),
			"editAs":     sEditAs,
			"type":       nTypeToWrite != null ? ToXML_ST_EditAs(nTypeToWrite) : undefined
		}
	};
	WriterToJSON.prototype.SerClientData = function(oClientData)
	{
		return {
			"fLocksWithSheet":  oClientData.fLocksWithSheet != null ? oClientData.fLocksWithSheet : undefined,
			"fPrintsWithSheet": oClientData.fPrintsWithSheet != null ? oClientData.fPrintsWithSheet : undefined
		}
	};
	WriterToJSON.prototype.SerFromTo = function(oFromTo) // CCellObjectInfo
	{
		if (!oFromTo)
			return oFromTo;

		return {
			"col":    oFromTo.col,
			"colOff": private_MM2EMU(oFromTo.colOff),
			"row":    oFromTo.row,
			"rowOff": private_MM2EMU(oFromTo.rowOff)
		}
	};
	WriterToJSON.prototype.SerHyperlinkExcel = function(oHyperlink)
	{
		return {
			"ref":      oHyperlink.bbox ? this.SerRef(oHyperlink.bbox) : undefined,
			"display":  oHyperlink.data.Hyperlink != null ? oHyperlink.data.Hyperlink : undefined,
			"location": oHyperlink.data.getLocation() != null ? oHyperlink.data.getLocation() : undefined,
			"tooltip":  oHyperlink.data.Tooltip != null ? oHyperlink.data.Tooltip : undefined
		}
	};
	WriterToJSON.prototype.SerMergeCells = function(aInfo)
	{
		var aRefs = [];
		for (var nRef = 0; nRef < aInfo.length; nRef++)
			aRefs.push(this.SerRef(aInfo[nRef].bbox));

		return aRefs;
	};
	WriterToJSON.prototype.SerPageSetupExcel = function(oPageSetup)
	{
		if (!oPageSetup)
			return oPageSetup;

		var sOrientType = undefined;
		switch(oPageSetup.orientation)
		{
			case Asc.EPageOrientation.pageorientPortrait:
				sOrientType = "portrait";
				break;
			case Asc.EPageOrientation.pageorientLandscape:
				sOrientType = "landscape";
				break;
		}

		var dWidth = oPageSetup.asc_getWidth();
		var dHeight = oPageSetup.asc_getHeight();
		var nSizeId = undefined;
		if(null != dWidth && null != dHeight)
			nSizeId = DocumentPageSize.getSizeByWH(dWidth, dHeight).id;

		return {
			"blackAndWhite":      oPageSetup.blackAndWhite != null ? oPageSetup.blackAndWhite : undefined,
			"cellComments":       oPageSetup.cellComments != null ? ToXML_ST_CellComments(oPageSetup.cellComments) : undefined,
			"copies":             oPageSetup.copies != null ? oPageSetup.copies : undefined,
			"draft":              oPageSetup.draft != null ? oPageSetup.draft : undefined,
			"errors":             oPageSetup.errors != null ? ToXML_ST_PrintError(oPageSetup.errors) : undefined,
			"firstPageNumber":    oPageSetup.firstPageNumber != null ? oPageSetup.firstPageNumber : undefined,
			"fitToHeight":        oPageSetup.fitToHeight != null ? oPageSetup.fitToHeight : undefined,
			"fitToWidth":         oPageSetup.fitToWidth != null ? oPageSetup.fitToWidth : undefined,
			"horizontalDpi":      oPageSetup.horizontalDpi != null ? oPageSetup.horizontalDpi : undefined,
			"orientation":        sOrientType,
			"pageOrder":          oPageSetup.pageOrder != null ? ToXML_ST_PageOrder(oPageSetup.pageOrder) : undefined,
			"paperSize":          nSizeId,
			"scale":              oPageSetup.scale != null ? oPageSetup.scale : undefined,
			"useFirstPageNumber": oPageSetup.useFirstPageNumber != null ? oPageSetup.useFirstPageNumber : undefined,
			"usePrinterDefaults": oPageSetup.usePrinterDefaults != null ? oPageSetup.usePrinterDefaults : undefined,
			"verticalDpi":        oPageSetup.verticalDpi != null ? oPageSetup.verticalDpi : undefined,

			// not reading
			//paperHeight:        oPageSetup.paperHeight,
			//paperWidth:         oPageSetup.paperWidth
		}
	};
	WriterToJSON.prototype.SerPrintOptionsExcel = function(oOptions)
	{
		return {
			"gridLines": oOptions.gridLines != null ? oOptions.gridLines : undefined,
			"headings":  oOptions.headings != null ? oOptions.headings : undefined
		}
	};
	WriterToJSON.prototype.SerProtectedRanges = function(aRanges)
	{
		var aResult = [];
		for (var nRange = 0; nRange < aRanges.length; nRange++)
			aResult.push(this.SerProtectedRange(aRanges[nRange]));

		return aResult;
	};
	WriterToJSON.prototype.SerProtectedRange = function(oRange)
	{
		return {
			"algorithmName":      oRange.algorithmName != null ? To_XML_CryptoAlgorithmName(oRange.AlgorithmName) : undefined,
			"spinCount":          oRange.spinCount != null ? oRange.spinCount : undefined,
			"hashValue":          oRange.hashValue != null ? oRange.hashValue : undefined,
			"saltValue":          oRange.saltValue != null ? oRange.saltValue : undefined,
			"name":               oRange.name != null ? oRange.name : undefined,
			"sqref":              oRange.sqref != null ? AscCommonExcel.getSqRefString(oRange.sqref) : undefined,
			"securityDescriptor": oRange.securityDescriptor != null ? oRange.securityDescriptor : undefined
		}
	};
	WriterToJSON.prototype.SerSheetFormatPr = function(oPr)
	{
		if (!oPr)
			return oPr;

		return {
			"defaultColWidth":  oPr.dDefaultColWidth != undefined ? oPr.dDefaultColWidth : undefined,
			"baseColWidth":     oPr.nBaseColWidth != undefined ? oPr.nBaseColWidth : undefined,
			"defaultRowHeight": oPr.oAllRow && oPr.oAllRow.h != undefined ? oPr.oAllRow.h : undefined,
			"customHeight":     oPr.oAllRow && oPr.oAllRow.getCustomHeight() ? true : undefined,
			"zeroHeight":       oPr.oAllRow && oPr.oAllRow.getHidden() ? true : undefined,
			"outlineLevelCol":  oPr.nOutlineLevelCol > 0 ? oPr.nOutlineLevelCol : undefined,
			"outlineLevelRow":  oPr.oAllRow && oPr.oAllRow.getOutlineLevel() > 0 ? oPr.oAllRow.getOutlineLevel() : undefined
			//thickBottom: null // not supported
		}
	};
	WriterToJSON.prototype.SerPageMargins = function(oMargins)
	{
		if (!oMargins)
			return oMargins;

		return {
			"bottom": oMargins.bottom != null ? oMargins.bottom: undefined,
			"footer": oMargins.footer != null ? oMargins.footer: undefined,
			"header": oMargins.header != null ? oMargins.header: undefined,
			"left":   oMargins.left != null ? oMargins.left: undefined,
			"right":  oMargins.right != null ? oMargins.right: undefined,
			"top":    oMargins.top != null ? oMargins.top: undefined,
		}
	};
	WriterToJSON.prototype.SerSheetPr = function(oPr)
	{
		if (!oPr)
			return oPr;

		return {
			"codeName":                          oPr.CodeName != null ? oPr.CodeName : undefined,
			"enableFormatConditionsCalculation": oPr.EnableFormatConditionsCalculation != null ? oPr.EnableFormatConditionsCalculation : undefined,
			"filterMode":                        oPr.FilterMode != null ? oPr.FilterMode : undefined,
			"published":                         oPr.Published != null ? oPr.Published : undefined,
			"syncHorizontal":                    oPr.SyncHorizontal != null ? oPr.SyncHorizontal : undefined,
			"syncRef":                           oPr.SyncRef != null ? oPr.SyncRef : undefined,
			"syncVertical":                      oPr.SyncVertical != null ? oPr.SyncVertical : undefined,
			"transitionEntry":                   oPr.TransitionEntry != null ? oPr.TransitionEntry : undefined,
			"transitionEvaluation":              oPr.TransitionEvaluation != null ? oPr.TransitionEvaluation : undefined,

			"tabColor":           oPr.TabColor != null ? this.SerColorExcell(oPr.TabColor) : undefined,
			"pageSetUpPr": {
				"autoPageBreaks": oPr.AutoPageBreaks != null ? oPr.AutoPageBreaks : undefined,
				"fitToPage":      oPr.FitToPage != null ? oPr.FitToPage : undefined
			},
			"outlinePr": {
				"applyStyles":        oPr.ApplyStyles != null ? oPr.ApplyStyles : undefined,
				"showOutlineSymbols": oPr.ShowOutlineSymbols != null ? oPr.ShowOutlineSymbols : undefined,
				"summaryBelow":       oPr.SummaryBelow != null ? oPr.SummaryBelow : undefined,
				"summaryRight":       oPr.SummaryRight != null ? oPr.SummaryRight : undefined
			}
		}
	};
	WriterToJSON.prototype.SerSheetProtection = function(oProtection)
	{
		if (!oProtection)
			return oProtection;

		return {
			"algorithmName":        oProtection.algorithmName != null ? To_XML_CryptoAlgorithmName(oProtection.algorithmName) : undefined,
			"autoFilter":           oProtection.autoFilter != null ? oProtection.autoFilter : undefined,
			"deleteColumns":        oProtection.deleteColumns != null ? oProtection.deleteColumns : undefined,
			"deleteRows":           oProtection.deleteRows != null ? oProtection.deleteRows : undefined,
			"formatCells":          oProtection.formatCells != null ? oProtection.formatCells : undefined,
			"formatColumns":        oProtection.formatColumns != null ? oProtection.formatColumns : undefined,
			"formatRows":           oProtection.formatRows != null ? oProtection.formatRows : undefined,
			"hashValue":            oProtection.hashValue != null ? oProtection.hashValue : undefined,
			"insertColumns":        oProtection.insertColumns != null ? oProtection.insertColumns : undefined,
			"insertHyperlinks":     oProtection.insertHyperlinks != null ? oProtection.insertHyperlinks : undefined,
			"insertRows":           oProtection.insertRows != null ? oProtection.insertRows : undefined,
			"objects":              oProtection.objects != null ? oProtection.objects : undefined,
			"pivotTables":          oProtection.pivotTables != null ? oProtection.pivotTables : undefined,
			"saltValue":            oProtection.saltValue != null ? oProtection.saltValue : undefined,
			"scenarios":            oProtection.scenarios != null ? oProtection.scenarios : undefined,
			"selectLockedCells":    oProtection.selectLockedCells != null ? oProtection.selectLockedCells : undefined,
			"selectUnlockedCells":  oProtection.selectUnlockedCells != null ? oProtection.selectUnlockedCells : undefined,
			"sheet":                oProtection.sheet != null ? oProtection.sheet : undefined,
			"sort":                 oProtection.sort != null ? oProtection.sort : undefined,
			"spinCount":            oProtection.spinCount != null ? oProtection.spinCount : undefined,
			"password":             oProtection.password != null ? oProtection.password : undefined,
			"content":              oProtection.content != null ? oProtection.content : undefined
		}
	};
	WriterToJSON.prototype.SerSheetViews = function(aViews, oWorksheet)
	{
		var aResult = [];
		for(var nView = 0; nView < aViews.length; nView++)
			aResult.push(this.SerSheetView(aViews[nView], oWorksheet));

		return aResult;
	};
	WriterToJSON.prototype.SerSheetView = function(oView, oWorksheet)
	{
		return {
			"showGridLines":     oView.showGridLines != null ? oView.showGridLines : undefined,
			"showRowColHeaders": oView.showRowColHeaders != null ? oView.showRowColHeaders : undefined,
			"showZeros":         oView.showZeros != null ? oView.showZeros : undefined,
			"topLeftCell":       oView.topLeftCell != null ? oView.topLeftCell.getName() : undefined,
			"zoomScale":         oView.zoomScale != null ? oView.zoomScale : undefined,
			"pane":              oView.pane != null ? this.SerPane(oView.pane) : undefined, 
			"selection":         oWorksheet.selectionRange != null ? this.SerSelectionRange(oWorksheet.selectionRange) : undefined
		}
	};
	WriterToJSON.prototype.SerPane = function(oPane)
	{
		if (!oPane)
			return oPane;

		var nActivePane = null;
		var nCol = oPane.topLeftFrozenCell.getCol0();
		var nRow = oPane.topLeftFrozenCell.getRow0();
		if (0 < nCol && 0 < nRow) {
			nActivePane = AscCommonExcel.EActivePane.bottomRight;
		} else if (0 < nRow) {
			nActivePane = AscCommonExcel.EActivePane.bottomLeft;
		} else if (0 < nCol) {
			nActivePane = AscCommonExcel.EActivePane.topRight;
		}
		
		return {
			"activePane":  nActivePane != null ? ToXML_EActivePane(nActivePane) : undefined,
			"state":       "frozen", // Всегда пишем Frozen
			"topLeftCell": oPane.topLeftFrozenCell.getID(),
			"xSplit":      0 < nCol ? nCol : undefined,
			"ySplit":      0 < nRow ? nRow : undefined
		}
	};
	WriterToJSON.prototype.SerSelectionRange = function(oSelection)
	{
		if (!oSelection)
			return oSelection;

		return {
			"activeCell":   oSelection.activeCell != null ? this.SerRefCell(oSelection.activeCell) : undefined,
			"activeCellId": oSelection.activeCellId != null ? oSelection.activeCellId : undefined,
			"sqref":        oSelection.ranges != null ? AscCommonExcel.getSqRefString(oSelection.ranges) : undefined
		}
	};
	WriterToJSON.prototype.SerTableParts = function(aParts)
	{
		var aResult = [];
		for (var nPart = 0; nPart < aParts.length; nPart++)
			aResult.push(this.SerTablePart(aParts[nPart]));

		return aResult;
	};
	WriterToJSON.prototype.SerTablePart = function(oPart)
	{
		return {
			"ref":            oPart.Ref != null ? oPart.Ref.getName() : undefined,
			"headerRowCount": oPart.HeaderRowCount != null ? oPart.HeaderRowCount : undefined,
			"totalsRowCount": oPart.TotalsRowCount != null ? oPart.TotalsRowCount : undefined,
			"displayName":    oPart.DisplayName != null ? oPart.DisplayName : undefined,
			"autoFilter":     oPart.AutoFilter != null ? this.SerAutoFilter(oPart.AutoFilter) : undefined,
			"sortState":      oPart.SortState != null ? this.SerSortState(oPart.SortState) : undefined,
			"tableColumns":   oPart.TableColumns != null ? this.SerTableColumns(oPart.TableColumns) : undefined,
			"tableStyleInfo": oPart.TableStyleInfo != null ? this.SerTableStyleInfo(oPart.TableStyleInfo) : undefined,
			"altText":        oPart.altText != null ? oPart.altText : undefined,
			"altTextSummary": oPart.altTextSummary != null ? oPart.altTextSummary : undefined,
			"id":             oPart.id != null ? oPart.id : undefined,
			"queryTable":     oPart.QueryTable != null ? this.SerQueryTable(oPart.QueryTable) : undefined,
			"tableType":      oPart.tableType != null ? ToXML_ST_TableType(oPart.tableType) : undefined
		}
	};
	WriterToJSON.prototype.SerTableColumns = function(aColumns)
	{
		var aResult = [];
		for (var nCol = 0; nCol < aColumns.length; nCol++)
			aResult.push(this.SerTableColumn(aColumns[nCol]));

		return aResult;
	};
	WriterToJSON.prototype.SerTableColumn = function(oColumn)
	{
		if (oColumn.dxf)
			this.InitSaveManager.aDxfs.push(oColumn.dxf);

		return {
			"name":              oColumn.Name != null ? oColumn.Name : undefined,
			"totalsRowLabel":    oColumn.TotalsRowLabel != null ? oColumn.TotalsRowLabel : undefined,
			"totalsRowFunction": oColumn.TotalsRowFunction != null ? ToXML_ETotalsRowFunction(oColumn.TotalsRowFunction) : undefined,
			"totalsRowFormula":  oColumn.TotalsRowFormula != null ? oColumn.TotalsRowFormula : undefined,
			"dataDxfId":         oColumn.dxf != null ? this.InitSaveManager.aDxfs.length : undefined,
			"queryTableFieldId": oColumn.queryTableFieldId != null ? oColumn.queryTableFieldId : undefined,
			"uniqueName":        oColumn.uniqueName != null ? oColumn.uniqueName : undefined,
			"id":                oColumn.id != null ? oColumn.id : undefined
		}
	};
	WriterToJSON.prototype.SerTableStyleInfo = function(oInfo)
	{
		if (!oInfo)
			return oInfo;
		
		return {
			"name":              oInfo.Name != null ? oInfo.Name : undefined,
			"showColumnStripes": oInfo.ShowColumnStripes != null ? oInfo.ShowColumnStripes : undefined,
			"showRowStripes":    oInfo.ShowRowStripes != null ? oInfo.ShowRowStripes : undefined,
			"showFirstColumn":   oInfo.ShowFirstColumn != null ? oInfo.ShowFirstColumn : undefined,
			"showLastColumn":    oInfo.ShowLastColumn != null ? oInfo.ShowLastColumn : undefined
		}
	};
	WriterToJSON.prototype.SerQueryTable = function(oTable)
	{
		if (!oTable)
			return oTable;

		return {
			"connectionId":            oTable.connectionId != null ? oTable.connectionId : undefined,
			"name":                    oTable.name != null ? oTable.name : undefined,
			"autoFormatId":            oTable.autoFormatId != null ? oTable.autoFormatId : undefined,
			"growShrinkType":          oTable.growShrinkType != null ? oTable.growShrinkType : undefined,
			"adjustColumnWidth":       oTable.adjustColumnWidth != null ? oTable.adjustColumnWidth : undefined,
			"applyAlignmentFormats":   oTable.applyAlignmentFormats != null ? oTable.applyAlignmentFormats : undefined,
			"applyBorderFormats":      oTable.applyBorderFormats != null ? oTable.applyBorderFormats : undefined,
			"applyFontFormats":        oTable.applyFontFormats != null ? oTable.applyFontFormats : undefined,
			"applyNumberFormats":      oTable.applyNumberFormats != null ? oTable.applyNumberFormats : undefined,
			"applyPatternFormats":     oTable.ApplyPatternFormats != null ? oTable.ApplyPatternFormats : undefined,
			"applyWidthHeightFormats": oTable.applyWidthHeightFormats != null ? oTable.applyWidthHeightFormats : undefined,
			"backgroundRefresh":       oTable.backgroundRefresh != null ? oTable.backgroundRefresh : undefined,
			"disableEdit":             oTable.disableEdit != null ? oTable.disableEdit : undefined,
			"disableRefresh":          oTable.disableRefresh != null ? oTable.disableRefresh : undefined,
			"fillFormulas":            oTable.fillFormulas != null ? oTable.fillFormulas : undefined,
			"firstBackgroundRefresh":  oTable.firstBackgroundRefresh != null ? oTable.firstBackgroundRefresh : undefined,
			"headers":                 oTable.headers != null ? oTable.headers : undefined,
			"intermediate":            oTable.intermediate != null ? oTable.intermediate : undefined,
			"preserveFormatting":      oTable.preserveFormatting != null ? oTable.preserveFormatting : undefined,
			"refreshOnLoad":           oTable.refreshOnLoad != null ? oTable.refreshOnLoad : undefined,
			"removeDataOnSave":        oTable.removeDataOnSave != null ? oTable.removeDataOnSave : undefined,
			"rowNumbers":              oTable.rowNumbers != null ? oTable.rowNumbers : undefined,
			"queryTableRefresh":       oTable.queryTableRefresh != null ? this.SerQueryTableRefresh(oTable.queryTableRefresh) : undefined
		}
	};
	WriterToJSON.prototype.SerQueryTableRefresh = function(oTable)
	{
		if (!oTable)
			return oTable;

		return {
			"nextId":                   oTable.nextId != null ? oTable.nextId : undefined,
			"minimumVersion":           oTable.minimumVersion != null ? oTable.minimumVersion : undefined,
			"unboundColumnsLeft":       oTable.unboundColumnsLeft != null ? oTable.unboundColumnsLeft : undefined,
			"unboundColumnsRight":      oTable.unboundColumnsRight != null ? oTable.unboundColumnsRight : undefined,
			"fieldIdWrapped":           oTable.fieldIdWrapped != null ? oTable.fieldIdWrapped : undefined,
			"headersInLastRefresh":     oTable.headersInLastRefresh != null ? oTable.headersInLastRefresh : undefined,
			"preserveSortFilterLayout": oTable.preserveSortFilterLayout != null ? oTable.preserveSortFilterLayout : undefined,
			"sortState":                oTable.sortState != null ? this.SerSortState(oTable.sortState) : undefined,
			"queryTableFields":         oTable.queryTableFields != null ? this.SerQueryTableFields(oTable.queryTableFields) : undefined,
			"queryTableDeletedFields":  oTable.queryTableDeletedFields != null ? this.SerQueryTableDeletedFields(oTable.queryTableDeletedFields) : undefined
		}
	};
	WriterToJSON.prototype.SerQueryTableFields = function(aFields)
	{
		var aResult = [];
		for (var nFld = 0; nFld < aFields.length; nFld++)
			aResult.push(this.SerQueryTableField(aFields[nFld]));

		return aResult;
	};
	WriterToJSON.prototype.SerQueryTableField = function(oField)
	{
		return {
			"name":          oField.name != null ? oField.name : undefined,
			"id":            oField.id != null ? oField.id : undefined,
			"tableColumnId": oField.tableColumnId != null ? oField.tableColumnId : undefined,
			"rowNumbers":    oField.rowNumbers != null ? oField.rowNumbers : undefined,
			"fillFormulas":  oField.fillFormulas != null ? oField.fillFormulas : undefined,
			"dataBound":     oField.dataBound != null ? oField.dataBound : undefined,
			"clipped":       oField.clipped != null ? oField.clipped : undefined
		}
	};
	WriterToJSON.prototype.SerQueryTableDeletedFields = function(aFields)
	{
		var aResult = [];
		var oCurField;
		for (var nFld = 0; nFld < aFields.length; nFld++)
		{
			oCurField = aFields[nFld];
			aResult.push({
				"name": oCurField.name != null ? oCurField.name : undefined
			});
		}

		return aResult;
	};
	WriterToJSON.prototype.SerComments = function(aComments)
	{
		var aResult = [];
		var oCurComm;
		var aPersonLst = [];
		for (var nComm = 0; nComm < aComments.length; nComm++)
		{
			oCurComm = aComments[nComm];
			aResult.push({
				"coord": this.SerCommentCoords(oCurComm.coords),
				"data":  this.SerCommentData(oCurComm, aPersonLst)
			});
		}

		return aResult;
	};
	WriterToJSON.prototype.SerCommentCoords = function(oCoords)
	{
		if (!oCoords)
			return oCoords;

		return {
			"row":           oCoords.nRow != null ? oCoords.nRow : undefined,
			"col":           oCoords.nCol != null ? oCoords.nCol : undefined,
			"left":          oCoords.nLeft != null ? oCoords.nLeft : undefined,
			"leftOff":       oCoords.nLeftOffset != null ? oCoords.nLeftOffset : undefined,
			"topOff":        oCoords.nTopOffset != null ? oCoords.nTopOffset : undefined,
			"right":         oCoords.nRight != null ? oCoords.nRight : undefined,
			"rightOff":      oCoords.nRightOffset != null ? oCoords.nRightOffset : undefined,
			"bottom":        oCoords.nBottom != null ? oCoords.nBottom : undefined,
			"bottomOff":     oCoords.nBottomOffset != null ? oCoords.nBottomOffset : undefined,
			"leftMM":        oCoords.dLeftMM != null ? oCoords.dLeftMM : undefined,
			"topMM":         oCoords.dTopMM != null ? oCoords.dTopMM : undefined,
			"widthMM":       oCoords.dWidthMM != null ? oCoords.dWidthMM : undefined,
			"heightMM":      oCoords.dHeightMM != null ? oCoords.dHeightMM : undefined,
			"moveWithCells": oCoords.bMoveWithCells != null ? oCoords.bMoveWithCells : undefined,
			"sizeWithCells": oCoords.bSizeWithCells != null ? oCoords.bSizeWithCells : undefined,
		}
	};
	WriterToJSON.prototype.SerCommentData = function(oData, aPersonLst)
	{
		var aReplies = [];
		for (var nReply = 0; nReply < oData.aReplies.length; nReply++)
			aReplies.push(this.SerCommentData(oData.aReplies[nReply], aPersonLst));

		var sOOTime = oData.sOOTime;
		var sdTime = oData.sTime;
		if (sOOTime != null && sOOTime != "")
			sOOTime = new Date(sOOTime - 0).toISOString().slice(0, 22) + "Z";
		else if (sdTime != null && sdTime != "")
			sOOTime = new Date(sdTime - 0).toISOString().slice(0, 22) + "Z";

		var userId = oData.sUserId;
		var displayName = oData.sUserName;
		var providerId = oData.sProviderId;

		var person = aPersonLst.find(function isPrime(element) {
			return userId === element.userId && displayName === element.displayName && providerId === element.providerId;
		});
		if (!person) {
			person = {id: AscCommon.CreateGUID(), userId: userId, displayName: displayName, providerId: providerId};
			aPersonLst.push(person);
		}

		var guid = oData.sGuid;
		var solved = oData.bSolved;
		var text = oData.sText;
		var sUserId = oData.sUserId;
		var sUserName = oData.sUserName;
		//var userData = oData.m_sUserData;

		return {
			"text":        text != null ? text : undefined,
			"dt":          sOOTime != null ? sOOTime : undefined,
			"personId":    person.id,
			"displayName": sUserName != null ? sUserName : undefined,
			"userId":      sUserId != null ? sUserId : undefined,
			"id":          guid != null ? guid : undefined,
			"providerId":  providerId != null ? providerId : undefined,
			"done":        solved != null ? solved : undefined, 
			"replies":     aReplies.length > 0 ? aReplies : undefined
		}
	};
	WriterToJSON.prototype.SerSparklineGroups = function(aGroup)
	{
		var aResult = [];
		for (var nIndex = 0; nIndex < aGroup.length; nIndex++)
			aResult.push(this.SerSparklineGroup(aGroup[nIndex]));
		
		return aResult;
	};
	WriterToJSON.prototype.SerSparklineGroup = function(oGroup)
	{
		var aSparklines = [];

		for (var nLine = 0; nLine < oGroup.arrSparklines.length; nLine++)
			aSparklines.push(this.SerSparkLine(oGroup.arrSparklines[nLine]));

		return {
			"manualMax":           oGroup.manualMax != null ? oGroup.manualMax : undefined,
			"manualMin":           oGroup.manualMin != null ? oGroup.manualMin : undefined,
			"lineWeight":          oGroup.lineWeight != null ? oGroup.lineWeight : undefined,
			"type":                oGroup.type != null ? ToXML_ST_SparklineType(oGroup.type) : undefined,
			"dateAxis":            oGroup.dateAxis != null ? oGroup.dateAxis : undefined,
			"displayEmptyCellsAs": oGroup.displayEmptyCellsAs != null ? ToXML_ST_DispBlanksAs(oGroup.displayEmptyCellsAs) : undefined,
			"markers":             oGroup.markers != null ? oGroup.markers : undefined,
			"high":                oGroup.high != null ? oGroup.high : undefined,
			"low":                 oGroup.low != null ? oGroup.low : undefined,
			"first":               oGroup.first != null ? oGroup.first : undefined,
			"last":                oGroup.last != null ? oGroup.last : undefined,
			"negative":            oGroup.negative != null ? oGroup.negative : undefined,
			"displayXAxis":        oGroup.displayXAxis != null ? oGroup.displayXAxis : undefined,
			"displayHidden":       oGroup.displayHidden != null ? oGroup.displayHidden : undefined,
			"minAxisType":         oGroup.minAxisType != null ? ToXML_ST_SparklineAxisMinMax(oGroup.minAxisType) : undefined,
			"maxAxisType":         oGroup.maxAxisType != null ? ToXML_ST_SparklineAxisMinMax(oGroup.maxAxisType) : undefined,
			"rightToLeft":         oGroup.rightToLeft != null ? oGroup.rightToLeft : undefined,
			"colorSeries":         oGroup.colorSeries != null ? this.SerColorExcell(oGroup.colorSeries) : undefined,
			"colorNegative":       oGroup.colorNegative != null ? this.SerColorExcell(oGroup.colorNegative) : undefined,
			"colorAxis":           oGroup.colorAxis != null ? this.SerColorExcell(oGroup.colorAxis) : undefined,
			"colorMarkers":        oGroup.colorMarkers != null ? this.SerColorExcell(oGroup.colorMarkers) : undefined,
			"colorFirst":          oGroup.colorFirst != null ? this.SerColorExcell(oGroup.colorFirst) : undefined,
			"colorLast":           oGroup.colorLast != null ? this.SerColorExcell(oGroup.colorLast) : undefined,
			"colorHigh":           oGroup.colorHigh != null ? this.SerColorExcell(oGroup.colorHigh) : undefined,
			"colorLow":            oGroup.colorLow != null ? this.SerColorExcell(oGroup.colorLow) : undefined,
			"f":                   oGroup.f != null ? oGroup.f : undefined,
			"sparklines":          aSparklines
		}
	};
	WriterToJSON.prototype.SerSparkLine = function(oSparkLine)
	{
		return {
			"f":     oSparkLine.f != null ? oSparkLine.f : undefined,
			"sqRef": oSparkLine.sqRef != null ? oSparkLine.sqRef.getName() : undefined
		}
	};
	WriterToJSON.prototype.SerHdrFtrExcell = function(oHdrFtr)
	{
		if (!oHdrFtr)
			return oHdrFtr;

		return {
			"alignWithMargins": oHdrFtr.alignWithMargins != null ? oHdrFtr.alignWithMargins : undefined,
			"differentFirst":   oHdrFtr.differentFirst != null ? oHdrFtr.differentFirst : undefined,
			"differentOddEven": oHdrFtr.differentOddEven != null ? oHdrFtr.differentOddEven : undefined,
			"scaleWithDoc":     oHdrFtr.scaleWithDoc != null ? oHdrFtr.scaleWithDoc : undefined,
			"evenFooter":       oHdrFtr.evenFooter != null ? oHdrFtr.evenFooter.getStr() : undefined,
			"evenHeader":       oHdrFtr.evenHeader != null ? oHdrFtr.evenHeader.getStr() : undefined,
			"firstFooter":      oHdrFtr.firstFooter != null ? oHdrFtr.firstFooter.getStr() : undefined,
			"firstHeader":      oHdrFtr.firstHeader != null ? oHdrFtr.firstHeader.getStr() : undefined,
			"oddFooter":        oHdrFtr.oddFooter != null ? oHdrFtr.oddFooter.getStr() : undefined,
			"oddHeader":        oHdrFtr.oddHeader != null ? oHdrFtr.oddHeader.getStr() : undefined
		}
	};
	WriterToJSON.prototype.SerPivotTables = function(aTables)
	{
		var aResult = [];
		for (var nTable = 0; nTable < aTables.length; nTable++)
			aResult.push(this.SerPivotTable(aTables[nTable]));

		return aResult;
	};
	WriterToJSON.prototype.SerPivotTable = function(oTable) // CT_pivotTableDefinition
	{
		return {
			"name":                    oTable.name != null ? oTable.name : undefined,
			"cacheId":                 oTable.cacheId != null ? oTable.cacheId : undefined,
			"dataOnRows":              oTable.dataOnRows != null ? oTable.dataOnRows : undefined,
			"dataPosition":            oTable.dataPosition != null ? oTable.dataPosition : undefined,
			"autoFormatId":            oTable.autoFormatId != null ? oTable.autoFormatId : undefined,
			"applyNumberFormats":      oTable.applyNumberFormats != null ? oTable.applyNumberFormats : undefined,
			"applyBorderFormats":      oTable.applyBorderFormats != null ? oTable.applyBorderFormats : undefined,
			"applyFontFormats":        oTable.applyFontFormats != null ? oTable.applyFontFormats : undefined,
			"applyPatternFormats":     oTable.applyPatternFormats != null ? oTable.applyPatternFormats : undefined,
			"applyAlignmentFormats":   oTable.applyAlignmentFormats != null ? oTable.applyAlignmentFormats : undefined,
			"applyWidthHeightFormats": oTable.applyWidthHeightFormats != null ? oTable.applyWidthHeightFormats : undefined,
			"dataCaption":             oTable.dataCaption != null ? oTable.dataCaption : undefined,
			"grandTotalCaption":       oTable.grandTotalCaption != null ? oTable.grandTotalCaption : undefined,
			"errorCaption":            oTable.errorCaption != null ? oTable.errorCaption : undefined,
			"showError":               oTable.showError != null ? oTable.showError : undefined,
			"missingCaption":          oTable.missingCaption != null ? oTable.missingCaption : undefined,
			"showMissing":             oTable.showMissing != null ? oTable.showMissing : undefined,
			"pageStyle":               oTable.pageStyle != null ? oTable.pageStyle : undefined,
			"pivotTableStyle":         oTable.pivotTableStyle != null ? oTable.pivotTableStyle : undefined,
			"vacatedStyle":            oTable.vacatedStyle != null ? oTable.vacatedStyle : undefined,
			"tag":                     oTable.tag != null ? oTable.tag : undefined,
			"updatedVersion":          oTable.updatedVersion != null ? oTable.updatedVersion : undefined,
			"minRefreshableVersion":   oTable.minRefreshableVersion != null ? oTable.minRefreshableVersion : undefined,
			"asteriskTotals":          oTable.asteriskTotals != null ? oTable.asteriskTotals : undefined,
			"showItems":               oTable.showItems != null ? oTable.showItems : undefined,
			"editData":                oTable.editData != null ? oTable.editData : undefined,
			"disableFieldList":        oTable.disableFieldList != null ? oTable.disableFieldList : undefined,
			"showCalcMbrs":            oTable.showCalcMbrs != null ? oTable.showCalcMbrs : undefined,
			"visualTotals":            oTable.visualTotals != null ? oTable.visualTotals : undefined,
			"showMultipleLabel":       oTable.showMultipleLabel != null ? oTable.showMultipleLabel : undefined,
			"showDataDropDown":        oTable.showDataDropDown != null ? oTable.showDataDropDown : undefined,
			"showDrill":               oTable.showDrill != null ? oTable.showDrill : undefined,
			"printDrill":              oTable.printDrill != null ? oTable.printDrill : undefined,
			"showMemberPropertyTips":  oTable.showMemberPropertyTips != null ? oTable.showMemberPropertyTips : undefined,
			"showDataTips":            oTable.showDataTips != null ? oTable.showDataTips : undefined,
			"enableWizard":            oTable.enableWizard != null ? oTable.enableWizard : undefined,
			"enableDrill":             oTable.enableDrill != null ? oTable.enableDrill : undefined,
			"enableFieldProperties":   oTable.enableFieldProperties != null ? oTable.enableFieldProperties : undefined,
			"preserveFormatting":      oTable.preserveFormatting != null ? oTable.preserveFormatting : undefined,
			"useAutoFormatting":       oTable.useAutoFormatting != null ? oTable.useAutoFormatting : undefined,
			"pageWrap":                oTable.pageWrap != null ? oTable.pageWrap : undefined,
			"pageOverThenDown":        oTable.pageOverThenDown != null ? oTable.pageOverThenDown : undefined,
			"subtotalHiddenItems":     oTable.subtotalHiddenItems != null ? oTable.subtotalHiddenItems : undefined,
			"rowGrandTotals":          oTable.rowGrandTotals != null ? oTable.rowGrandTotals : undefined,
			"colGrandTotals":          oTable.colGrandTotals != null ? oTable.colGrandTotals : undefined,
			"fieldPrintTitles":        oTable.fieldPrintTitles != null ? oTable.fieldPrintTitles : undefined,
			"itemPrintTitles":         oTable.itemPrintTitles != null ? oTable.itemPrintTitles : undefined,
			"mergeItem":               oTable.mergeItem != null ? oTable.mergeItem : undefined,
			"showDropZones":           oTable.showDropZones != null ? oTable.showDropZones : undefined,
			"createdVersion":          oTable.createdVersion != null ? oTable.createdVersion : undefined,
			"indent":                  oTable.indent != null ? oTable.indent : undefined,
			"showEmptyRow":            oTable.showEmptyRow != null ? oTable.showEmptyRow : undefined,
			"showEmptyCol":            oTable.showEmptyCol != null ? oTable.showEmptyCol : undefined,
			"showHeaders":             oTable.showHeaders != null ? oTable.showHeaders : undefined,
			"compact":                 oTable.compact != null ? oTable.compact : undefined,
			"outline":                 oTable.outline != null ? oTable.outline : undefined,
			"outlineData":             oTable.outlineData != null ? oTable.outlineData : undefined,
			"compactData":             oTable.compactData != null ? oTable.compactData : undefined,
			"published":               oTable.published != null ? oTable.published : undefined,
			"gridDropZones":           oTable.gridDropZones != null ? oTable.gridDropZones : undefined,
			"immersive":               oTable.immersive != null ? oTable.immersive : undefined,
			"multipleFieldFilters":    oTable.multipleFieldFilters != null ? oTable.multipleFieldFilters : undefined,
			"chartFormat":             oTable.chartFormat != null ? oTable.chartFormat : undefined,
			"rowHeaderCaption":        oTable.rowHeaderCaption != null ? oTable.rowHeaderCaption : undefined,
			"colHeaderCaption":        oTable.colHeaderCaption != null ? oTable.colHeaderCaption : undefined,
			"fieldListSortAscending":  oTable.fieldListSortAscending != null ? oTable.fieldListSortAscending : undefined,
			"mdxSubqueries":           oTable.mdxSubqueries != null ? oTable.mdxSubqueries : undefined,
			"customListSort":          oTable.customListSort != null ? oTable.customListSort : undefined,
			"location":                oTable.location != null ? this.SerLocation(oTable.location) : undefined,
			"pivotFields":             oTable.pivotFields != null ? this.SerPivotFields(oTable.pivotFields) : undefined,
			"rowFields":               oTable.rowFields != null ? this.SerRowFields(oTable.rowFields) : undefined,
			"rowItems":                oTable.rowItems != null ? this.SerRowItems(oTable.rowItems) : undefined,
			"colFields":               oTable.colFields != null ? this.SerColFields(oTable.colFields) : undefined,
			"colItems":                oTable.colItems != null ? this.SerColItems(oTable.colItems) : undefined,
			"pageFields":              oTable.pageFields != null ? this.SerPageFields(oTable.pageFields) : undefined,
			"dataFields":              oTable.dataFields != null ? this.SerDataFields(oTable.dataFields) : undefined,
			"formats":                 oTable.formats != null ? this.SerFormats(oTable.formats) : undefined,
			"conditionalFormats":      oTable.conditionalFormats != null ? this.SerConditionalFormats(oTable.conditionalFormats) : undefined,
			"chartFormats":            oTable.chartFormats != null ? this.SerChartFormats(oTable.chartFormats) : undefined,
			"pivotHierarchies":        oTable.pivotHierarchies != null ? this.SerPivotHierarchies(oTable.pivotHierarchies) : undefined,
			"pivotTableStyleInfo":     oTable.pivotTableStyleInfo != null ? this.SerPivotTableStyleInfo(oTable.pivotTableStyleInfo) : undefined,
			"filters":                 oTable.filters != null ? this.SerPivotFilters(oTable.filters) : undefined,
			"rowHierarchiesUsage":     oTable.rowHierarchiesUsage != null ? this.SerRowHierarchiesUsage(oTable.rowHierarchiesUsage) : undefined,
			"colHierarchiesUsage":     oTable.colHierarchiesUsage != null ? this.SerColHierarchiesUsage(oTable.colHierarchiesUsage) : undefined,
			"pivotTableDefinitionX14": oTable.pivotTableDefinitionX14 != null ? this.SerPivotTableDefinitionX14(oTable.pivotTableDefinitionX14) : undefined
		}
	};
	WriterToJSON.prototype.SerLocation = function(oLocation)
	{
		if (!oLocation)
			return oLocation;

		return {
			"ref":            oLocation.ref != null ? this.SerRef(oLocation.ref) : undefined,
			"firstHeaderRow": oLocation.firstHeaderRow != null ? oLocation.firstHeaderRow : undefined,
			"firstDataRow":   oLocation.firstDataRow != null ? oLocation.firstDataRow : undefined,
			"firstDataCol":   oLocation.firstDataCol != null ? oLocation.firstDataCol : undefined,
			"rowPageCount":   oLocation.rowPageCount != null ? oLocation.rowPageCount : undefined,
			"colPageCount":   oLocation.colPageCount != null ? oLocation.colPageCount : undefined
		}
	};
	WriterToJSON.prototype.SerPivotFields = function(oPivotFields)
	{
		var aFields = [];
		for (var nField = 0; nField < oPivotFields.pivotField.length; nField++)
			aFields.push(this.SerPivotField(oPivotFields.pivotField[nField]));

		return {
			"count":      oPivotFields.pivotField.length,
			"pivotField": aFields
		}
	};
	WriterToJSON.prototype.SerPivotField = function(oField)
	{
		return {
			"name":                         oField.name != null ? oField.name : undefined,
			"axis":                         oField.axis != null ? ToXml_ST_Axis(oField.axis) : undefined,
			"dataField":                    oField.dataField != null ? oField.dataField : undefined,
			"subtotalCaption":              oField.subtotalCaption != null ? oField.subtotalCaption : undefined,
			"showDropDowns":                oField.showDropDowns != null ? oField.showDropDowns : undefined,
			"hiddenLevel":                  oField.hiddenLevel != null ? oField.hiddenLevel : undefined,
			"uniqueMemberProperty":         oField.uniqueMemberProperty != null ? oField.uniqueMemberProperty : undefined,
			"compact":                      oField.compact != null ? oField.compact : undefined,
			"allDrilled":                   oField.allDrilled != null ? oField.allDrilled : undefined,
			"numFmtId":                     oField.num != null ? this.SerNumFmtExcell(oField.num, undefined, true) : undefined,
			"outline":                      oField.outline != null ? oField.outline : undefined,
			"subtotalTop":                  oField.subtotalTop != null ? oField.subtotalTop : undefined,
			"dragToRow":                    oField.dragToRow != null ? oField.dragToRow : undefined,
			"dragToCol":                    oField.dragToCol != null ? oField.dragToCol : undefined,
			"multipleItemSelectionAllowed": oField.multipleItemSelectionAllowed != null ? oField.multipleItemSelectionAllowed : undefined,
			"dragToPage":                   oField.dragToPage != null ? oField.dragToPage : undefined,
			"dragToData":                   oField.dragToData != null ? oField.dragToData : undefined,
			"dragOff":                      oField.dragOff != null ? oField.dragOff : undefined,
			"showAll":                      oField.showAll != null ? oField.showAll : undefined,
			"insertBlankRow":               oField.insertBlankRow != null ? oField.insertBlankRow : undefined,
			"serverField":                  oField.serverField != null ? oField.serverField : undefined,
			"insertPageBreak":              oField.insertPageBreak != null ? oField.insertPageBreak : undefined,
			"autoShow":                     oField.autoShow != null ? oField.autoShow : undefined,
			"topAutoShow":                  oField.topAutoShow != null ? oField.topAutoShow : undefined,
			"hideNewItems":                 oField.hideNewItems != null ? oField.hideNewItems : undefined,
			"measureFilter":                oField.measureFilter != null ? oField.measureFilter : undefined,
			"includeNewItemsInFilter":      oField.includeNewItemsInFilter != null ? oField.includeNewItemsInFilter : undefined,
			"itemPageCount":                oField.itemPageCount != null ? oField.itemPageCount : undefined, // if (10 !== this.itemPageCount) {
			"sortType":                     oField.sortType != null ? ToXml_ST_FieldSortType(oField.sortType) : undefined,
			"dataSourceSort":               oField.dataSourceSort != null ? oField.dataSourceSort : undefined,
			"nonAutoSortDefault":           oField.nonAutoSortDefault != null ? oField.nonAutoSortDefault : undefined,
			"rankBy":                       oField.rankBy != null ? oField.rankBy : undefined,
			"defaultSubtotal":              oField.defaultSubtotal != null ? oField.defaultSubtotal : undefined,
			"sumSubtotal":                  oField.sumSubtotal != null ? oField.sumSubtotal : undefined,
			"countASubtotal":               oField.countASubtotal != null ? oField.countASubtotal : undefined,
			"avgSubtotal":                  oField.avgSubtotal != null ? oField.avgSubtotal : undefined,
			"maxSubtotal":                  oField.maxSubtotal != null ? oField.maxSubtotal : undefined,
			"minSubtotal":                  oField.minSubtotal != null ? oField.minSubtotal : undefined,
			"productSubtotal":              oField.productSubtotal != null ? oField.productSubtotal : undefined,
			"countSubtotal":                oField.countSubtotal != null ? oField.countSubtotal : undefined,
			"stdDevSubtotal":               oField.stdDevSubtotal != null ? oField.stdDevSubtotal : undefined,
			"stdDevPSubtotal":              oField.stdDevPSubtotal != null ? oField.stdDevPSubtotal : undefined,
			"varSubtotal":                  oField.varSubtotal != null ? oField.varSubtotal : undefined,
			"varPSubtotal":                 oField.varPSubtotal != null ? oField.varPSubtotal : undefined,
			"showPropCell":                 oField.showPropCell != null ? oField.showPropCell : undefined,
			"showPropTip":                  oField.showPropTip != null ? oField.showPropTip : undefined,
			"showPropAsCaption":            oField.showPropAsCaption != null ? oField.showPropAsCaption : undefined,
			"defaultAttributeDrillState":   oField.defaultAttributeDrillState != null ? oField.defaultAttributeDrillState : undefined,
			"pivotFieldX14":                oField.pivotFieldX14 != null ? this.SerPivotFieldX14(oField.pivotFieldX14) : undefined,
			"items":                        oField.items != null ? this.SerPivotFieldItems(oField.items) : undefined,
			"autoSortScope":                oField.autoSortScope != null ? this.SerAutoSortScope(oField.autoSortScope) : undefined
		}
	};
	WriterToJSON.prototype.SerExtensionList = function(oPivotField)
	{
		if (!oPivotField)
			return oPivotField;

		var aResult = [];
		for (var nExt = 0; nExt < oPivotField.ext.length; nExt++)
			aResult.push(this.SerExtension(oPivotField.ext[nExt]));
	
		return {
			"ext": aResult
		};
	};
	WriterToJSON.prototype.SerExtension = function(oExt)
	{
		if (!oExt || !oExt.elem)
			return null;

		var oElem;
		switch (oExt.uri)
		{
			case "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}":
				oElem = this.SerPivotTableDefinitionX14(oExt.elem);
				break;
			case "{725AE2AE-9491-48be-B2B4-4EB974FC3084}":
				oElem = this.SerPivotCacheDefinitionX14(oExt.elem);
				break;
			case "{2946ED86-A175-432a-8AC1-64E0C546D7DE}":
				oElem = this.SerPivotFieldX14(oExt.elem);
				break;
		}

		return {
			"uri":  oExt.uri != null ? oExt.uri : undefined,
			"elem": oElem
		}
	};
	WriterToJSON.prototype.SerPivotFieldX14 = function(oField)
	{
		if (!oField)
			return oField;

		return {
			"fillDownLabels": oField.fillDownLabels != null ? oField.fillDownLabels : undefined,
			"ignore":         oField.ignore != null ? oField.ignore : undefined,
			"type":           "x14:pivotField"
		}
	};
	WriterToJSON.prototype.SerPivotCacheDefinitionX14 = function(oDef)
	{
		if (!oDef)
			return oDef;

		return {
			"slicerData":               oDef.slicerData != null ? oDef.slicerData : undefined,
			"pivotCacheId":             oDef.pivotCacheId != null ? oDef.pivotCacheId : undefined,
			"supportSubqueryNonVisual": oDef.supportSubqueryNonVisual != null ? oDef.supportSubqueryNonVisual : undefined,
			"supportSubqueryCalcMem":   oDef.supportSubqueryCalcMem != null ? oDef.supportSubqueryCalcMem : undefined,
			"supportAddCalcMems":       oDef.supportAddCalcMems != null ? oDef.supportAddCalcMems : undefined,
			"type":                     "x14:pivotCacheDefinition"
		}
	};
	WriterToJSON.prototype.SerPivotTableDefinitionX14 = function(oDef)
	{
		if (!oDef)
			return oDef;

		return {
			"fillDownLabelsDefault":      oDef.fillDownLabelsDefault != null ? oDef.fillDownLabelsDefault : undefined,
			"visualTotalsForSets":        oDef.visualTotalsForSets != null ? oDef.visualTotalsForSets : undefined,
			"calculatedMembersInFilters": oDef.calculatedMembersInFilters != null ? oDef.calculatedMembersInFilters : undefined,
			"altText":                    oDef.altText != null ? oDef.altText : undefined,
			"altTextSummary":             oDef.altTextSummary != null ? oDef.altTextSummary : undefined,
			"enableEdit":                 oDef.enableEdit != null ? oDef.enableEdit : undefined,
			"autoApply":                  oDef.autoApply != null ? oDef.autoApply : undefined,
			"allocationMethod":           oDef.allocationMethod != null ? ToXml_ST_AllocationMethod(oDef.allocationMethod) : undefined,
			"weightExpression":           oDef.weightExpression != null ? oDef.weightExpression : undefined,
			"hideValuesRow":              oDef.hideValuesRow != null ? oDef.hideValuesRow : undefined,
			"type":                       "x14:pivotTableDefinition"
		}
	};
	WriterToJSON.prototype.SerPivotFieldItems = function(oItems) // CT_Items
	{
		if (!oItems)
			return oItems;

		var aResult = [];
		for (var nItem = 0; nItem < oItems.item.length; nItem++)
			aResult.push(this.SerPivotFieldItem(oItems.item[nItem]));

		return {
			"item": aResult
		}
	};
	WriterToJSON.prototype.SerPivotFieldItem = function(oItem) // CT_Item
	{
		return {
			"n":  oItem.n != null ? oItem.n : undefined,
			"t":  oItem.t != null ? ToXml_ST_ItemType(oItem.t) : undefined,
			"h":  oItem.h != null ? oItem.h : undefined,
			"s":  oItem.s != null ? oItem.s : undefined,
			"sd": oItem.sd != null ? oItem.sd : undefined,
			"f":  oItem.f != null ? oItem.f : undefined,
			"m":  oItem.m != null ? oItem.m : undefined,
			"c":  oItem.c != null ? oItem.c : undefined,
			"x":  oItem.x != null ? oItem.x : undefined,
			"d":  oItem.d != null ? oItem.d : undefined,
			"e":  oItem.e != null ? oItem.e : undefined
		}
	};
	WriterToJSON.prototype.SerAutoSortScope = function(oSortScope)
	{
		if (!oSortScope)
			return oSortScope;

		return {
			"pivotArea": oSortScope.pivotArea != null ? this.SerPivotArea(oSortScope.pivotArea) : undefined
		}
	};
	WriterToJSON.prototype.SerPivotArea = function(oPivotArea)
	{
		if (!oPivotArea)
			return oPivotArea;

		return {
			"field":                       oPivotArea.field != null ? oPivotArea.field : undefined,
			"type":                        oPivotArea.type != null ? ToXml_ST_PivotAreaType(oPivotArea.type) : undefined,
			"dataOnly":                    oPivotArea.dataOnly != null ? oPivotArea.dataOnly : undefined,
			"labelOnly":                   oPivotArea.labelOnly != null ? oPivotArea.labelOnly : undefined,
			"grandRow":                    oPivotArea.grandRow != null ? oPivotArea.grandRow : undefined,
			"grandCol":                    oPivotArea.grandCol != null ? oPivotArea.grandCol : undefined,
			"cacheIndex":                  oPivotArea.cacheIndex != null ? oPivotArea.cacheIndex : undefined,
			"outline":                     oPivotArea.outline != null ? oPivotArea.outline : undefined,
			"offset":                      oPivotArea.offset != null ? oPivotArea.offset : undefined,
			"collapsedLevelsAreSubtotals": oPivotArea.collapsedLevelsAreSubtotals != null ? oPivotArea.collapsedLevelsAreSubtotals : undefined,
			"axis":                        oPivotArea.axis != null ? ToXml_ST_Axis(oPivotArea.axis) : undefined,
			"fieldPosition":               oPivotArea.fieldPosition != null ? oPivotArea.fieldPosition : undefined,
			"references":                  oPivotArea.references != null ? this.SerPivotAreaRefs(oPivotArea.references) : undefined,
			"extLst":                      oPivotArea.extLst != null ? this.SerExtensionList(oPivotArea.extLst) : undefined
		}
	};
	WriterToJSON.prototype.SerPivotAreaRefs = function(oRefs) // CT_PivotAreaReferences
	{
		if (!oRefs)
			return oRefs;

		var aRefs = [];
		for (var nRef = 0; nRef < oRefs.reference.length; nRef++)
			aRefs.push(this.SerPivotAreaRef(oRefs.reference[nRef]));

		return aRefs;
	};
	WriterToJSON.prototype.SerPivotAreaRef = function(oRef) // CT_PivotAreaReference
	{
		var aIndexes = [];
		for (var nIdx = 0; nIdx < oRef.x.length; nIdx++)
			aIndexes.push(this.SerIndex(oRef.x[nIdx]));

		return {
			"field":           oRef.field != null ? oRef.field : undefined,
			"count":           oRef.x.length,
			"selected":        oRef.selected != null ? oRef.selected : undefined,
			"byPosition":      oRef.byPosition != null ? oRef.byPosition : undefined,
			"relative":        oRef.relative != null ? oRef.relative : undefined,
			"defaultSubtotal": oRef.defaultSubtotal != null ? oRef.defaultSubtotal : undefined,
			"sumSubtotal":     oRef.sumSubtotal != null ? oRef.sumSubtotal : undefined,
			"countASubtotal":  oRef.countASubtotal != null ? oRef.countASubtotal : undefined,
			"avgSubtotal":     oRef.avgSubtotal != null ? oRef.avgSubtotal : undefined,
			"maxSubtotal":     oRef.maxSubtotal != null ? oRef.maxSubtotal : undefined,
			"minSubtotal":     oRef.minSubtotal != null ? oRef.minSubtotal : undefined,
			"productSubtotal": oRef.productSubtotal != null ? oRef.productSubtotal : undefined,
			"countSubtotal":   oRef.countSubtotal != null ? oRef.countSubtotal : undefined,
			"stdDevSubtotal":  oRef.stdDevSubtotal != null ? oRef.stdDevSubtotal : undefined,
			"stdDevPSubtotal": oRef.stdDevPSubtotal != null ? oRef.stdDevPSubtotal : undefined,
			"varSubtotal":     oRef.varSubtotal != null ? oRef.varSubtotal : undefined,
			"varPSubtotal":    oRef.varPSubtotal != null ? oRef.varPSubtotal : undefined,
			"x":               aIndexes,
			"extLst":          oRef.extLst != null ? this.SerExtensionList(oRef.extLst) : undefined
		}
	};
	WriterToJSON.prototype.SerIndex = function(oIndex) // CT_Index
	{
		return {
			"v": oIndex.v
		}
	};
	WriterToJSON.prototype.SerRowFields = function(oRowFields) //CT_RowFields
	{
		if (!oRowFields)
			return oRowFields;

		var aResult = [];
		for (var nField = 0; nField < oRowFields.field.length; nField++)
			aResult.push(this.SerField(oRowFields.field[nField]));

		return {
			"field": aResult
		};
	};
	WriterToJSON.prototype.SerField = function(oField) // CT_Field
	{
		return {
			"x": oField.x != null ? oField.x : undefined
		}
	};
	WriterToJSON.prototype.SerRowItems = function(oRowItems) // CT_rowItems
	{
		if (!oRowItems)
			return oRowItems;

		var aResult = [];
		for (var nItem = 0; nItem < oRowItems.i.length; nItem++)
			aResult.push(this.SerI(oRowItems.i[nItem]));

		return {
			"i": aResult	
		};
	};
	WriterToJSON.prototype.SerI = function(oI) // CT_I
	{
		var aX = [];
		for (var nX = 0; nX < oI.x.length; nX++)
			aX.push(this.SerX(oI.x[nX]));

		return {
			"t": oI.t != null ? ToXml_ST_ItemType(oI.t) : undefined,
			"r": oI.r != null ? oI.r : undefined,
			"i": oI.i != null ? oI.i : undefined,
			"x": aX
		}
	};
	WriterToJSON.prototype.SerX = function(oX) // CT_X
	{
		return {
			"v": oX.v !== 0 && oX.v != null ? oX.v : undefined
		}
	};
	WriterToJSON.prototype.SerColFields = function(oColFields) // CT_ColFields
	{
		if (!oColFields)
			return oColFields;

		var aResult = [];
		for (var nField = 0; nField < oColFields.field.length; nField++)
			aResult.push(this.SerField(oColFields.field[nField]));

		return {
			"field": aResult
		}
	};
	WriterToJSON.prototype.SerColItems = function(oColItems) // CT_colItems
	{
		if (!oColItems)
			return oColItems;

		var aResult = [];
		for (var nI = 0; nI < oColItems.i.length; nI++)
			aResult.push(this.SerI(oColItems.i[nI]));

		return {
			"i": aResult
		}
	};
	WriterToJSON.prototype.SerPageFields = function(oPageFields) // CT_PageFields
	{
		if (!oPageFields)
			return oPageFields;

		var aResult = [];
		for (var nField = 0; nField < oPageFields.pageField.length; nField++)
			aResult.push(this.SerPageField(oPageFields.pageField[nField]));
		
		return {
			"pageField": aResult
		}
	};
	WriterToJSON.prototype.SerPageField = function(oPageField) // CT_PageField
	{
		return {
			"fld":    oPageField.fld != null ? oPageField.fld : undefined,
			"item":   oPageField.item != null ? oPageField.item : undefined,
			"hier":   oPageField.hier != null ? oPageField.hier : undefined,
			"name":   oPageField.name != null ? oPageField.name : undefined,
			"cap":    oPageField.cap != null ? oPageField.cap : undefined,
			"extLst": oPageField.extLst != null ? this.SerExtensionList(oPageField.extLst) : undefined
		}
	};
	WriterToJSON.prototype.SerDataFields = function(oDataFields) // CT_DataFields
	{
		if (!oDataFields)
			return oDataFields;

		var aResult = [];
		for (var nField = 0; nField < oDataFields.dataField.length; nField++)
			aResult.push(this.SerDataField(oDataFields.dataField[nField]));

		return {
			"dataField": aResult
		}
	};
	WriterToJSON.prototype.SerDataField = function(oDataField) // CT_DataField
	{
		return {
			"name":       oDataField.name != null ? oDataField.name : undefined,
			"fld":        oDataField.fld != null ? oDataField.fld : undefined,
			"subtotal":   oDataField.subtotal != null ? ToXml_ST_DataConsolidateFunction(oDataField.subtotal) : undefined,
			"showDataAs": oDataField.showDataAs != null ? ToXml_ST_ShowDataAs(oDataField.showDataAs) : undefined,
			"baseField":  oDataField.baseField != null ? oDataField.baseField : undefined,
			"baseItem":   oDataField.baseItem != null ? oDataField.baseItem : undefined,
			"numFmtId":   oDataField.num != null ? this.stylesForWrite.getNumIdByFormat(oDataField.num) : undefined, // тут, отсюда
			"extLst":     oDataField.extLst != null ? this.SerExtensionList(oDataField.extLst) : undefined
		}
	};
	WriterToJSON.prototype.SerFormats = function(oFormats) // CT_Formats
	{
		if (!oFormats)
			return oFormats;

		var aResult = [];
		for (var nFmt = 0; nFmt < oFormats.format.length; nFmt++)
			aResult.push(this.SerFormat(oFormats.format[nFmt]));

		return {
			"format": aResult
		}
	};
	WriterToJSON.prototype.SerFormat = function(oFormat) // CT_Format
	{
		return { 
			"action":    ToXml_ST_FormatAction(oFormat.action),
			//dxfId:     oFormat.dxfId
			"pivotArea": this.SerPivotArea(oFormat.pivotArea),
			"extLst":    this.SerExtensionList(oFormat.extLst)
		}
	};
	WriterToJSON.prototype.SerConditionalFormats = function(oCondFormats) // CT_ConditionalFormats
	{
		if (!oCondFormats)
			return oCondFormats;

		var aResult = [];
		for (var nFmt = 0; nFmt < oCondFormats.conditionalFormat.length; nFmt++)
			aResult.push(this.SerConditionalFormat(oCondFormats.conditionalFormat[nFmt]));

		return {
			"conditionalFormat": aResult
		}
	};
	WriterToJSON.prototype.SerConditionalFormat = function(oFmt) // CT_ConditionalFormat
	{
		return {
			"scope":      oFmt.scope != null ? ToXml_ST_Scope(oFmt.scope) : undefined,
			"type":       oFmt.type != null ? ToXml_ST_Type(oFmt.type) : undefined,
			"priority":   oFmt.priority != null ? oFmt.priority : undefined,
			"pivotAreas": oFmt.pivotAreas != null ? this.SerPivotAreas(oFmt.pivotAreas) : undefined,
			"extLst":     oFmt.extLst != null ? this.SerExtensionList(oFmt.extLst) : undefined
		}
	};
	WriterToJSON.prototype.SerPivotAreas = function(oPivotAreas) // CT_PivotAreas
	{
		if (!oPivotAreas)
			return oPivotAreas;

		var aResult = 0;
		for (var nArea = 0; nArea < oPivotAreas.pivotArea.length; nArea++)
			aResult.push(this.SerPivotArea(oPivotAreas.pivotArea[nArea]));

		return {
			"pivotArea": aResult
		}
	};
	WriterToJSON.prototype.SerChartFormats = function(oChartFormats) // CT_ChartFormats
	{
		if (!oChartFormats)
			return oChartFormats;

		var aResult = [];
		for (var nFmt = 0; nFmt < oChartFormats.chartFormat.length; nFmt++)
			aResult.push(this.SerChartFormat(oChartFormats.chartFormat[nFmt]));

		return {
			"chartFormat": aResult
		}
	};
	WriterToJSON.prototype.SerChartFormat = function(oChartFormat) // CT_ChartFormat
	{
		return {
			"chart":     oChartFormat.chart != null ? oChartFormat.chart : undefined,
			"format":    oChartFormat.format != null ? oChartFormat.format : undefined,
			"series":    oChartFormat.series != null ? oChartFormat.series : undefined,
			"pivotArea": oChartFormat.pivotArea != null ? this.SerPivotArea(oChartFormat.pivotArea) : undefined
		}
	};
	WriterToJSON.prototype.SerPivotHierarchies = function(oPivotHier) // CT_PivotHierarchies
	{
		if (!oPivotHier)
			return oPivotHier;

		var aResult = [];
		for (var nPivot = 0; nPivot < oPivotHier.pivotHierarchy.length; nPivot++)
			aResult.push(this.SerPivotHierarchy(oPivotHier.pivotHierarchy[nPivot]));
		
		return {
			"pivotHierarchy": aResult
		}
	};
	WriterToJSON.prototype.SerPivotHierarchy = function(oPivot) // CT_PivotHierarchy
	{
		return {
			"outline":                      oPivot.outline != null ? oPivot.outline : undefined,
			"multipleItemSelectionAllowed": oPivot.multipleItemSelectionAllowed != null ? oPivot.multipleItemSelectionAllowed : undefined,
			"subtotalTop":                  oPivot.subtotalTop != null ? oPivot.subtotalTop : undefined,
			"showInFieldList":              oPivot.showInFieldList != null ? oPivot.showInFieldList : undefined,
			"dragToRow":                    oPivot.dragToRow != null ? oPivot.dragToRow : undefined,
			"dragToCol":                    oPivot.dragToCol != null ? oPivot.dragToCol : undefined,
			"dragToPage":                   oPivot.dragToPage != null ? oPivot.dragToPage : undefined,
			"dragToData":                   oPivot.dragToData != null ? oPivot.dragToData : undefined,
			"dragOff":                      oPivot.dragOff != null ? oPivot.dragOff : undefined,
			"includeNewItemsInFilter":      oPivot.includeNewItemsInFilter != null ? oPivot.includeNewItemsInFilter : undefined,
			"caption":                      oPivot.caption != null ? oPivot.caption : undefined,
			"mps":                          oPivot.mps != null ? this.SerMemberProps(oPivot.mps) : undefined,
			"members":                      oPivot.members != null ? this.SerMembers(oPivot.members) : undefined,
			"extLst":                       oPivot.extLst != null ? this.SerExtensionList(oPivot.extLst) : undefined
		}
	};
	WriterToJSON.prototype.SerMemberProps = function(oProps) // CT_MemberProperties
	{
		if (!oProps)
			return oProps;

		var aResult = [];
		for (var nMp = 0; nMp < oProps.mp.length; nMp++)
			aResult.push(this.SerMemberProp(oProps.mp[nMp]));
		
		return {
			"mp": aResult
		}
	};
	WriterToJSON.prototype.SerMemberProp = function(oProp) // CT_MemberProperty
	{
		return {
			"name":          oProp.name != null ? oProp.name : undefined,
			"showCell":      oProp.showCell != null ? oProp.showCell : undefined,
			"showTip":       oProp.showTip != null ? oProp.showTip : undefined,
			"showAsCaption": oProp.showAsCaption != null ? oProp.showAsCaption : undefined,
			"nameLen":       oProp.nameLen != null ? oProp.nameLen : undefined,
			"pPos":          oProp.pPos != null ? oProp.pPos : undefined,
			"pLen":          oProp.pLen != null ? oProp.pLen : undefined,
			"level":         oProp.level != null ? oProp.level : undefined,
			"field":         oProp.field != null ? oProp.field : undefined
		}
	};
	WriterToJSON.prototype.SerMembers = function(oMembers) // CT_Members
	{
		if (!oMembers)
			return oMembers;

		var aResult = [];
		for (var nMember = 0; nMember < oMembers.member.length; nMember++)
			aResult.push(this.SerMember(oMembers.member[nMember]));
		
		return {
			"level":  oMembers.level != null ? oMembers.level : undefined,
			"member": aResult
		}
	};
	WriterToJSON.prototype.SerMember = function(oMember) // CT_Member
	{
		return {
			"name": oMember.name != null ? oMember.name : undefined
		}
	};
	WriterToJSON.prototype.SerPivotTableStyleInfo = function(oInfo) // CT_PivotTableStyle
	{
		if (!oInfo)
			return oInfo;

		return {
			"name":           oInfo.name != null ?oInfo.name : undefined,
			"showRowHeaders": oInfo.showRowHeaders != null ?oInfo.showRowHeaders : undefined,
			"showColHeaders": oInfo.showColHeaders != null ?oInfo.showColHeaders : undefined,
			"showRowStripes": oInfo.showRowStripes != null ?oInfo.showRowStripes : undefined,
			"showColStripes": oInfo.showColStripes != null ?oInfo.showColStripes : undefined,
			"showLastColumn": oInfo.showLastColumn != null ?oInfo.showLastColumn : undefined
		}
	};
	WriterToJSON.prototype.SerPivotFilters = function(oFilters) // CT_PivotFilters
	{
		if (!oFilters)
			return oFilters;

		var aFilters = [];
		for (var nFilter = 0; nFilter < oFilters.filter.length; nFilter++)
			aFilters.push(this.SerPivotFilter(oFilters.filter[nFilter]));

		return {
			"filter": aFilters
		}
	};
	WriterToJSON.prototype.SerPivotFilter = function(oFilter) // CT_PivotFilter
	{
		return {
			"fld":          oFilter.fld != null ? oFilter.fld : undefined,
			"mpFld":        oFilter.mpFld != null ? oFilter.mpFld : undefined,
			"type":         oFilter.type != null ? ToXml_ST_PivotFilterType(oFilter.type) : undefined,
			"evalOrder":    oFilter.evalOrder != null ? oFilter.evalOrder : undefined,
			//"id":           oFilter.id != null ? oFilter.id : undefined,
			"iMeasureHier": oFilter.iMeasureHier != null ? oFilter.iMeasureHier : undefined,
			"iMeasureFld":  oFilter.iMeasureFld != null ? oFilter.iMeasureFld : undefined,
			"name":         oFilter.name != null ? oFilter.name : undefined,
			"description":  oFilter.description != null ? oFilter.description : undefined,
			"stringValue1": oFilter.stringValue1 != null ? oFilter.stringValue1 : undefined,
			"stringValue2": oFilter.stringValue2 != null ? oFilter.stringValue2 : undefined,
			"autoFilter":   oFilter.autoFilter != null ? this.SerAutoFilter(oFilter.autoFilter) : undefined,
			"extLst":       oFilter.extLst != null ? this.SerExtensionList(oFilter.extLst) : undefined
		}
	};
	WriterToJSON.prototype.SerRowHierarchiesUsage = function(oUsages) // CT_RowHierarchiesUsage
	{
		if (!oUsages)
			return oUsages;

		var aUsages = [];
		for (var nUsage = 0; nUsage < oUsages.rowHierarchyUsage.length; nUsage++)
			aUsages.push(this.SerHierarchyUsage(oUsages.rowHierarchyUsage[nUsage]));

		return {
			"rowHierarchyUsage": aUsages
		}
	};
	WriterToJSON.prototype.SerHierarchyUsage = function(oUsage) // CT_HierarchyUsage
	{
		if (!oUsage)
			return oUsage;

		return {
			"hierarchyUsage": oUsage.hierarchyUsage != null ? oUsage.hierarchyUsage : undefined
		}
	};
	WriterToJSON.prototype.SerColHierarchiesUsage = function(oUsages) // CT_ColHierarchiesUsage
	{
		if (!oUsages)
			return oUsages;

		var aUsages = [];
		for (var nUsage = 0; nUsage < oUsages.colHierarchyUsage.length; nUsage++)
			aUsages.push(this.SerHierarchyUsage(oUsages.colHierarchyUsage[nUsage]));

		return {
			"colHierarchyUsage": aUsages
		}
	};
	WriterToJSON.prototype.SerSlicers = function(aSlicers) // CT_slicers (такой объект создается на чтении, но по факту просто массив)
	{
		var aResult = [];
		for (var nSlicer = 0; nSlicer < aSlicers.length; nSlicer++)
			aResult.push(this.SerSlicer(aSlicers[nSlicer]));

		return aResult;
	};
	WriterToJSON.prototype.SerSlicer = function(oSlicer) // CT_slicer
	{
		return {
			"name":            oSlicer.name != null ? oSlicer.name : undefined,
			"uid":             oSlicer.uid != null ? oSlicer.uid : undefined,
			"cacheDefinition": oSlicer.cacheDefinition != null ? oSlicer.cacheDefinition.name : undefined,
			"caption":         oSlicer.caption != null ? oSlicer.caption : undefined,
			"startItem":       oSlicer.startItem != null ? oSlicer.startItem : undefined,
			"columnCount":     oSlicer.columnCount != null ? oSlicer.columnCount : undefined,
			"showCaption":     oSlicer.showCaption != null ? oSlicer.showCaption : undefined,
			"level":           oSlicer.level != null ? oSlicer.level : undefined,
			"style":           oSlicer.style != null ? oSlicer.style : undefined,
			"lockedPosition":  oSlicer.lockedPosition != null ? oSlicer.lockedPosition : undefined,
			"rowHeight":       oSlicer.rowHeight != null ? oSlicer.rowHeight  : undefined
		}
	};
	WriterToJSON.prototype.SerSlicerCacheDefinition = function(oChacheDef) // CT_slicerCacheDefinition
	{
		if (!oChacheDef)
			return oChacheDef;

		return {
			"name":                           oChacheDef.name != null ? oChacheDef.name : undefined,
			"uid":                            oChacheDef.uid != null ? oChacheDef.uid : undefined,
			"sourceName":                     oChacheDef.sourceName != null ? oChacheDef.sourceName : undefined,
			"pivotTables":                    oChacheDef.pivotTables != null ? this.SerSlicerCachePivotTables(oChacheDef.pivotTables) : undefined,
			"data":                           oChacheDef.data != null ? this.SerSlicerCacheData(oChacheDef.data) : undefined,
			"tableSlicerCache":               oChacheDef.tableSlicerCache != null ? this.SerTableSlicerCache(oChacheDef.tableSlicerCache) : undefined,
			"slicerCacheHideItemsWithNoData": oChacheDef.slicerCacheHideItemsWithNoData != null ? this.SerSlicerCacheHideNoData(oChacheDef.slicerCacheHideItemsWithNoData) : undefined
		}
	};
	WriterToJSON.prototype.SerSlicerCachePivotTables = function(aTables) // CT_slicerCachePivotTable[]
	{
		var aResult = [];
		for (var nTable = 0; nTable < aTables.length; nTable++)
			aResult.push(this.SerSlicerCachePivotTable(aTables[nTable]));

		return aResult;
	};
	WriterToJSON.prototype.SerSlicerCachePivotTable = function(oTable) // CT_slicerCachePivotTable
	{
		return {
			"sheetId":   oTable.sheetId,
			"tabIdOpen": oTable.tabIdOpen, // to do возможно не нужно
			"name":      oTable.name
		}
	};
	WriterToJSON.prototype.SerTableSlicerCache = function(oCache) // CT_tableSlicerCache
	{
		if (!oCache)
			return oCache;
	
		return {
			"tableId":        oCache.tableId != null ? oCache.tableId : undefined,
			"tableIdOpen":    oCache.tableIdOpen != null ? oCache.tableIdOpen : undefined, // возможно не нужно
			"column":         oCache.column != null ? oCache.column : undefined,
			"columnOpen":     oCache.columnOpen != null ? oCache.columnOpen : undefined, // возможно не нужно
			"sortOrder":      oCache.sortOrder != null ? ToXML_ST_tabularSlicerCacheSortOrder(oCache.sortOrder) : undefined,
			"customListSort": oCache.customListSort != null ? oCache.customListSort : undefined,
			"crossFilter":    oCache.crossFilter != null ? ToXML_ST_slicerCacheCrossFilter(oCache.crossFilter) : undefined
		}
	};
	WriterToJSON.prototype.SerSlicerCacheHideNoData = function(oCache) // CT_slicerCacheHideNoData
	{
		if (!oCache)
			return oCache;

		var aResult = [];
		for (var nItem = 0; nItem < oCache.slicerCacheOlapLevelName.length; nItem++)
			aResult.push(this.SerSlicerCacheOlapLevelName(oCache.slicerCacheOlapLevelName[nItem]));

		return {
			"count": oCache.count,
			"slicerCacheOlapLevelName": aResult
		}
	};
	WriterToJSON.prototype.SerSlicerCacheOlapLevelName = function(oLvlName) // CT_slicerCacheOlapLevelName
	{
		return {
			"uniqueName": oLvlName.uniqueName != null ? oLvlName.uniqueName : undefined,
			"count":      oLvlName.count != null ? oLvlName.count : undefined
		}
	};
	WriterToJSON.prototype.SerSlicerCacheData = function(oData) // CT_slicerCacheData
	{
		if (!oData)
			return oData;

		return {
			"olap":    oData.olap != null ? this.SerOlapSlicerCache(oData.olap) : undefined,
			"tabular": oData.tabular != null ? this.SerTabularSlicerCache(oData.tabular) : undefined
		}
	};
	WriterToJSON.prototype.SerOlapSlicerCache = function(oOlap) // CT_olapSlicerCache
	{
		if (!oOlap)
			return oOlap;

		return {
			"pivotCacheId":         oOlap.getPivotCacheId(),
			"levels":               this.SerOlapSlicerCacheLevelsData(oOlap.levels),
			"selections":           this.SerOlapSlicerCacheSelections(oOlap.selections)
		}
	};
	WriterToJSON.prototype.SerOlapSlicerCacheLevelsData = function(aLevels) // CT_olapSlicerCacheLevelData[]
	{
		var aResult = [];
		for (var nLvl = 0; nLvl < aLevels.length; nLvl++)
			aResult.push(this.SerOlapSlicerCacheLevelData(aLevels[nLvl]));

		return aResult;
	};
	WriterToJSON.prototype.SerOlapSlicerCacheLevelData = function(oLevel) // CT_olapSlicerCacheLevelData
	{
		return {
			"uniqueName":    oLevel.uniqueName != null ? oLevel.uniqueName : undefined,
			"sourceCaption": oLevel.sourceCaption != null ? oLevel.sourceCaption : undefined,
			"count":         oLevel.count != null ? oLevel.count : undefined,
			"sortOrder":     oLevel.sortOrder != null ? ToXML_ST_olapSlicerCacheSortOrder(oLevel.sortOrder) : undefined,
			"crossFilter":   oLevel.crossFilter != null ? ToXML_ST_slicerCacheCrossFilter(oLevel.crossFilter) : undefined,
			"ranges":        oLevel.ranges != null ? this.SerOlapSlicerCacheRanges(oLevel.ranges) : undefined
		}
	};
	WriterToJSON.prototype.SerOlapSlicerCacheRanges = function(aRanges) // CT_olapSlicerCacheRange[]
	{
		var aResult = [];
		for (var nRange = 0; nRange < aRanges.length; nRange++)
			aResult.push(this.SerOlapSlicerCacheRange(aRanges[nRange]));

		return aResult;
	};
	WriterToJSON.prototype.SerOlapSlicerCacheRange = function(oRange) // CT_olapSlicerCacheRange
	{
		var aItems = [];
		for (var nItem = 0; nItem < oRange.i.length; nItem++)
			aItems.push(this.SerOlapSlicerCacheItem(oRange.i[nItem]));

		return {
			"startItem": oRange.startItem != null ? oRange.startItem : undefined,
			"i":         aItems
		}
	};
	WriterToJSON.prototype.SerOlapSlicerCacheItem = function(oItem) // CT_olapSlicerCacheItem
	{
		return {
			"n":  oItem.n != null ? oItem.n : undefined,
			"c":  oItem.c != null ? oItem.c : undefined,
			"nd": oItem.nd != null ? oItem.nd : undefined,
			"p":  oItem.p != null ? this.SerOlapSlicerCacheItemParents(oItem.p) : undefined
		}
	};
	WriterToJSON.prototype.SerOlapSlicerCacheItemParents = function(aParents)
	{
		var aResult = [];
		for (var nIndex = 0; nIndex < aParents.length; nIndex++)
			aResult.push(this.SerOlapSlicerCacheItemParent(aParents[nIndex]));

		return aResult;
	};
	WriterToJSON.prototype.SerOlapSlicerCacheItemParent = function(oParent) // CT_olapSlicerCacheItemParent
	{
		return {
			"n": oParent.n != null ? oParent.n : undefined
		}
	};
	WriterToJSON.prototype.SerOlapSlicerCacheSelections = function(aSelections) // CT_olapSlicerCacheSelection[]
	{
		var aResult = [];
		for (var nItem = 0; nItem < aSelections.length; nItem++)
			aResult.push(this.SerOlapSlicerCacheSelection(aSelections[nItem]));

		return aResult;
	};
	WriterToJSON.prototype.SerOlapSlicerCacheSelection = function(oSelection) // CT_olapSlicerCacheSelection
	{
		return {
			"n": oSelection.n != null ? oSelection.n : undefined,
			"p": oSelection.p != null ? this.SerOlapSlicerCacheItemParents(oSelection.p) : undefined
		}
	};
	WriterToJSON.prototype.SerPivotCacheDefinition = function(oCache) // CT_PivotCacheDefinition
	{
		if (!oCache)
			return oCache;

		return {
			// attributes
			"id":                    oCache.id,
			"invalid":               oCache.invalid !== false ? oCache.invalid : undefined,
			"saveData":              oCache.saveData !== true ? oCache.saveData : undefined,
			"refreshOnLoad":         oCache.refreshOnLoad !== false ? oCache.refreshOnLoad : undefined,
			"optimizeMemory":        oCache.optimizeMemory !== false ? oCache.optimizeMemory : undefined,
			"enableRefresh":         oCache.enableRefresh !== true ? oCache.enableRefresh : undefined,
			"refreshedBy":           oCache.refreshedBy != null ? oCache.refreshedBy : undefined,
			"refreshedDate":         oCache.refreshedDate != null ? oCache.refreshedDate : undefined,
			"backgroundQuery":       oCache.backgroundQuery !== false ? oCache.backgroundQuery : undefined,
			"missingItemsLimit":     oCache.missingItemsLimit != null ? oCache.missingItemsLimit : undefined,
			"createdVersion":        oCache.createdVersion !== 0 ? oCache.createdVersion : undefined,
			"refreshedVersion":      oCache.refreshedVersion !== 0 ? oCache.refreshedVersion : undefined,
			"minRefreshableVersion": oCache.minRefreshableVersion !== 0 ? oCache.minRefreshableVersion : undefined,
			"cacheRecords":          oCache.cacheRecords != null ? this.SerPivotCacheRecords(oCache.cacheRecords) : undefined,
			"upgradeOnRefresh":      oCache.upgradeOnRefresh !== false ? oCache.upgradeOnRefresh : undefined,
			"supportSubquery":       oCache.supportSubquery !== false ? oCache.supportSubquery : undefined,
			"supportAdvancedDrill":  oCache.supportAdvancedDrill !== false ? oCache.supportAdvancedDrill : undefined,
			// members
			"cacheSource":           oCache.cacheSource !== null ? this.SerCacheSource(oCache.cacheSource) : undefined,
			"cacheFields":           oCache.cacheFields !== null ? this.SerCacheFields(oCache.cacheFields) : undefined,
			"cacheHierarchies":      oCache.cacheHierarchies !== null ? this.SerCacheHierarchies(oCache.cacheHierarchies) : undefined,
			"kpis":                  oCache.kpis !== null ? this.SerPCDKPIs(oCache.kpis) : undefined,
			"tupleCache":            oCache.tupleCache !== null ? this.SerTupleCache(oCache.tupleCache) : undefined,
			"calculatedItems":       oCache.calculatedItems !== null ? this.SerCalculatedItems(oCache.calculatedItems) : undefined,
			"calculatedMembers":     oCache.calculatedMembers !== null ? this.SerCalculatedMembers(oCache.calculatedMembers) : undefined,
			"dimensions":            oCache.dimensions !== null ? this.SerDimensions(oCache.dimensions) : undefined,
			"measureGroups":         oCache.measureGroups !== null ? this.SerMeasureGroups(oCache.measureGroups) : undefined,
			"maps":                  oCache.maps !== null ? this.SerMeasureDimensionMaps(oCache.maps) : undefined,
			// ext
			"pivotCacheDefinitionX14": oCache.pivotCacheDefinitionX14 !== null ? this.SerPivotCacheDefinitionX14(oCache.pivotCacheDefinitionX14) : undefined // to do (возожно стоит записать как extLst)
		}
	};
	WriterToJSON.prototype.SerPivotCacheRecords = function(oRecords) // CT_PivotCacheRecords
	{
		var count = oRecords.getRowsCount();

		let oResult = {
			"records": [],
			"extLst":  oRecords.extLst != null ? this.SerExtensionList(oRecords.extLst) : undefined,
			"count": count
		}

		for (let i = 0; i < count; i++) {
			let aRow = [];
			for (let j = 0; j < oRecords._cols.length; j++)
				aRow = aRow.concat(this.SerPivotRecords(oRecords._cols[j], i));
			
			oResult["records"].push(aRow);
		}

		return oResult;
	};
	WriterToJSON.prototype.SerPivotCaches = function(oCaches)
	{
		var oResult = {};

		for (var key in oCaches)
			oResult[key] = this.SerPivotCacheDefinition(oCaches[key].cache);

		return oResult;
	};
	WriterToJSON.prototype.SerSlicerCaches = function(oCaches)
	{
		var oResult = {};

		for (var key in oCaches)
			oResult[key] = this.SerSlicerCacheDefinition(oCaches[key]);

		return oResult;
	};
	WriterToJSON.prototype.SerCacheSource = function(oSource) // CT_CacheSource
	{
		if (!oSource)
			return oSource;

		return {
			"type":            oSource.type != null ? ToXml_ST_SourceType(oSource.type) : undefined,
			"connectionId":    oSource.connectionId !== 0 ? oSource.connectionId : undefined,
			"consolidation":   oSource.consolidation != null ? this.SerConsolidation(oSource.consolidation) : undefined,
			"extLst":          oSource.extLst != null ? this.SerExtensionList(oSource.extLst) : undefined,
			"worksheetSource": oSource.worksheetSource != null ? this.SerWorksheetSource(oSource.worksheetSource) : undefined
		}
	};
	WriterToJSON.prototype.SerConsolidation = function(oConsolidation) // CT_Consolidation
	{
		if (!oConsolidation)
			return oConsolidation;

		return {
			"autoPage":  oConsolidation.autoPage != null ? oConsolidation.autoPage : undefined,
			"pages":     oConsolidation.pages != null ? this.SerPages(oConsolidation.pages) : undefined,
			"rangeSets": oConsolidation.rangeSets != null ? this.SerRangeSets(oConsolidation.rangeSets) : undefined
		}
	};
	WriterToJSON.prototype.SerPages = function(oPages) // CT_Pages
	{
		if (!oPages)
			return oPages;

		var aPages = [];
		for (var nElm = 0; nElm < oPages.page.length; nElm++)
			aPages.push(this.SerPage(oPages.page[nElm]));

		return {
			"count": aPages.length > 0 ? aPages.length : undefined,
			"page":  aPages
		}
	};
	WriterToJSON.prototype.SerPage = function(oPage) // CT_PCDSCPage
	{
		var aPageItems = [];
		for (var nElm = 0; nElm < oPage.pageItem.length; nElm++)
			aPageItems.push(this.SerPageItem(oPage.pageItem[nElm]));

		return {
			"count":    oPage.pageItem.length > 0 ? oPage.pageItem.length : undefined,
			"pageItem": aPageItems
		}
	};
	WriterToJSON.prototype.SerPageItem = function(oItem) // CT_PageItem
	{
		return {
			"name": oItem.name != null ? oItem.name : undefined
		}
	};
	WriterToJSON.prototype.SerRangeSets = function(oSets) // CT_RangeSets
	{
		if (!oSets)
			return oSets;

		var aSets = [];
		for (var nElm = 0; nElm < oSets.rangeSet.length; nElm++)
			aSets.push(this.SerRangeSet(oSets.rangeSet[nElm]));

		return {
			"count":    aSets.length > 0 ? aSets.length : undefined,
			"rangeSet": aSets
		}
	};
	WriterToJSON.prototype.SerRangeSet = function(oSet) // CT_RangeSet
	{
		return {
			"i1":    oSet.i1 != null ? oSet.i1 : undefined,
			"i2":    oSet.i2 != null ? oSet.i2 : undefined,
			"i3":    oSet.i3 != null ? oSet.i3 : undefined,
			"i4":    oSet.i4 != null ? oSet.i4 : undefined,
			"ref":   oSet.ref != null ? oSet.ref : undefined,
			"name":  oSet.name != null ? oSet.name : undefined,
			"sheet": oSet.sheet != null ? oSet.sheet : undefined
		}
	};
	WriterToJSON.prototype.SerWorksheetSource = function(oWSource) // CT_WorksheetSource
	{
		if (!oWSource)
			return oWSource;

		return {
			"ref":   oWSource.ref != null ? oWSource.ref : undefined,
			"name":  oWSource.name != null ? oWSource.name : undefined,
			"sheet": oWSource.sheet != null ? oWSource.sheet : undefined
			//"r:id":  oSource.id != null ? oSource.id : undefined // мб не надо
		}
	};
	WriterToJSON.prototype.SerCacheFields = function(oFields) // CT_CacheFields
	{
		var aFields = [];
		for (var nElm = 0; nElm < oFields.cacheField.length; nElm++)
			aFields.push(this.SerCacheField(oFields.cacheField[nElm]));

		return {
			"count":      oFields.cacheField.length > 0 ? oFields.cacheField.length : undefined,
			"cacheField": aFields
		}
	};
	WriterToJSON.prototype.SerCacheField = function(oField) // CT_CacheField
	{
		var aMpMap = [];
		for (var nElm = 0; nElm < oField.mpMap.length; nElm++)
			aMpMap.push(this.SerX(oField.mpMap[nElm]));

		return {
			"name":                oField.name != null ? oField.name : undefined,
			"caption":             oField.caption != null ? oField.caption : undefined,
			"propertyName":        oField.propertyName != null ? oField.propertyName : undefined,
			"serverField":         oField.serverField !== false ? oField.serverField : undefined,
			"uniqueList":          oField.uniqueList !== true ? oField.uniqueList : undefined,
			"numFmtId":            oField.num != null ? this.SerNumFmtExcell(oField.num, undefined, true) : undefined,
			"formula":             oField.formula != null ? oField.formula : undefined,
			"sqlType":             oField.sqlType != 0 ? oField.sqlType : undefined,
			"hierarchy":           oField.hierarchy != 0 ? oField.hierarchy : undefined,
			"level":               oField.level != 0 ? oField.level : undefined,
			"databaseField":       oField.databaseField !== true ? oField.databaseField : undefined,
			"mappingCount":        oField.mappingCount != null ? oField.mappingCount : undefined,
			"memberPropertyField": oField.memberPropertyField != false ? oField.memberPropertyField : undefined,
			"sharedItems":         oField.sharedItems != null ? this.SerSharedItems(oField.sharedItems) : undefined,
			"fieldGroup":          oField.fieldGroup != null ? this.SerFieldGroup(oField.fieldGroup) : undefined,
			"mpMap":               aMpMap,
			"extLst":              oField.extLst != null ? this.SerExtensionList(oField.extLst) : undefined
		}
	};
	WriterToJSON.prototype.SerSharedItems = function(oItems) // CT_SharedItems
	{
		if (!oItems)
			return oItems;

		var nCount = oItems.Items.getSize();

		return {
			// attributes
			"containsSemiMixedTypes": oItems.containsSemiMixedTypes !== true ? oItems.containsSemiMixedTypes : undefined,
			"containsNonDate":        oItems.containsNonDate !== true ? oItems.containsNonDate : undefined,
			"containsDate":           oItems.containsDate !== false ? oItems.containsDate : undefined,
			"containsString":         oItems.containsString !== true ? oItems.containsString : undefined,
			"containsBlank":          oItems.containsBlank !== false ? oItems.containsBlank : undefined,
			"containsMixedTypes":     oItems.containsMixedTypes !== false ? oItems.containsMixedTypes : undefined,
			"containsNumber":         oItems.containsNumber !== false ? oItems.containsNumber : undefined,
			"containsInteger":        oItems.containsInteger !== false ? oItems.containsInteger : undefined,
			"minValue":               oItems.minValue != null ? oItems.minValue : undefined,
			"maxValue":               oItems.maxValue != null ? oItems.maxValue : undefined,
			"minDate":                oItems.minDate != null ? oItems.minDate.toISOString().slice(0, 19) : undefined,
			"maxDate":                oItems.maxDate != null ? oItems.maxDate.toISOString().slice(0, 19) : undefined,
			"count":                  nCount > 0 ? nCount : undefined,
			"longText":               oItems.longText !== false ? oItems.longText : undefined,
			// members
			"items":                  this.SerPivotRecords(oItems.Items)
		}
	};
	WriterToJSON.prototype.SerPivotRecords = function(oRecords, nOptIndex) // PivotRecords
	{
		if (!oRecords)
			return oRecords;

		var aRecords = [];
		var oTempData, sTempType;
		if (nOptIndex !== undefined)
		{
			oTempData = oRecords.get(nOptIndex);
			sTempType = ToXML_PivotRecType(oTempData.type);
			aRecords.push(this.SerPivotRecord(sTempType, oTempData.val, oTempData.addition));
		}
		else
		{
			for (var nElm = 0; nElm < oRecords.size; nElm++)
			{
				oTempData = oRecords.get(nElm);
				sTempType = ToXML_PivotRecType(oTempData.type);
				aRecords.push(this.SerPivotRecord(sTempType, oTempData.val, oTempData.addition));
			}
		}
		
		return aRecords;
	};
	WriterToJSON.prototype.SerPivotRecord = function(sType, val, oAddition)
	{
		var oResult = {};

		if (val != null)
			oResult["v"] = val;

		oResult["type"] = sType;

		if (oAddition)
		{
			if (oAddition.u != null)
				oResult["u"] = oAddition.u;
			if (oAddition.f != null)
				oResult["f"] = oAddition.f;
			if (oAddition.c != null)
				oResult["c"] = oAddition.c;
			if (oAddition.cp != null)
				oResult["cp"] = oAddition.cp;
			if (oAddition.in != null)
				oResult["in"] = oAddition.in;
			if (oAddition.bc != null)
				oResult["bc"] = oAddition.bc;
			if (oAddition.fc != null)
				oResult["fc"] = oAddition.fc;
			if (oAddition.i != null)
				oResult["i"] = oAddition.i;
			if (oAddition.un != null)
				oResult["un"] = oAddition.un;
			if (oAddition.st != null)
				oResult["st"] = oAddition.st;
			if (oAddition.b != null)
				oResult["b"] = oAddition.b;

			if (oAddition.tpls.length > 0)
			{
				oResult["tpls"] = [];
				for (var nElm = 0; nElm < oAddition.tpls.length; nElm++)
					oResult["tpls"].push(this.SerTuples(oAddition.tpls[nElm]));
			}
			if (oAddition.tpls.length > 0)
			{
				oResult["x"] = [];
				for (var nElm = 0; nElm < oAddition.x.length; nElm++)
					oResult["x"].push(this.SerX(oAddition.x[nElm]));
			}
		}

		return oResult;
	};
	WriterToJSON.prototype.SerFieldGroup = function(oGroup) // CT_FieldGroup
	{
		if (!oGroup)
			return oGroup;

		return {
			"par":        oGroup.par != null ? oGroup.par : undefined,
			"base":       oGroup.base != null ? oGroup.base : undefined,
			"rangePr":    oGroup.rangePr != null ? this.SerRangePr(oGroup.rangePr) : undefined,
			"discretePr": oGroup.discretePr != null ? this.SerDiscretePr(oGroup.discretePr) : undefined,
			"groupItems": oGroup.groupItems != null ? this.SerSharedItems(oGroup.groupItems) : undefined
		}
	};
	WriterToJSON.prototype.SerRangePr = function(oPr) // CT_RangePr
	{
		if (!oPr)
			return oPr;

		return {
			"autoStart":     oPr.autoStart !== true ? oPr.autoStart : undefined,
			"autoEnd":       oPr.autoEnd !== true ? oPr.autoEnd : undefined,
			"groupBy":       c_oAscGroupBy.Range !== oPr.groupBy ? ToXml_ST_GroupBy(oPr.groupBy) : undefined,
			"startNum":      oPr.startNum != null ? oPr.startNum : undefined,
			"endNum":        oPr.endNum != null ? oPr.endNum : undefined,
			"startDate":     oPr.startDate != null ? oPr.startDate.toISOString().slice(0, 19) : undefined,
			"endDate":       oPr.endDate != null ? oPr.endDate.toISOString().slice(0, 19) : undefined,
			"groupInterval": oPr.groupInterval !== -1 ? oPr.groupInterval : undefined,
		}
	};
	WriterToJSON.prototype.SerDiscretePr = function(oPr) // CT_DiscretePr
	{
		if (!oPr)
			return oPr;

		var aIndex = [];
		for (var nElm = 0; nElm < oPr.x.length; nElm++)
			aIndex.push(this.SerIndex(oPr.x[nElm]));

		return {
			"count": oPr.x.length > 0 ? oPr.x.length : undefined,
			"x":     aIndex
		}
	};
	WriterToJSON.prototype.SerTuples = function(oTuples) // CT_Tuples
	{
		var aTpl = [];
		for (var nElm = 0; nElm < oTuples.tpl.length; nElm++)
			aTpl.push(this.SerTuple(oTuples.tpl[nElm]));

		return {
			"tpl": aTpl,
			"c":   oTuples.c != null ? oTuples.c : undefined
		}
	};
	WriterToJSON.prototype.SerTuple = function(oTuple) // CT_Tuple
	{
		return {
			"fld":  oTuple.fld != null ? oTuple.fld : undefined,
			"hier": oTuple.hier != null ? oTuple.hier : undefined,
			"item": oTuple.item != null ? oTuple.item : undefined
		}
	};
	WriterToJSON.prototype.SerCacheHierarchies = function(oCacheHiers) // CT_CacheHierarchies
	{
		var aCaches = [];
		for (var nElm = 0; nElm < oCacheHiers.cacheHierarchy.length; nElm++)
			aCaches.push(this.SerCacheHierarchy(oCacheHiers.cacheHierarchy[nElm]));

		return {
			"count":          oCacheHiers.cacheHierarchy.length > 0 ? oCacheHiers.cacheHierarchy.length : undefined,
			"cacheHierarchy": aCaches
		}
	};
	WriterToJSON.prototype.SerCacheHierarchy = function(oCacheHier) // CT_CacheHierarchy
	{
		return {
			"uniqueName":              oCacheHier.uniqueName != null ? oCacheHier.uniqueName : undefined,
			"caption":                 oCacheHier.caption != null ? oCacheHier.caption : undefined,
			"measure":                 oCacheHier.measure !== false ? oCacheHier.measure : undefined,
			"set":                     oCacheHier.set !== false ? oCacheHier.set : undefined,
			"parentSet":               oCacheHier.parentSet != null ? oCacheHier.parentSet : undefined,
			"iconSet":                 oCacheHier.iconSet !== 0 ? oCacheHier.iconSet : undefined,
			"attribute":               oCacheHier.attribute !== false ? oCacheHier.attribute : undefined,
			"time":                    oCacheHier.time !== false ? oCacheHier.time : undefined,
			"keyAttribute":            oCacheHier.keyAttribute !== false ? oCacheHier.keyAttribute : undefined,
			"defaultMemberUniqueName": oCacheHier.defaultMemberUniqueName != null ? oCacheHier.defaultMemberUniqueName : undefined,
			"allUniqueName":           oCacheHier.allUniqueName != null ? oCacheHier.allUniqueName : undefined,
			"allCaption":              oCacheHier.allCaption != null ? oCacheHier.allCaption : undefined,
			"dimensionUniqueName":     oCacheHier.dimensionUniqueName != null ? oCacheHier.dimensionUniqueName : undefined,
			"displayFolder":           oCacheHier.displayFolder != null ? oCacheHier.displayFolder : undefined,
			"measureGroup":            oCacheHier.measureGroup != null ? oCacheHier.measureGroup : undefined,
			"measures":                oCacheHier.measures !== false ? oCacheHier.measures : undefined,
			"count":                   oCacheHier.count != null ? oCacheHier.count : undefined,
			"oneField":                oCacheHier.oneField !== false ? oCacheHier.oneField : undefined,
			"memberValueDatatype":     oCacheHier.memberValueDatatype !== null ? oCacheHier.memberValueDatatype : undefined,
			"unbalanced":              oCacheHier.unbalanced !== null ? oCacheHier.unbalanced : undefined,
			"unbalancedGroup":         oCacheHier.unbalancedGroup !== null ? oCacheHier.unbalancedGroup : undefined,
			"hidden":                  oCacheHier.hidden !== false ? oCacheHier.hidden : undefined,
			"fieldsUsage":             oCacheHier.fieldsUsage != null ? this.SerFieldsUsage(oCacheHier.fieldsUsage) : undefined,
			"groupLevels":             oCacheHier.groupLevels != null ? this.SerGroupLevels(oCacheHier.groupLevels) : undefined,
			"extLst":                  oCacheHier.extLst != null ? this.SerExtensionList(oCacheHier.extLst) : undefined

		}
	};
	WriterToJSON.prototype.SerFieldsUsage = function(oFields) // CT_FieldsUsage
	{
		if (!oFields)
			return oFields;

		var aUsages = [];
		for (var nElm = 0; nElm < oFields.fieldUsage.length; nElm++)
			aUsages.push(this.SerFieldUsage(oFields.fieldUsage[nElm]));

		return {
			"count":      oFields.fieldUsage.length > 0 ? oFields.fieldUsage.length : undefined,
			"fieldUsage": aUsages
		}
	};
	WriterToJSON.prototype.SerFieldUsage = function(oField) // CT_FieldUsage
	{
		return {
			"x": oField.x != null ? oField.x : undefined
		}
	};
	WriterToJSON.prototype.SerGroupLevels = function(oLevels) // CT_GroupLevels
	{
		if (!oLevels)
		return oLevels;

		var aLevel = [];
		for (var nElm = 0; nElm < oLevels.groupLevel.length; nElm++)
			aLevel.push(this.SerGroupLevel(oLevels.groupLevel[nElm]));

		return {
			"count":      oLevels.groupLevel.length > 0 ? oLevels.groupLevel.length : undefined,
			"groupLevel": aLevel
		}
	};
	WriterToJSON.prototype.SerGroupLevel = function(oLevel) // CT_GroupLevel
	{
		return {
			"uniqueName":   oLevel.uniqueName != null ? oLevel.uniqueName : undefined,
			"caption":      oLevel.caption != null ? oLevel.caption : undefined,
			"user":         oLevel.user !== false ? oLevel.user : undefined,
			"customRollUp": oLevel.customRollUp !== false ? oLevel.customRollUp : undefined,
			"groups":       oLevel.groups != null ? this.SerGroups(oLevel.groups) : undefined,
			"extLst":       oLevel.extLst != null ? this.SerExtensionList(oLevel.extLst) : undefined
		}
	};
	WriterToJSON.prototype.SerGroups = function(oGroups) // CT_Groups
	{
		if (!oGroups)
			return oGroups;

		var aGroup = [];
		for (var nElm = 0; nElm < oGroups.group.length; nElm++)
			aGroup.push(this.SerGroup(oGroups.group[nElm]));

		return {
			"count": oGroups.group.length > 0 ? oGroups.group.length : undefined,
			"group": aGroup
		}
	};
	WriterToJSON.prototype.SerGroup = function(oGroup) // CT_LevelGroup
	{
		return {
			"name":         oGroup.name != null ? oGroup.name : undefined,
			"uniqueName":   oGroup.uniqueName != null ? oGroup.uniqueName : undefined,
			"caption":      oGroup.caption != null ? oGroup.caption : undefined,
			"uniqueParent": oGroup.uniqueParent != null ? oGroup.uniqueParent : undefined,
			"id":           oGroup.id != null ? oGroup.id : undefined,
			"groupMembers": oGroup.groupMembers != null ? this.SerGroupMembers(oGroup.groupMembers) : undefined
		}
	};
	WriterToJSON.prototype.SerGroupMembers = function(oMembers) // CT_GroupMembers
	{
		if (!oMembers)
		return oMembers;

		var aMembers = [];
		for (var nElm = 0; nElm < oMembers.groupMember.length; nElm++)
			aMembers.push(this.SerGroupMember(oMembers.groupMember[nElm]));

		return {
			"count":       oMembers.groupMember.length > 0 ? oMembers.groupMember.length : undefined,
			"groupMember": aMembers
		}
	};
	WriterToJSON.prototype.SerGroupMember = function(oMember) // CT_GroupMember
	{
		return {
			"uniqueName": oMember.uniqueName != null ? oMember.uniqueName : undefined,
			"group":      oMember.group !== false ? oMember.group : undefined
		}
	};
	WriterToJSON.prototype.SerPCDKPIs = function(oKPIs) // CT_PCDKPIs
	{
		if (!oKPIs)
		return oKPIs;

		var aKpi = [];
		for (var nElm = 0; nElm < oKPIs.kpi.length; nElm++)
			aKpi.push(this.SerKPI(oKPIs.kpi[nElm]));

		return {
			"count": oKPIs.kpi.length > 0 ? oKPIs.kpi.length : undefined,
			"kpi":   aKpi
		}
	};
	WriterToJSON.prototype.SerKPI = function(oKPI) // CT_PCDKPI
	{
		return {
			// attributes 
			"uniqueName":    oKPI.uniqueName !== null ? oKPI.uniqueName : undefined,
			"caption":       oKPI.caption !== null ? oKPI.caption : undefined,
			"displayFolder": oKPI.displayFolder !== null ? oKPI.displayFolder : undefined,
			"measureGroup":  oKPI.measureGroup !== null ? oKPI.measureGroup : undefined,
			"parent":        oKPI.parent !== null ? oKPI.parent : undefined,
			"value":         oKPI.value !== null ? oKPI.value : undefined,
			"goal":          oKPI.goal !== null ? oKPI.goal : undefined,
			"status":        oKPI.status !== null ? oKPI.status : undefined,
			"trend":         oKPI.trend !== null ? oKPI.trend : undefined,
			"weight":        oKPI.weight !== null ? oKPI.weight : undefined,
			"time":          oKPI.time !== null ? oKPI.time : undefined
		}
	};
	WriterToJSON.prototype.SerTupleCache = function(oCache) // CT_TupleCache
	{
		if (!oCache)
			return oCache;

		return {
			"entries":       this.SerPCDSDTCEntries(oCache.entries),
			"sets":          this.SerSets(oCache.sets),
			"queryCache":    this.SerQueryCache(oCache.queryCache),
			"serverFormats": this.SerServerFormats(oCache.serverFormats),
			"extLst":        this.SerExtensionList(oCache.extLst)
		}
	};
	WriterToJSON.prototype.SerPCDSDTCEntries = function(oEntries) // CT_PCDSDTCEntries
	{
		if (!oEntries)
		return oEntries;

		var aItems = [];
		for (var nElm = 0; nElm < oEntries.Items.length; nElm++)
			aItems.push(this.SerPCDSDTCItem(oEntries.Items[nElm]));

		return {
			"count": oEntries.Items.length > 0 ? oEntries.Items.length : undefined,
			"items": aItems
		}
	};
	WriterToJSON.prototype.SerPCDSDTCItem = function(oItem)
	{
		var oResult;
		if (oItem instanceof Asc.CT_Error)
			oResult = this.SerPivotRecord(oItem, oItem.v, "e");
		else if (oItem instanceof Asc.CT_Missing)
			oResult = this.SerPivotRecord(oItem, oItem.v, "m");
		else if (oItem instanceof Asc.CT_Number)
			oResult = this.SerPivotRecord(oItem, oItem.v, "n");
		else if (oItem instanceof Asc.CT_String)
			oResult = this.SerPivotRecord(oItem, oItem.v, "s");

		return oResult;
	};
	WriterToJSON.prototype.SerSets = function(oSets) // CT_Sets
	{
		if (!oSets)
			return oSets;

		var aSet = [];
		for (var nElm = 0; nElm < oSets.set.length; nElm++)
			aSet.push(this.SerSet(oSets.set[nElm]));

		return {
			"count": oSets.set.length > 0 ? oSets.set.length : undefined,
			"set":   aSet
		}
	};
	WriterToJSON.prototype.SerSet = function(oSet) // CT_Set
	{
		var aTuples = [];
		for (var nElm = 0; nElm < oSet.tpls.length; nElm++)
			aTuples.push(this.SerTuples(oSet.tpls[nElm]));

		return {
			"count":         oSet.tpls.length > 0 ? oSet.tpls.length : undefined,
			"maxRank":       oSet.maxRank != null ? oSet.maxRank : undefined,
			"setDefinition": oSet.setDefinition != null ? oSet.setDefinition : undefined,
			"sortType":      c_oAscSortType.None !== oSet.sortType ? ToXml_ST_SortType(oSet.sortType) : undefined,
			"queryFailed":   oSet.queryFailed !== false ? oSet.queryFailed : undefined,
			"tpls":          aTuples.length > 0 ? aTuples : undefined,
			"sortByTuple":   oSet.sortByTuple !== null ? oSet.sortByTuple : undefined
		}
	};
	WriterToJSON.prototype.SerQueryCache = function(oQuery) // CT_QueryCache
	{
		if (!oQuery)
			return oQuery;

		var aQuery = [];
		for (var nElm = 0; nElm < oQuery.query.length; nElm++)
			aQuery.push(this.SerQuery(oQuery.query[nElm]));

		return {
			"count": oQuery.query.length > 0 ? oQuery.query.length : undefined,
			"query": aQuery
		}
	};
	WriterToJSON.prototype.SerQuery = function(oQuery) // CT_Query
	{
		return {
			"mdx":  oQuery.mdx != null ? oQuery.mdx : undefined,
			"tpls": oQuery.tpls != null ? this.SerTuples(oQuery.tpls) : undefined
		}
	};
	WriterToJSON.prototype.SerServerFormats = function(oFormats) // CT_ServerFormats
	{
		if (!oFormats)
			return oFormats;

		var aFormat = [];
		for (var nElm = 0; nElm < oFormats.serverFormat.length; nElm++)
			aFormat.push(this.SerServerFormat(oFormats.serverFormat[nElm]));

		return {
			"count":        oFormats.serverFormat.length > 0 ? oFormats.serverFormat.length : undefined,
			"serverFormat": aFormat
		}
	};
	WriterToJSON.prototype.SerServerFormat = function(oFormat) // CT_ServerFormat
	{
		return {
			"culture": oFormat.culture != null ? oFormat.culture : undefined,
			"format":  oFormat.format != null ? oFormat.format : undefined
		}
	};
	WriterToJSON.prototype.SerCalculatedItems = function(oItems) // CT_CalculatedItems
	{
		if (!oItems)
			return oItems;

		var aItem = [];
		for (var nElm = 0; nElm < oItems.calculatedItem.length; nElm++)
			aItem.push(this.SerCalculatedItem(oItems.calculatedItem[nElm]));

		return {
			"count":          oItems.calculatedItem.length > 0 ? oItems.calculatedItem.length : undefined,
			"calculatedItem": aItem
		}
	};
	WriterToJSON.prototype.SerCalculatedItem = function(oItem) // CT_CalculatedItem
	{
		return {
			"field":     oItem.field != null ? oItem.field : undefined,
			"formula":   oItem.formula != null ? oItem.formula : undefined,
			"pivotArea": oItem.pivotArea != null ? this.SerPivotArea(oItem.pivotArea) : undefined,
			"extLst":    oItem.extLst != null ? this.SerExtensionList(oItem.extLst) : undefined
		}
	};
	WriterToJSON.prototype.SerCalculatedMembers = function(oMembers) // CT_CalculatedMembers
	{
		if (!oMembers)
			return oMembers;

		var aMember = [];
		for (var nElm = 0; nElm < oMembers.calculatedMember.length; nElm++)
			aMember.push(this.SerCalculatedMember(oMembers.calculatedMember[nElm]));

		return {
			"count":            oMembers.calculatedMember.length > 0 ? oMembers.calculatedMember.length : undefined,
			"calculatedMember": aMember
		}
	};
	WriterToJSON.prototype.SerCalculatedMember = function(oMember) // CT_CalculatedMember
	{
		return {
			// attributes
			"name":       oMember.name != null ? oMember.name : undefined,
			"mdx":        oMember.mdx != null ? oMember.mdx : undefined,
			"memberName": oMember.memberName != null ? oMember.memberName : undefined,
			"hierarchy":  oMember.hierarchy != null ? oMember.hierarchy : undefined,
			"parent":     oMember.parent != null ? oMember.parent : undefined,
			"solveOrder": oMember.solveOrder !== 0 ? oMember.solveOrder : undefined,
			"set":        oMember.set !== false ? oMember.set : undefined,
			// members
			"extLst":     oMember.extLst != null ? this.SerExtensionList(oMember.extLst) : undefined
		}
	};
	WriterToJSON.prototype.SerDimensions = function(oDimensions) // CT_Dimensions
	{
		if (!oDimensions)
			return oDimensions;

		var aDimension = [];
		for (var nElm = 0; nElm < oDimensions.dimension.length; nElm++)
			aDimension.push(this.SerDimension(oDimensions.dimension[nElm]));

		return {
			"count":     oDimensions.dimension.length > 0 ? oDimensions.dimension.length : undefined,
			"dimension": aDimension
		}
	};
	WriterToJSON.prototype.SerDimension = function(oDimension) // CT_PivotDimension
	{
		return {
			"measure":    oDimension.measure !== false ? oDimension.measure : undefined,
			"name":       oDimension.name != null ? oDimension.name : undefined,
			"uniqueName": oDimension.uniqueName != null ? oDimension.uniqueName : undefined,
			"caption":    oDimension.caption != null ? oDimension.caption : undefined
		}
	};
	WriterToJSON.prototype.SerMeasureGroups = function(oGroups) // CT_MeasureGroups
	{
		if (!oGroups)
			return oGroups;

		var aGroup = [];
		for (var nElm = 0; nElm < oGroups.measureGroup.length; nElm++)
			aGroup.push(this.SerMeasureGroup(oGroups.measureGroup[nElm]));

		return {
			"count":        oGroups.measureGroup.length > 0 ? oGroups.measureGroup.length : undefined,
			"measureGroup": aGroup
		}
	};
	WriterToJSON.prototype.SerMeasureGroup = function(oGroup) // CT_MeasureGroup
	{
		return {
			"name":    oGroup.name != null ? oGroup.name : undefined,
			"caption": oGroup.caption != null ? oGroup.caption : undefined
		}
	};
	WriterToJSON.prototype.SerMeasureDimensionMaps = function(oMaps) // CT_MeasureDimensionMaps
	{
		if (!oMaps)
			return oMaps;

		var aMap = [];
		for (var nElm = 0; nElm < oMaps.map.length; nElm++)
			aMap.push(this.SerMeasureDimensionMap(oMaps.map[nElm]));

		return {
			"count": oMaps.map.length > 0 ? oMaps.map.length : undefined,
			"map":   aMap
		}
	};
	WriterToJSON.prototype.SerMeasureDimensionMap = function(oMap) // CT_MeasureDimensionMap
	{
		return {
			"measureGroup": oMap.measureGroup != null ? oMap.measureGroup : undefined,
			"dimension":    oMap.dimension != null ? oMap.dimension : undefined
		}
	};
	WriterToJSON.prototype.SerTabularSlicerCache = function(oCache) // CT_tabularSlicerCache
	{
		if (!oCache)
			return oCache;

		return {
			"pivotCacheId":         oCache.getPivotCacheId(),
			"sortOrder":            oCache.sortOrder != null ? ToXML_ST_tabularSlicerCacheSortOrder(oCache.sortOrder) : undefined,
			"customListSort":       oCache.customListSort != null ? oCache.customListSort : undefined,
			"showMissing":          oCache.showMissing != null ? oCache.showMissing : undefined,
			"crossFilter":          oCache.crossFilter != null ? ToXML_ST_slicerCacheCrossFilter(oCache.crossFilter) : undefined,
			"items":                oCache.items.length > 0 ? this.SerTabularSlicerCacheItems(oCache.items) : undefined
		}
	};
	WriterToJSON.prototype.SerTabularSlicerCacheItems = function(aItems) // CT_tabularSlicerCacheItem[]
	{
		var aResult = [];
		for (var nItem = 0; nItem < aItems.length; nItem++)
			aResult.push(this.SerTabularSlicerCacheItem(aItems[nItem]));

		return aResult;
	};
	WriterToJSON.prototype.SerTabularSlicerCacheItem = function(oItem) // CT_tabularSlicerCacheItem
	{
		return {
			"x":  oItem.x != null ? oItem.x : undefined,
			"s":  oItem.s != null ? oItem.s : undefined,
			"nd": oItem.nd != null ? oItem.nd : undefined
		}
	};
	WriterToJSON.prototype.SerNamedSheetViews = function(aViews) // CT_NamedSheetView[]
	{
		var aResult = [];
		for (var nView = 0; nView < aViews.length; nView++)
			aResult.push(this.SerNamedSheetView(aViews[nView]));
	
		return aResult;
	};
	WriterToJSON.prototype.SerNamedSheetView = function(oView) // CT_NamedSheetView
	{
		var aNsvFilter = [];
		for (var nFilter = 0; nFilter < oView.nsvFilters.length; nFilter++)
			aNsvFilter.push(this.SerNsvFilter(oView.nsvFilters[nFilter]));
			
		return {
			"name":       oView.name != null ? oView.name : undefined,
			"id":         oView.id != null ? oView.id : undefined,
			"nsvFilters": aNsvFilter.length > 0 ? aNsvFilter : undefined
		}
	};
	WriterToJSON.prototype.SerNsvFilter = function(oFilter) // CT_NsvFilter
	{
		var aColFilter = [];
		for (var nFilter = 0; nFilter < oFilter.columnsFilter.length; nFilter++)
			aColFilter.push(this.SerFilterColumn(oFilter.columnsFilter[nFilter]));

		return {
			"filterId":      oFilter.filterId != null ? oFilter.filterId : undefined,
			"ref":           oFilter.ref != null ? this.SerRef(oFilter.ref) : undefined,
			"tableIdOpen":   oFilter.tableIdOpen != null ? oFilter.tableIdOpen : undefined, // тут хз
			"tableId":       oFilter.tableId != null ? oFilter.tableId : undefined, // мб не надо, надо смотреть
			"columnsFilter": aColFilter.length > 0 ? aColFilter : undefined,
			"sortRules":     oFilter.sortRules != null ? this.SerSortRules(oFilter.sortRules) : undefined
		}
	};
	WriterToJSON.prototype.SerSortRules = function(aRules) // CT_SortRules
	{
		if (!aRules)
			return aRules;

		var aResult = [];
		for (var nRule = 0; aRules.length; nRule++)
			aResult.push(this.SerSortRule(aRules[nRule]));

		return {
			"sortRule": aResult.length > 0 ? aResult : undefined
		}
	};
	WriterToJSON.prototype.SerSortRule = function(oRule) // CT_SortRule
	{
		return {
			"colId":         oRule.colId != null ? oRule.colId : undefined,
			"id":            oRule.id != null ? oRule.id : undefined,
			"sortCondition": oRule.sortCondition != null ? this.SerSortCondition(oRule.sortCondition) : undefined,
		}
	};
	WriterToJSON.prototype.SerSheetData = function(oWorksheet)
	{
		var oThis = this;
		var aRows = [];
		var oRange = oWorksheet.getRange3(0, 0, AscCommon.gc_nMaxRow0, AscCommon.gc_nMaxCol0);
		var oTempRow, oTempCell;

		function SerRow(oRow)
		{
			oTempRow = {
				"collapsed":    oRow.getCollapsed(),
				"customHeight": oRow.getCustomHeight(),
				"hidden":       oRow.getHidden(),
				"ht":           oRow.getHeight(),
				"outlineLevel": oRow.getOutlineLevel(),
				"r":            oRow.index + 1,
				"s":            oThis.stylesForWrite.add(oRow.getStyle()),
				"cell":         []
			}
			aRows.push(oTempRow);
		}

		function SerCell(oCell)
		{
			var sRef = AscCommon.g_oCellAddressUtils.colnumToColstr(oCell.nCol + 1) + (oCell.nRow + 1);
			var oFormulaForWrite = oCell.isFormula() ? oThis.InitSaveManager.PrepareFormulaToWrite(oCell) : null;
			oTempCell = {
				"f": SerFormula(oFormulaForWrite),
				"r": sRef,
				"s": oThis.stylesForWrite.add(oCell.getStyle()),
				"t": ToXml_ST_CellValueType(oCell.type)
			}
			if (oCell.multiText)
			{
				oTempCell["v"] = [];
				for (let Index = 0; Index < oCell.multiText.length; Index++)
					oTempCell["v"].push(SerMultiTextElem(oCell.multiText[Index]));
			}
			else
			{
				switch (oCell.type)
				{
					case AscCommon.CellValueType.String:
					case AscCommon.CellValueType.Error:
						oTempCell["v"] = oCell.text;
						break;
					case AscCommon.CellValueType.Bool:
					default:
						oTempCell["v"] = oCell.number;
						break;
				}
			}

			oTempRow["cell"].push(oTempCell);
		}

		function SerFormula(oFormulaForWrite)
		{
			if (!oFormulaForWrite)
				return undefined;
				
			return {
				"t":   ToXml_ST_CellFormulaType(oFormulaForWrite.type),
				"ref": oFormulaForWrite.ref != null ? oFormulaForWrite.ref.getName() : undefined,
				"v":   oFormulaForWrite.formula,
				"ca":  oFormulaForWrite.ca != null ? oFormulaForWrite.ca : undefined,
				"si":  oFormulaForWrite.si != null ? oFormulaForWrite.si : undefined
			}
		}

		function SerMultiTextElem(oMultiTextElem)
		{
			return {
				"text":   oMultiTextElem.text,
				"format": oThis.SerFontExcell(oMultiTextElem.format)
			}
		}

		oRange._foreachNoEmpty(SerCell, SerRow);
		return aRows;
	};
	WriterToJSON.prototype.SerDataValidations = function(oDataValidations)
	{
		var aElems = [];
		for (var nElem = 0; nElem < oDataValidations.elems.length; nElem++)
			aElems.push(this.SerDataValidation(oDataValidations.elems[nElem]));

		return {
			"dataValidation": aElems,
			"disablePrompts": oDataValidations.disablePrompts != null ? oDataValidations.disablePrompts : undefined,
			"xWindow":        oDataValidations.xWindow != null ? oDataValidations.xWindow : undefined,
			"yWindow":        oDataValidations.yWindow != null ? oDataValidations.yWindow : undefined
		}
	};
	WriterToJSON.prototype.SerDataValidation = function(oDataValidation)
	{
		return {
			"formula1": oDataValidation.formula1 != null ? this.SerFormula(oDataValidation.formula1) : undefined,
			"formula2": oDataValidation.formula2 != null ? this.SerFormula(oDataValidation.formula2) : undefined,

			"allowBlank":       oDataValidation.allowBlank != null ? oDataValidation.allowBlank : undefined,
			"error":            oDataValidation.error != null ? oDataValidation.error : undefined,
			"errorStyle":       oDataValidation.errorStyle != null ? ToXML_ST_DataValidationErrorStyle(oDataValidation.errorStyle) : undefined,
			"errorTitle":       oDataValidation.errorTitle != null ? oDataValidation.errorTitle : undefined,
			"imeMode":          oDataValidation.imeMode != null ? ToXML_ST_DataValidationImeMode(oDataValidation.imeMode) : undefined,
			"operator":         oDataValidation.operator != null ? ToXML_ST_DataValidationOperator(oDataValidation.operator) : undefined,
			"prompt":           oDataValidation.prompt != null ? oDataValidation.prompt : undefined,
			"promptTitle":      oDataValidation.promptTitle != null ? oDataValidation.promptTitle : undefined,
			"showDropDown":     oDataValidation.showDropDown != null ? oDataValidation.showDropDown : undefined,
			"showErrorMessage": oDataValidation.showErrorMessage != null ? oDataValidation.showErrorMessage : undefined,
			"showInputMessage": oDataValidation.showInputMessage != null ? oDataValidation.showInputMessage : undefined,
			"sqref":            oDataValidation.ranges != null ? AscCommonExcel.getSqRefString(oDataValidation.ranges) : undefined,
			"type":             oDataValidation.type != null ? ToXML_ST_DataValidationType(oDataValidation.type) : undefined
		}
	};
	WriterToJSON.prototype.SerFormula = function(oFormula) // Asc.CDataFormula
	{
		return {
			"formula": oFormula.text
		}
	};
	WriterToJSON.prototype.SerCondFormatting = function(aCondFormattingRules)
	{
		var aCondRules = [];
		for (var nRule = 0; nRule < aCondFormattingRules.length; nRule++)
			aCondRules.push(this.SerCondFormattingRule(aCondFormattingRules[nRule]));

		return aCondRules;
	};
	WriterToJSON.prototype.SerCondFormattingRule = function(oCondRule)
	{
		var aRules = [];
		for (var nRule = 0; nRule < oCondRule.aRuleElements.length; nRule++)
			aRules.push(this.SerCondFormRuleElement(oCondRule.aRuleElements[nRule]));

		return {
			"aboveAverage": oCondRule.aboveAverage != null ? oCondRule.aboveAverage : undefined,
			"bottom":       oCondRule.bottom != null ? oCondRule.bottom : undefined,
			"dxf":          oCondRule.dxf != null ? this.SerDxf(oCondRule.dxf) : undefined,
			"equalAverage": oCondRule.equalAverage != null ? oCondRule.equalAverage : undefined,
			"operator":     oCondRule.operator != null ? ToXML_CFOperatorType(oCondRule.operator) : undefined,
			"percent":      oCondRule.percent != null ? oCondRule.percent : undefined,
			"priority":     oCondRule.priority != null ? oCondRule.priority : undefined,
			"rank":         oCondRule.rank != null ? oCondRule.rank : undefined,
			"stdDev":       oCondRule.stdDev != null ? oCondRule.stdDev : undefined,
			"stopIfTrue":   oCondRule.stopIfTrue != null ? oCondRule.stopIfTrue : undefined,
			"text":         oCondRule.text != null ? oCondRule.text : undefined,
			"timePeriod":   oCondRule.timePeriod != null ? oCondRule.timePeriod : undefined,
			"type":         oCondRule.type != null ? ToXML_CfRuleType(oCondRule.type) : undefined,
			"sqref":        oCondRule.ranges != null ? AscCommonExcel.getSqRefString(oCondRule.ranges) : undefined,
			"pivot":        oCondRule.pivot != null ? oCondRule.pivot : undefined,
			"rules":        aRules
		}
	};
	WriterToJSON.prototype.SerCondFormRuleElement = function(oRule)
	{
		if (oRule instanceof AscCommonExcel.CColorScale)
			return this.SerColorScale(oRule);
		if (oRule instanceof AscCommonExcel.CDataBar)
			return this.SerDataBar(oRule);
		if (oRule instanceof AscCommonExcel.CIconSet)
			return this.SerIconSet(oRule);
		if (oRule instanceof AscCommonExcel.CFormulaCF)
			return this.SerFormulaCF(oRule);
	};
	WriterToJSON.prototype.SerFormulaCF = function(oFormulaCf)
	{
		return {
			"formula": oFormulaCf.Text,
			"type":    "formulaCf"
		}
	};
	WriterToJSON.prototype.SerIconSet = function(oIconSet)
	{
		var aCFVO = [];
		for (var nElem = 0; nElem < oIconSet.aCFVOs.length; nElem++)
			aCFVO.push(this.SerCondFmtValObj(oIconSet.aCFVOs[nElem]));

		var aCFIS = [];
		for (nElem = 0; nElem < oIconSet.aIconSets.length; nElem++)
			aCFIS.push(this.SerCondFmtIconSet(oIconSet.aIconSets[nElem]));

		return {
			"cfvo":      aCFVO,
			"cfIcon":    aCFIS,
			"iconSet":   oIconSet.IconSet != null ? ToXML_IconSetType(oIconSet.IconSet) : undefined,
			"percent":   oIconSet.Percent != null ? oIconSet.Percent : undefined,
			"reverse":   oIconSet.Reverse != null ? oIconSet.Reverse : undefined,
			"showValue": oIconSet.ShowValue != null ? oIconSet.ShowValue : undefined,
			"type":      "iconSet"
		}
	};
	WriterToJSON.prototype.SerCondFmtIconSet = function(oCFIS)
	{
		return {
			"iconSet": oCFIS.IconSet != null ? ToXML_IconSetType(oCFIS.IconSet) : undefined,
			"iconId":  oCFIS.IconId != null ? oCFIS.IconId : undefined
		}
	};
	WriterToJSON.prototype.SerDataBar = function(oDataBar)
	{
		var aCFVO = [];
		for (var nElem = 0; nElem < oDataBar.aCFVOs.length; nElem++)
			aCFVO.push(this.SerCondFmtValObj(oDataBar.aCFVOs[nElem]));

		return {
			"cfvo":                  aCFVO,
			"color":               oDataBar.Color != null ? this.SerColorExcell(oDataBar.Color) : undefined,
			"negativeColor":       oDataBar.NegativeColor != null ? this.SerColorExcell(oDataBar.NegativeColor) : undefined,
			"borderColor":         oDataBar.BorderColor != null ? this.SerColorExcell(oDataBar.BorderColor) : undefined,
			"axisColor":           oDataBar.AxisColor != null ? this.SerColorExcell(oDataBar.AxisColor) : undefined,
			"negativeBorderColor": oDataBar.NegativeBorderColor != null ? this.SerColorExcell(oDataBar.NegativeBorderColor) : undefined,
			
			"maxLength": oDataBar.MaxLength != null ? oDataBar.MaxLength : undefined,
			"minLength": oDataBar.MinLength != null ? oDataBar.MinLength : undefined,
			"showValue": oDataBar.ShowValue != null ? oDataBar.ShowValue : undefined,
			"axPos":     oDataBar.AxisPosition != null ? ToXML_EDataBarAxisPosition(oDataBar.AxisPosition) : undefined,
			"dir":       oDataBar.Direction != null ? ToXML_EDataBarDirection(oDataBar.Direction) : undefined,
			"gradient":  oDataBar.Gradient != null ? oDataBar.Gradient : undefined,
			"negBarClrSameAsPositive":    oDataBar.NegativeBarColorSameAsPositive != null ? oDataBar.NegativeBarColorSameAsPositive : undefined,
			"negBarBdrClrSameAsPositive": oDataBar.NegativeBarBorderColorSameAsPositive != null ? oDataBar.NegativeBarBorderColorSameAsPositive : undefined,
			"type":                       "dataBar"
		}
	};
	WriterToJSON.prototype.SerColorScale = function(oColorScale)
	{
		var aCFVO = [];
		for (var nElem = 0; nElem < oColorScale.aCFVOs.length; nElem++)
			aCFVO.push(this.SerCondFmtValObj(oColorScale.aCFVOs[nElem]));

		var aColors = [];
		for (var nColor = 0; nColor < oColorScale.aColors.length; nColor++)
			aColors.push(this.SerColorExcell(oColorScale.aColors[nColor]));

		return {
			"cfvo":  aCFVO,
			"color": aColors,
			"type":  "clrScale"
		}
	};
	WriterToJSON.prototype.SerCondFmtValObj = function(oCondFmtValObj)
	{
		return {
			"gte":  oCondFmtValObj.Gte != null ? oCondFmtValObj.Gte : undefined,
			"type": oCondFmtValObj.Type != null ? ToXML_ST_CfvoType(oCondFmtValObj.Type) : undefined,
			"val":  oCondFmtValObj.Val != null ? oCondFmtValObj.Val : undefined
		}
	};
	WriterToJSON.prototype.SerColorExcell = function(oColor)
	{
		if (!oColor)
			return oColor;

		var res;
		if (oColor instanceof AscCommonExcel.ThemeColor)
			res = {
				"theme": oColor.theme != null ? oColor.theme : undefined,
				"tint":  oColor.tint != null ? oColor.tint : undefined,
				"type":  "themeClr"

			}
		else if (oColor instanceof AscCommonExcel.RgbColor)
			res = {
				"rgb":   oColor.rgb != null ? oColor.rgb : undefined,
				"type":  "rgbClr"
			}

		return res;
	};
	WriterToJSON.prototype.SerCols = function(oWs)
	{
		var aResult = [];

		var aCols = oWs.aCols;
		var oPrevCol = null;
		var nPrevIndexStart = null;
		var nPrevIndex = null;
		var aIndexes = [];
		for(var i in aCols)
			aIndexes.push(i - 0);
		aIndexes.sort(AscCommon.fSortAscending);
		var fInitCol = function(col, nMin, nMax){
			var oRes = {col: col, Max: nMax, Min: nMin, xfsid: null, width: col.width};
			if(null == oRes.width)
			{
				if(null != oWs.oSheetFormatPr.dDefaultColWidth)
					oRes.width = oWs.oSheetFormatPr.dDefaultColWidth;
				else
					oRes.width = AscCommonExcel.oDefaultMetrics.ColWidthChars;
			}
			return oRes;
		};
		var oAllCol = null;
		if(null != oWs.oAllCol)
		{
			oAllCol = fInitCol(oWs.oAllCol, 0, AscCommon.gc_nMaxCol0);
		}
		for(var i = 0 , length = aIndexes.length; i < length; ++i)
		{
			var nIndex = aIndexes[i];
			var col = aCols[nIndex];
			if(null != col)
			{
				if(false == col.isEmpty())
				{
					if(null != oAllCol && null == nPrevIndex && nIndex > 0)
					{
						oAllCol.Min = 1;
						oAllCol.Max = nIndex;
						aResult.push(this.SerCol(oAllCol));
					}
					if(null != nPrevIndex && (nPrevIndex + 1 != nIndex || false == oPrevCol.isEqual(col)))
					{
						var oColToWrite = fInitCol(oPrevCol, nPrevIndexStart + 1, nPrevIndex + 1);
						aResult.push(this.SerCol(oColToWrite));
						nPrevIndexStart = null;
						if(null != oAllCol && nPrevIndex + 1 != nIndex)
						{
							oAllCol.Min = nPrevIndex + 2;
							oAllCol.Max = nIndex;
							aResult.push(this.SerCol(oAllCol));
						}
					}
					oPrevCol = col;
					nPrevIndex = nIndex;
					if(null == nPrevIndexStart)
						nPrevIndexStart = nPrevIndex;
				}
			}
		}
		if(null != nPrevIndexStart && null != nPrevIndex && null != oPrevCol)
		{
			var oColToWrite = fInitCol(oPrevCol, nPrevIndexStart + 1, nPrevIndex + 1);
			aResult.push(this.SerCol(oColToWrite));
		}
		if(null != oAllCol)
		{
			if(null == nPrevIndex)
			{
				oAllCol.Min = 1;
				oAllCol.Max = AscCommon.gc_nMaxCol0 + 1;
				aResult.push(this.SerCol(oAllCol));
			}
			else if(AscCommon.gc_nMaxCol0 != nPrevIndex)
			{
				oAllCol.Min = nPrevIndex + 2;
				oAllCol.Max = AscCommon.gc_nMaxCol0 + 1;
				aResult.push(this.SerCol(oAllCol));
			}
		}

		return aResult;
	};
	WriterToJSON.prototype.SerCol = function(oCol)
	{
		return {
			"bestFit":      oCol.col.BestFit != null ? oCol.col.BestFit : undefined,
			"hidden":       oCol.col.hd ? oCol.col.hd : undefined,
			"max":          oCol.Max != null ? oCol.Max : undefined,
			"min":          oCol.Min != null ? oCol.Min : undefined,
			"xfs":          oCol.col.xfs != null ? this.stylesForWrite.add(oCol.col.xfs) : undefined,
			"width":        oCol.col.width != null ? oCol.col.width : undefined,
			"customWidth":  oCol.col.CustomWidth != null ? oCol.col.CustomWidth : undefined,
			"outlineLevel": oCol.col.outlineLevel > 0 ? oCol.col.outlineLevel : undefined,
			"collapsed":    oCol.col.collapsed ? oCol.col.collapsed : undefined
		}
	};
	WriterToJSON.prototype.SerDxf = function(oDxf)
	{
		return {
			"alignment": oDxf.align != null ? this.SerAlign(oDxf.align) : undefined,
			"border":    oDxf.border != null ? this.SerBorderExcell(oDxf.border) : undefined,
			"fill":      oDxf.fill != null ? this.SerFillExcell(oDxf.fill) : undefined,
			"font":      oDxf.font != null ? this.SerFontExcell(oDxf.font) : undefined,
			"numFmt":    oDxf.num != null ? this.SerNumFmtExcell(oDxf.num) : undefined
		}
	};
	WriterToJSON.prototype.SerXfs = function(oXfs, isCellStyle)
	{
		var xf = oXfs.xf;
		return {
			"applyBorder":       oXfs.ApplyBorder != null ? oXfs.ApplyBorder : undefined,
			"borderId":          oXfs.borderid != null ? oXfs.borderid : undefined,
			"applyFill":         oXfs.ApplyFill != null ? oXfs.ApplyFill : undefined,
			"fillId":            oXfs.fillid != null ? oXfs.fillid : undefined,
			"applyFont":         oXfs.ApplyFont != null ? oXfs.ApplyFont : undefined,
			"fontId":            oXfs.fontid != null ? oXfs.fontid : undefined,
			"applyNumberFormat": oXfs.ApplyNumberFormat != null ? oXfs.ApplyNumberFormat : undefined,
			"numFmtId":          oXfs.numid != null ? oXfs.numid : undefined,
			"applyAlignment":    oXfs.ApplyAlignment != null ? oXfs.ApplyAlignment : undefined,
			"alignment":         oXfs.alignMinimized != null ? this.SerAlign(oXfs.alignMinimized) : undefined,
			"quotePrefix":       xf.QuotePrefix != null ? xf.QuotePrefix : undefined,
			"pivotButton":       xf.PivotButton != null ? xf.PivotButton : undefined,
			"xfId":              xf.XfId != null && !isCellStyle ? xf.XfId : undefined,
			"applyProtection":   xf.applyProtection != null ? xf.applyProtection : undefined,
			"protection": null != xf.locked || null != xf.hidden ? {
				"hidden": xf.hidden != null ? xf.hidden : undefined,
				"locked": xf.locked != null ? xf.locked : undefined
			} : undefined
		}
	};
	WriterToJSON.prototype.SerCellStyle = function(oStyle)
	{
		return {
			"builtinId":     oStyle.BuiltinId != null ? oStyle.BuiltinId : undefined,
			"customBuiltin": oStyle.CustomBuiltin != null ? oStyle.CustomBuiltin : undefined,
			"hidden":        oStyle.Hidden != null ? oStyle.Hidden : undefined,
			"iLevel":        oStyle.ILevel != null ? oStyle.ILevel : undefined,
			"name":          oStyle.Name != null ? oStyle.Name : undefined,
			"xfId":          oStyle.XfId != null ? oStyle.XfId : undefined
		}
	};
	WriterToJSON.prototype.SerTableStyles = function(oStyles)
	{
		var bEmptyCustom = true;
		for (var i in oStyles.CustomStyles)
		{
			bEmptyCustom = false;
			break;
		}

		return {
			"defaultTableStyle": oStyles.DefaultTableStyle != null ? oStyles.DefaultTableStyle : undefined,
			"defaultPivotStyle": oStyles.DefaultPivotStyle != null ? oStyles.DefaultPivotStyle : undefined,
			"customStyle":       bEmptyCustom === false ? this.SerTableCustomStyles(oStyles.CustomStyles) : undefined
		}
	};
	WriterToJSON.prototype.SerTableCustomStyles = function(oCustomStyles)
	{
		var aResult = [];
		for (var i in oCustomStyles)
			aResult.push(this.SerTableCustomStyle(oCustomStyles[i]));

		return aResult;
	};
	WriterToJSON.prototype.SerTableCustomStyle = function(oCustomStyle)
	{
		var aElements = [];
		if (oCustomStyle.blankRow != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.blankRow, "blankRow"));	
		if (oCustomStyle.firstColumn != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.firstColumn, "firstColumn"));
		if (oCustomStyle.firstColumnStripe != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.firstColumnStripe, "firstColumnStripe"));
		if (oCustomStyle.firstColumnSubheading != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.firstColumnSubheading, "firstColumnSubheading"));
		if (oCustomStyle.firstHeaderCell != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.firstHeaderCell, "firstHeaderCell"));
		if (oCustomStyle.firstRowStripe != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.firstRowStripe, "firstRowStripe"));
		if (oCustomStyle.firstRowSubheading != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.firstRowSubheading, "firstRowSubheading"));
		if (oCustomStyle.firstSubtotalColumn != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.firstSubtotalColumn, "firstSubtotalColumn"));
		if (oCustomStyle.firstSubtotalRow != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.firstSubtotalRow, "firstSubtotalRow"));
		if (oCustomStyle.firstTotalCell != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.firstTotalCell, "firstTotalCell"));
		if (oCustomStyle.headerRow != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.headerRow, "headerRow"));
		if (oCustomStyle.lastColumn != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.lastColumn, "lastColumn"));
		if (oCustomStyle.lastHeaderCell != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.lastHeaderCell, "lastHeaderCell"));
		if (oCustomStyle.lastTotalCell != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.lastTotalCell, "lastTotalCell"));
		if (oCustomStyle.pageFieldLabels != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.pageFieldLabels, "pageFieldLabels"));
		if (oCustomStyle.pageFieldValues != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.pageFieldValues, "pageFieldValues"));
		if (oCustomStyle.secondColumnStripe != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.secondColumnStripe, "secondColumnStripe"));
		if (oCustomStyle.secondColumnSubheading != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.secondColumnSubheading, "secondColumnSubheading"));
		if (oCustomStyle.secondRowStripe != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.secondRowStripe, "secondRowStripe"));
		if (oCustomStyle.secondRowSubheading != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.secondRowSubheading, "secondRowSubheading"));
		if (oCustomStyle.secondSubtotalColumn != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.secondSubtotalColumn, "secondSubtotalColumn"));
		if (oCustomStyle.secondSubtotalRow != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.secondSubtotalRow, "secondSubtotalRow"));
		if (oCustomStyle.thirdColumnSubheading != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.thirdColumnSubheading, "thirdColumnSubheading"));
		if (oCustomStyle.thirdRowSubheading != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.thirdRowSubheading, "thirdRowSubheading"));
		if (oCustomStyle.thirdSubtotalColumn != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.thirdSubtotalColumn, "thirdSubtotalColumn"));
		if (oCustomStyle.thirdSubtotalRow != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.thirdSubtotalRow, "thirdSubtotalRow"));
		if (oCustomStyle.totalRow != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.totalRow, "totalRow"));
		if (oCustomStyle.wholeTable != null)
			aElements.push(this.SerTableCustomStyleElement(oCustomStyle.wholeTable, "wholeTable"));

		return {
			"name":     oCustomStyle.name != null ? oCustomStyle.name : undefined,
			"pivot":    oCustomStyle.pivot != null ? oCustomStyle.pivot : undefined,
			"table":    oCustomStyle.table != null ? oCustomStyle.table : undefined,
			"elements": aElements.length > 0 ? aElements : undefined
		}
	};
	WriterToJSON.prototype.SerTableCustomStyleElement = function(oCustomStyleElement, sType)
	{
		if (oCustomStyleElement.dxf)
			this.InitSaveManager.aDxfs.push(oCustomStyleElement.dxf);

		return {
			"type":  sType,
			"size":  oCustomStyleElement.size != null ? oCustomStyleElement.size : undefined,
			"dxfId": oCustomStyleElement.dxf != null ? this.InitSaveManager.aDxfs.length : undefined
		}
	};
	WriterToJSON.prototype.SerBorderExcell = function(oBorder)
	{
		if (!oBorder)
			return oBorder;

		return {
			"bottom":        oBorder.b != null ? this.SerBorderProp(oBorder.b) : undefined,
			"diagonal":      oBorder.d != null ? this.SerBorderProp(oBorder.d) : undefined,
			"end":           oBorder.r != null ? this.SerBorderProp(oBorder.r) : undefined, // right
			"horizontal":    oBorder.ih != null ? this.SerBorderProp(oBorder.ih) : undefined,
			"start":         oBorder.l != null ? this.SerBorderProp(oBorder.l) : undefined, // left
			"top":           oBorder.t != null ? this.SerBorderProp(oBorder.t) : undefined,
			"vertical":      oBorder.iv != null ? this.SerBorderProp(oBorder.iv) : undefined,
			"diagonalDown":  oBorder.dd != null ? oBorder.dd : undefined,
			"diagonalUp":    oBorder.du != null ? oBorder.du : undefined
		}
	};
	WriterToJSON.prototype.SerBorderProp = function(oProp)
	{
		if (!oProp)
			return oProp;

		return {
			"style": oProp.s != null ? To_XML_EBorderStyle(oProp.s) : undefined,
			"color": oProp.c != null ? this.SerColorExcell(oProp.c) : undefined
		}
	};
	WriterToJSON.prototype.SerFillExcell = function(oFill)
	{
		if (!oFill)
			return oFill;

		oFill.checkEmptyContent();
		return {
			"patternFill":  oFill.patternFill != null ? this.SerPatternFillExcell(oFill.patternFill) : undefined,
			"gradientFill": oFill.gradientFill != null ? this.SerGradientFillExcell(oFill.gradientFill) : undefined
		}
	};
	WriterToJSON.prototype.SerPatternFillExcell = function(oFill)
	{
		if (!oFill)
			return oFill;

		return {
			"fgColor":     oFill.fgColor != null ? this.SerColorExcell(oFill.fgColor) : undefined,
			"bgColor":     oFill.bgColor != null ? this.SerColorExcell(oFill.bgColor) : undefined,
			"patternType": oFill.patternType != null ? ToXml_ST_PatternType(oFill.patternType) : undefined
		}
	};
	WriterToJSON.prototype.SerGradientFillExcell = function(oFill)
	{
		if (!oFill)
			return oFill;

		var aGrStop = [];
		for (var nGrStop = 0; nGrStop < oFill.stop.length; nGrStop++)
			aGrStop.push(this.SerGradStopEcxell(oFill.stop[nGrStop]));
			
		return {
			"left":    oFill.left != null ? oFill.left : undefined,
			"top":     oFill.top != null ? oFill.top : undefined,
			"right":   oFill.right != null ? oFill.right : undefined,
			"bottom":  oFill.bottom != null ? oFill.bottom : undefined,
			"degree":  oFill.degree != null ? oFill.degree : undefined,
			"type":    oFill.type != null ? ToXml_ST_GradientType(oFill.type) : undefined,
			"aGrStop": aGrStop
		}
	};
	WriterToJSON.prototype.SerGradStopEcxell = function(oGrStop)
	{
		if (!oGrStop)
			return null;

		return {
			"color":    oGrStop.color != null ? this.SerColorExcell(oGrStop.color) : undefined,
			"position": oGrStop.position != null ? oGrStop.position : undefined
		}
	};
	WriterToJSON.prototype.SerFontExcell = function(oFont)
	{
		if (!oFont)
			return oFont;

		return {
			"b":         oFont.b != null ? oFont.b : undefined,
			"color":     oFont.c != null ? this.SerColorExcell(oFont.c) : undefined,
			"i":         oFont.i != null ? oFont.i : undefined,
			"name":      oFont.fn != null ? oFont.fn : undefined,
			"strike":    oFont.s != null ? oFont.s : undefined,
			"sz":        oFont.fs != null ? oFont.fs : undefined,
			"u":         oFont.u != null ? ToXML_EUnderline(oFont.u) : undefined,
			"vertAlign": oFont.va != null ? ToXML_ST_VerticalAlignRun(oFont.va) : undefined,
			"scheme":    oFont.scheme != null ? ToXML_EFontScheme(oFont.scheme) : undefined
		}
	};
	WriterToJSON.prototype.SerNumFmtExcell = function(oNumFmt, nId, bGetNumFmtId)
	{
		if (!oNumFmt)
			return oNumFmt;

		var numId;
		if (nId != null)
			numId = nId; 
		else 
		{
			numId = this.stylesForWrite.getNumIdByFormat(oNumFmt);
			if (numId == null)
				numId = oNumFmt.id;
		}

		if (bGetNumFmtId === true)
			return numId;
			
		var sFormat = oNumFmt.getFormat();
		
		return {
			"formatCode": sFormat != null ? sFormat : undefined,
			"numFmtId":   numId != null ? numId : undefined
		}
	};
	WriterToJSON.prototype.SerExcelStylesForWrite = function()
	{
		//var g_oDefaultFormat = AscCommonExcel.g_oDefaultFormat;

		var aBorders = [], aFills = [], aFonts = [],
		aCellStyleXfs = [], aCellXfs = [], aCellStyles = [],
		oTableStyles, aDxfs = [], aExtDxfs = [], aNumFmts = [];

		//borders
		var elems = this.stylesForWrite.oBorderMap.elems;
		for (var i = 0; i < elems.length; ++i)
		{
			aBorders.push(this.SerBorderExcell(elems[i]));
		}

		//fills
		elems = this.stylesForWrite.oFillMap.elems;
		for (i = 0; i < elems.length; ++i)
			aFills.push(this.SerFillExcell(elems[i]));

		//fonts
		elems = this.stylesForWrite.oFontMap.elems;
		for (i = 0; i < elems.length; ++i)
			aFonts.push(this.SerFontExcell(elems[i]));

		//CellStyleXfs
		elems = this.stylesForWrite.oXfsStylesMap;
		for (i = 0; i < elems.length; ++i)
			aCellStyleXfs.push(this.SerXfs(elems[i], true));

		//cellxfs
		elems = this.stylesForWrite.oXfsMap.elems;
		for (i = 0; i < elems.length; ++i)
			aCellXfs.push(this.SerXfs(elems[i], false));

		// cell styles
		for (i = 0; i < this.Workbook.CellStyles.CustomStyles.length; i++)
			aCellStyles.push(this.SerCellStyle(this.Workbook.CellStyles.CustomStyles[i]));
		
		// table styles
		oTableStyles = this.Workbook.TableStyles != null ? this.SerTableStyles(this.Workbook.TableStyles) : undefined;

		//Dxfs пишется после TableStyles, потому что Dxfs может пополниться при записи TableStyles
		for (i = 0; i < this.InitSaveManager.aDxfs.length; i++)
			aDxfs.push(this.SerDxf(this.InitSaveManager.aDxfs[i]));

		var aExtDxfsTemp = [];
		var oSlicerStyles = this.PrepareSlicerStyles(this.Workbook.SlicerStyles, aExtDxfsTemp);
		for (i = 0; i < aExtDxfsTemp.length; i++)
			aExtDxfs.push(this.SerDxf(aExtDxfsTemp[i]));

		// slicer styles
		oSlicerStyles = this.SerSlicerStyles(oSlicerStyles);

		//numfmts пишется в конце потому что они могут пополниться при записи Dxfs
		elems = this.stylesForWrite.oNumMap.elems;
		for (var i = 0; i < elems.length; ++i)
			aNumFmts.push(this.SerNumFmtExcell(elems[i], Asc.g_nNumsMaxId + i));

		return {
			"borders":      aBorders.length > 0 ? aBorders : undefined,
			"fills":        aFills.length > 0 ? aFills : undefined,
			"fonts":        aFonts.length > 0 ? aFonts : undefined,
			"numFmts":      aNumFmts.length > 0 ? aNumFmts : undefined,
			
			"cellStyleXfs": aCellStyleXfs.length > 0 ? aCellStyleXfs : undefined,
			"cellXfs":      aCellXfs.length > 0 ? aCellXfs : undefined,
			"cellStyles":   aCellStyles.length > 0 ? aCellStyles : undefined,
			"dxfs":         aDxfs.length > 0 ? aDxfs : undefined,
			"tableStyles":  oTableStyles,
			"slicerStyles": oSlicerStyles,
			"extDxfs":      aExtDxfs.length > 0 ? aExtDxfs : undefined
		}
	};
	WriterToJSON.prototype.PrepareSlicerStyles = function(slicerStyles, aDxfs)
	{
		var styles = new Asc.CT_slicerStyles();
		styles.defaultSlicerStyle = slicerStyles.DefaultStyle;
		for(var name in slicerStyles.CustomStyles){
			if(slicerStyles.CustomStyles.hasOwnProperty(name)){
				var slicerStyle = new Asc.CT_slicerStyle();
				slicerStyle.name = name;
				var elems = slicerStyles.CustomStyles[name];
				for(var type in elems){
					if(elems.hasOwnProperty(type)) {
						var styleElement = new Asc.CT_slicerStyleElement();
						styleElement.type = parseInt(type);
						styleElement.dxfId = aDxfs.length;
						aDxfs.push(elems[type]);
						slicerStyle.slicerStyleElements.push(styleElement);
					}
				}
				styles.slicerStyle.push(slicerStyle);
			}
		}
		return styles;
	};
	WriterToJSON.prototype.SerSlicerStyles = function(oStyles) // CT_slicerStyles
	{
		var aStyles = [];
		for (var nElm = 0; nElm < oStyles.slicerStyle.length; nElm++)
			aStyles.push(this.SerSlicerStyle(oStyles.slicerStyle[nElm]));

		return {
			"defaultSlicerStyle": oStyles.defaultSlicerStyle != null ? oStyles.defaultSlicerStyle : undefined,
			"slicerStyle":        aStyles.length > 0 ? aStyles : undefined
		}
	};
	WriterToJSON.prototype.SerSlicerStyle = function(oStyle) // CT_slicerStyle
	{
		var aElements = [];
		for (var nElm = 0; nElm < oStyle.slicerStyleElements.length; nElm++)
			aElements.push(this.SerSlicerStyleElement(oStyle.slicerStyleElements[nElm]));

		return {
			"name":                oStyle.name != null ? oStyle.name : undefined,
			"slicerStyleElements": aElements.length > 0 ? aElements : undefined
		}
	};
	WriterToJSON.prototype.SerSlicerStyleElement = function(oStyleElm) // CT_slicerStyleElement
	{
		return {
			"type":  oStyleElm.type != null ? To_XML_ST_slicerStyleType(oStyleElm.type) : undefined,
			"dxfId": oStyleElm.dxfId != null ? oStyleElm.dxfId : undefined
		}
	};
	WriterToJSON.prototype.SerAlign = function(oAlign)
	{
		if (!oAlign)
			return oAlign;

		return {
			"horizontal":     oAlign.hor != null ? ToXml_ST_HorizontalAlignment(oAlign.hor) : undefined,
			"indent":         oAlign.indent != null ? oAlign.indent : undefined,
			"relativeIndent": oAlign.RelativeIndent != null ? oAlign.RelativeIndent : undefined,
			"shrinkToFit":    oAlign.shrink != null ? oAlign.shrink : undefined,
			"textRotation":   oAlign.angle != null ? oAlign.angle : undefined,
			"vertical":       oAlign.ver != null ? ToXml_ST_VerticalAlignment(oAlign.ver) : undefined,
			"wrapText":       oAlign.wrap != null ? oAlign.wrap : undefined
		}
	};
	WriterToJSON.prototype.SerAutoFilter = function(oAutoFilter)
	{
		if (!oAutoFilter)
			return oAutoFilter;

		var aFilterColumns = [];
		if (oAutoFilter.FilterColumns)
		{
			for (var nFilterCol = 0; nFilterCol < oAutoFilter.FilterColumns.length; nFilterCol++)
				aFilterColumns.push(this.SerFilterColumn(oAutoFilter.FilterColumns[nFilterCol]));
		}
		
		return {
			"filterColumn": aFilterColumns.length > 0 ? aFilterColumns : undefined,
			"sortState":    oAutoFilter.SortState != null ? this.SerSortState(oAutoFilter.SortState) : undefined,
			"ref":          oAutoFilter.Ref != null ? oAutoFilter.Ref.getName() : undefined
		}
	};
	WriterToJSON.prototype.SerFilterColumn = function(oFilterColumn) // CT_ColumnFilter
	{
		if (!oFilterColumn)
			return oFilterColumn;

		return {
			"colorFilter":   oFilterColumn.ColorFilter != null ? this.SerColorFilter(oFilterColumn.ColorFilter) : undefined,
			"customFilters": oFilterColumn.CustomFiltersObj != null ? this.SerCustomFilters(oFilterColumn.CustomFiltersObj) : undefined,
			"dynamicFilter": oFilterColumn.DynamicFilter != null ? this.SerDynamicFilter(oFilterColumn.DynamicFilter) : undefined,
			"filters":       oFilterColumn.Filters != null ? this.SerFilters(oFilterColumn.Filters) : undefined,
			"top10":         oFilterColumn.Top10 != null ? this.SerTop10(oFilterColumn.Top10) : undefined,
			"colId":         oFilterColumn.ColId != null ? oFilterColumn.ColId : undefined,
			"showButton":    oFilterColumn.ShowButton != null ? oFilterColumn.ShowButton : undefined
		}
	};
	WriterToJSON.prototype.SerColorFilter = function(oColorFilter)
	{
		if (!oColorFilter)
			return oColorFilter;

		if (oColorFilter.dxf)
			this.InitSaveManager.aDxfs.push(oColorFilter.dxf);

		return {
			"cellColor": oColorFilter.CellColor != null ? oColorFilter.CellColor : undefined,
			"dxfId":     oColorFilter.dxf != null ? this.InitSaveManager.aDxfs.length : undefined
		}
	};
	WriterToJSON.prototype.SerCustomFilters = function(oCustomFilters)
	{
		if (!oCustomFilters)
			return oCustomFilters;

		var aFilters = [];
		if (oCustomFilters.CustomFilters)
		{
			for (var nFilter = 0; nFilter < oCustomFilters.CustomFilters.length; nFilter++)
				aFilters.push(this.SerCustomFilter(oCustomFilters.CustomFilters[nFilter]));
		}

		return {
			"and":           oCustomFilters.And != null ? oCustomFilters.And : undefined,
			"customFilters": aFilters
		}
	};
	WriterToJSON.prototype.SerCustomFilter = function(oCustomFilter)
	{
		return {
			"operator": oCustomFilter.Operator != null ? ToXml_ST_FilterOperator(oCustomFilter.Operator) : undefined,
			"val":      oCustomFilter.Val != null ? oCustomFilter.Val : undefined
		}
	};
	WriterToJSON.prototype.SerDynamicFilter = function(oDynamicFilter)
	{
		if (!oDynamicFilter)
			return oDynamicFilter;

		return {
			"type":      oDynamicFilter.Type != null ? ToXml_ST_DynamicFilterType(oDynamicFilter.Type) : undefined,
			"maxValIso": oDynamicFilter.MaxVal != null ? oDynamicFilter.MaxVal : undefined,
			"valIso":    oDynamicFilter.Val != null ? oDynamicFilter.Val : undefined
		}
	};
	WriterToJSON.prototype.SerFilters = function(oFilters) // AscCommonExcel.Filters()
	{
		if (!oFilters)
			return oFilters;

		var aDateGroupItem = [];
		var oDateGroupItem;
		for (var nItem = 0; nItem < oFilters.Dates.length; nItem++)
		{
			oDateGroupItem = new AscCommonExcel.DateGroupItem();
			oDateGroupItem.convertRangeToDateGroupItem(oFilters.Dates[nItem]);
			aDateGroupItem.push(this.SerDateGroupItem(oDateGroupItem));
		}
			
		var aFilter = [];
		for (var sFilter in oFilters.Values)
			aFilter.push(sFilter);

		return {
			"dateGroupItem": aDateGroupItem,
			"filter":        aFilter,
			"blank":         oFilters.Blank != null ? oFilters.Blank : undefined
		}
	};
	WriterToJSON.prototype.SerDateGroupItem = function(oItem)
	{
		return {
			"day":              oItem.Day != null ? oItem.Day : undefined,
			"hour":             oItem.Hour != null ? oItem.Hour : undefined,
			"minute":           oItem.Minute != null ? oItem.Minute : undefined,
			"month":            oItem.Month != null ? oItem.Month : undefined,
			"second":           oItem.Second != null ? oItem.Second : undefined,
			"year":             oItem.Year != null ? oItem.Year : undefined,
			"dateTimeGrouping": ToXml_ST_DateTimeGrouping(oItem.dateTimeGrouping)
		}
	};
	WriterToJSON.prototype.SerTop10 = function(oTop10)
	{
		if (!oTop10)
			return oTop10;

		return {
			"filterVal": oTop10.FilterVal != null ? oTop10.FilterVal : undefined,
			"percent":   oTop10.Percent != null ? oTop10.Percent : undefined,
			"top":       oTop10.Top != null ? oTop10.Top : undefined,
			"val":       oTop10.Val != null ? oTop10.Val : undefined
		}
	};
	WriterToJSON.prototype.SerSortState = function(oSortState)
	{
		if (!oSortState)
			return oSortState;

		var aSortConditions= [];
		if (oSortState.SortConditions)
		{
			for (var nSort = 0; nSort < oSortState.SortConditions.length; nSort++)
				aSortConditions.push(this.SerSortCondition(oSortState.SortConditions[nSort]));
		}

		return {
			"sortCondition": aSortConditions,
			"caseSensitive": oSortState.CaseSensitive != null ? oSortState.CaseSensitive : undefined,
			"columnSort":    oSortState.ColumnSort != null ? oSortState.ColumnSort : undefined,
			"ref":           oSortState.Ref != null ? this.SerRef(oSortState.Ref) : undefined,
			"sortMethod":    oSortState.SortMethod != null ? ToXML_ST_SortMethod(oSortState.SortMethod) : undefined
		}
	};
	WriterToJSON.prototype.SerSortCondition = function(oSortCond)
	{
		if (!oSortCond)
			return oSortCond;

		if (oSortCond.dxf)
			this.InitSaveManager.aDxfs.push(oSortCond.dxf);

		return {
			"sortBy":     oSortCond.ConditionSortBy != null ? ToXML_ESortBy(oSortCond.ConditionSortBy) : undefined,
			"descending": oSortCond.ConditionDescending != null ? oSortCond.ConditionDescending : undefined,
			"ref":        oSortCond.Ref != null ? this.SerRef(oSortCond.Ref) : undefined,
			"dxfId":      oSortCond.dxf != null ? this.InitSaveManager.aDxfs.length : undefined
		}
	};
	WriterToJSON.prototype.SerRef = function(oRef)
	{
		if (!oRef)
			return oRef;

		var sColumn1 = AscCommon.g_oCellAddressUtils.colnumToColstr(oRef.c1 + 1);
		var sColumn2 = AscCommon.g_oCellAddressUtils.colnumToColstr(oRef.c2 + 1);

		return sColumn1 + (oRef.r1 + 1) + ":" + sColumn2 + (oRef.r2 + 1);
	};
	WriterToJSON.prototype.SerRefCell = function(oRefCell)
	{
		if (!oRefCell)
			return oRefCell;

		var sColumn = AscCommon.g_oCellAddressUtils.colnumToColstr(oRefCell.col + 1);
		return sColumn + (oRefCell.row + 1);
	};
	
	ReaderFromJSON.prototype.WorksheetFromJSON = function(oParsedSheet, oWorkbook)
	{
		History.TurnOff();

		let oWorksheet = new AscCommonExcel.Worksheet(oWorkbook, -1);
		this.curWorksheet = oWorksheet;

		// worksheet props
		oWorksheet.sName = oParsedSheet["name"];
		oWorksheet.bHidden = oParsedSheet["hidden"] === "hidden" ? true : false;

		if (oParsedSheet["cols"] != null)
			this.ColsFromJSON(oParsedSheet["cols"], oWorksheet);
		if (oParsedSheet["sheetFormatPr"] != null)
			this.SheetFormatPrFromJSON(oParsedSheet["sheetFormatPr"], oWorksheet);
		if (oParsedSheet["pageMargins"] != null)
			this.PageMarginsFromJSON(oParsedSheet["pageMargins"], oWorksheet.PagePrintOptions.pageMargins);
		if (oParsedSheet["pageSetup"] != null)
			this.PageSetupExcellFromJSON(oParsedSheet["pageSetup"], oWorksheet.PagePrintOptions.pageSetup); // ???
		if (oParsedSheet["printOptions"] != null)
			this.PrintOptionsExcelFromJSON(oParsedSheet["printOptions"], oWorksheet.PagePrintOptions);
		if (oParsedSheet["mergeCells"] != null)
			this.MergeCellsFromJSON(oParsedSheet["mergeCells"], oWorksheet);
		if (oParsedSheet["sheetData"] != null)
			this.SheetDataFromJSON(oParsedSheet["sheetData"], oWorksheet);
		if (oParsedSheet["hiperlinks"] != null)
			this.HyperlinksExcelFromJSON(oParsedSheet["hiperlinks"], oWorksheet);
		if (oParsedSheet["drawings"] != null)
			this.DrawingsExcellFromJSON(oParsedSheet["drawings"], oWorksheet);
		if (oParsedSheet["autoFilter"] != null)
			oWorksheet.AutoFilter = this.AutoFilterFromJSON(oParsedSheet["autoFilter"], true)
		if (oParsedSheet["sortState"] != null)
			oWorksheet.sortState = this.SortStateFromJSON(oParsedSheet["sortState"]);
		if (oParsedSheet["tableParts"] != null)
			this.TablePartsFromJSON(oParsedSheet["tableParts"], oWorksheet);
		if (oParsedSheet["comments"] != null) 
			this.CommentsFromJSON(oParsedSheet["comments"], oWorksheet);
		if (oParsedSheet["conditionalFormatting"] != null)
			this.CondFormattingFromJSON(oParsedSheet["conditionalFormatting"], oWorksheet);
		if (oParsedSheet["sheetViews"] != null)
			oWorksheet.sheetViews = this.SheetViewsFromJSON(oParsedSheet["sheetViews"], oWorksheet);
		if (oParsedSheet["sheetPr"] != null)
			oWorksheet.sheetPr = this.SheetPrFromJSON(oParsedSheet["sheetPr"]); /// ???
		if (oParsedSheet["sparklineGroup"] != null)
			oWorksheet.aSparklineGroups = this.SparklineGroupsFromJSON(oParsedSheet["sparklineGroup"]);
		if (oParsedSheet["headerFooter"] != null)
			this.HdrFtrExcellFromJSON(oParsedSheet["headerFooter"], oWorksheet.headerFooter);
		if (oParsedSheet["dataValidations"] != null)
			oWorksheet.dataValidations = this.DataValidationsFromJSON(oParsedSheet["dataValidations"]);
		if (oParsedSheet["pivotTables"] != null)
			this.PivotTablesFromJSON(oParsedSheet["pivotTables"], oWorksheet);
		if (oParsedSheet["slicers"] != null)
			this.SlicersFromJSON(oParsedSheet["slicers"], oWorksheet);
		if (oParsedSheet["namedSheetViews"] != null)
			this.NamedSheetViewsFromJSON(oParsedSheet["namedSheetViews"], oWorksheet);
		if (oParsedSheet["sheetProtection"] != null)
			this.SheetProtectionFromJSON(oParsedSheet["sheetProtection"], oWorksheet);
		if (oParsedSheet["protectedRanges"] != null)
			oWorksheet.protectedRanges = this.ProtectedRangesFromJSON(oParsedSheet["protectedRanges"]);

		oWorksheet.initPostOpenZip(this.pivotCaches, this.oNumFmtsOpen);
		History.TurnOn();

		return oWorksheet;
	};
	ReaderFromJSON.prototype.WorksheetsFromJSON = function(oParsedSheets, oWorkbook)
	{
		let api = window["Asc"]["editor"];
		let WorkbookView = api.wb;
		let renameSheetMap = {};
		let oTempWorkBook = new AscCommonExcel.Workbook(undefined, undefined, false);
		let aRestoredSheets = [];
		let oThis = this;
		oTempWorkBook.DrawingDocument = Asc.editor.wbModel.DrawingDocument;
		oTempWorkBook.setCommonIndexObjectsFrom(WorkbookView.model);
		oTempWorkBook.oApi = api;
		oTempWorkBook.theme = oWorkbook.theme;
		this.slicerCachePivotTableSheetIdMap = {};

		this.Workbook = oWorkbook;

		if (this.StyleObject == null)
		{
			let oStyleObject = {aBorders: [], aFills: [], aFonts: [], oNumFmts: {}, aCellStyleXfs: [],
			aCellXfs: [], aDxfs: [], aExtDxfs: [], aCellStyles: [], oCustomTableStyles: {}, oCustomSlicerStyles: null};
			this.StylesFromJSON(oParsedSheets["styles"], oStyleObject);
			this.InitOpenManager = new AscCommonExcel.InitOpenManager(null, this.Workbook, [], false);
			this.InitOpenManager.InitStyleManager(oStyleObject, this.aCellXfs);
			this.oNumFmtsOpen = oStyleObject.oNumFmts;
			this.aDxfs = oStyleObject.aDxfs;
		}

		this.slicerCaches = {};
		this.slicerCachesExt = {};

		if (oParsedSheets["pivotCaches"] != null)
			this.pivotCaches = this.PivotCachesFromJSON(oParsedSheets["pivotCaches"]);
		
		if (oParsedSheets["slicerCaches"] != null)
			this.slicerCaches = this.SlicerCachesFromJSON(oParsedSheets["slicerCaches"]);
		
		if (oParsedSheets["slicerCachesExt"] != null)
			this.slicerCachesExt = this.SlicerCachesFromJSON(oParsedSheets["slicerCachesExt"]);

		for (var Index = 0; Index < oParsedSheets["sheets"].length; Index++)
		{
			let oWorksheet = this.WorksheetFromJSON(oParsedSheets["sheets"][Index], oTempWorkBook);
			oTempWorkBook.aWorksheets.push(oWorksheet);
		}

		let newFonts = {};
		let pasteProcessor = AscCommonExcel.g_clipboardExcel.pasteProcessor;
		newFonts = this.Workbook.generateFontMap2();
		newFonts = pasteProcessor._convertFonts(newFonts);
		
		for (let nSheet = 0; nSheet < oTempWorkBook.aWorksheets.length; nSheet++)
		{
			let oWorksheet = oTempWorkBook.aWorksheets[nSheet];
			for (let i = 0; i < oWorksheet.Drawings.length; i++) {
				oWorksheet.Drawings[i].graphicObject.getAllFonts(newFonts);
			}
		}

		let newSlicerCaches = {};
		let newSlicerCachesExt = {};
		let pasteSheets = function() {
			for (let nSheet = 0; nSheet < oTempWorkBook.aWorksheets.length; nSheet++)
			{
				let oWorksheet = oTempWorkBook.aWorksheets[nSheet];
				oTempWorkBook._updateWorksheetIndexes();
				oTempWorkBook.nActive = oWorksheet.getIndex();
				let sBinarySheet = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(oWorksheet, null, null, true, true);

				let where = WorkbookView.model.aWorksheets.length;
				let renameParams = WorkbookView.model.copyWorksheet(0, where, undefined, undefined, undefined, undefined, oWorksheet, sBinarySheet);
				let oNewWorksheet = WorkbookView.model.aWorksheets[where];
				oThis.slicerCachePivotTableSheetIdMap[oParsedSheets["sheets"][nSheet]["id"]] = oNewWorksheet.getId();

				if (renameParams && renameParams.copySlicerError && WorkbookView.handlers) {
					WorkbookView.handlers.trigger("asc_onError", Asc.c_oAscError.ID.MoveSlicerError, Asc.c_oAscError.Level.NoCritical);
				}

				renameSheetMap[renameParams.lastName] = renameParams.newName;
				api.asc_showWorksheet(where);
				api.asc_setZoom(1);
				// Посылаем callback об изменении списка листов
				api.sheetsChanged();

				for (let i = 0; i < oNewWorksheet.aSlicers.length; ++i) {
					var slicerCache = oNewWorksheet.aSlicers[i].getSlicerCache();
					if (slicerCache) {
						if (oNewWorksheet.aSlicers[i].isExt()) {
							newSlicerCachesExt[slicerCache.name] = slicerCache;
						} else {
							newSlicerCaches[slicerCache.name] = slicerCache;
						}
					}
				}

				aRestoredSheets.push(oNewWorksheet);
			}

			oThis.slicerCaches = newSlicerCaches;
			oThis.slicerCachesExt = newSlicerCachesExt;
			oThis.ResetSlicerCachePivotTablesSheetId();
		};

		api._loadFonts(newFonts, function () {
			pasteSheets();
		});

		return aRestoredSheets;
	};
	ReaderFromJSON.prototype.ResetSlicerCachePivotTablesSheetId = function()
	{
		for (let key in this.slicerCaches)
		{
			let oSlicerCachePivotTable = this.slicerCaches[key].pivotTables[0];
			if (oSlicerCachePivotTable)
			{
				if (this.slicerCachePivotTableSheetIdMap[oSlicerCachePivotTable.sheetId])
					this.slicerCaches[key].forCopySheet(oSlicerCachePivotTable.sheetId, this.slicerCachePivotTableSheetIdMap[oSlicerCachePivotTable.sheetId]);
				else
					this.slicerCaches[key].forCopySheet(oSlicerCachePivotTable.sheetId, null);
			}
		}
		for (let key in this.slicerCachesExt)
		{
			let oSlicerCachePivotTable = this.slicerCachesExt[key].pivotTables[0];
			if (oSlicerCachePivotTable)
			{
				if (this.slicerCachePivotTableSheetIdMap[oSlicerCachePivotTable.sheetId])
					this.slicerCachesExt[key].forCopySheet(oSlicerCachePivotTable.sheetId, this.slicerCachePivotTableSheetIdMap[oSlicerCachePivotTable.sheetId]);
				else
					this.slicerCachesExt[key].forCopySheet(oSlicerCachePivotTable.sheetId, null);
			}
		}
	};
	ReaderFromJSON.prototype.StylesFromJSON = function(oParsed, oStyleObject)
	{
		if (oParsed["borders"] != null)
		{
			for (var nElm = 0; nElm < oParsed["borders"].length; nElm++)
				oStyleObject.aBorders.push(this.BorderExcellFromJSON(oParsed["borders"][nElm]));
		}
		if (oParsed["fills"] != null)
		{
			for (var nElm = 0; nElm < oParsed["fills"].length; nElm++)
				oStyleObject.aFills.push(this.FillExcellFromJSON(oParsed["fills"][nElm]));
		}
		if (oParsed["fonts"] != null)
		{
			for (var nElm = 0; nElm < oParsed["fonts"].length; nElm++)
				oStyleObject.aFonts.push(this.FontExcellFromJSON(oParsed["fonts"][nElm]));
		}
		if (oParsed["numFmts"] != null)
		{
			for (var nElm = 0; nElm < oParsed["numFmts"].length; nElm++)
				this.NumFmtExcellFromJSON(oParsed["numFmts"][nElm], oStyleObject.oNumFmts);
		}
		if (oParsed["cellStyleXfs"] != null)
		{
			for (var nElm = 0; nElm < oParsed["cellStyleXfs"].length; nElm++)
				oStyleObject.aCellStyleXfs.push(this.XfsFromJSON(oParsed["cellStyleXfs"][nElm]));
		}
		if (oParsed["cellXfs"] != null)
		{
			for (var nElm = 0; nElm < oParsed["cellXfs"].length; nElm++)
				oStyleObject.aCellXfs.push(this.XfsFromJSON(oParsed["cellXfs"][nElm]));
		}
		if (oParsed["cellStyles"] != null)
		{
			for (var nElm = 0; nElm < oParsed["cellStyles"].length; nElm++)
				oStyleObject.aCellStyles.push(this.CellStyleFromJSON(oParsed["cellStyles"][nElm]));
		}
		if (oParsed["dxfs"] != null)
		{
			for (var nElm = 0; nElm < oParsed["dxfs"].length; nElm++)
				oStyleObject.aDxfs.push(this.DxfFromJSON(oParsed["dxfs"][nElm]));
		}
		if (oParsed["tableStyles"] != null)
		{
			this.TableStylesFromJSON(oParsed["tableStyles"], this.Workbook.TableStyles, oStyleObject.oCustomTableStyles);
		}
		if (oParsed["extDxfs"] != null)
		{
			for (var nElm = 0; nElm < oParsed["extDxfs"].length; nElm++)
				oStyleObject.aExtDxfs.push(this.DxfFromJSON(oParsed["extDxfs"][nElm]));
		}
		if (oParsed["slicerStyles"] != null)
			oStyleObject.oCustomSlicerStyles = this.SlicerStylesFromJSON(oParsed["slicerStyles"]);
	};
	ReaderFromJSON.prototype.SlicerStylesFromJSON = function(oParsed)
	{
		var oSlicerStyles = new Asc.CT_slicerStyles();

		if (oParsed["defaultSlicerStyle"] != null)
			oSlicerStyles.defaultSlicerStyle = oParsed["defaultSlicerStyle"];
		if (oParsed["slicerStyle"] != null)
		{
			for (var nElm = 0; nElm < oParsed["slicerStyle"].length; nElm++)
				oSlicerStyles.slicerStyle.push(this.SlicerStyleFromJSON(oParsed["slicerStyle"][nElm]));
		}

		return oSlicerStyles;
	};
	ReaderFromJSON.prototype.SlicerStyleFromJSON = function(oParsed)
	{
		var oSlicerStyle = new Asc.CT_slicerStyle();

		if (oParsed["name"] != null)
			oSlicerStyle.name = oParsed["name"];
		if (oParsed["slicerStyleElements"] != null)
		{
			for (var nElm = 0; nElm < oParsed["slicerStyleElements"].length; nElm++)
				oSlicerStyle.slicerStyleElements.push(this.SlicerStyleElementFromJSON(oParsed["slicerStyleElements"][nElm]));
		}

		return oSlicerStyle;
	};
	ReaderFromJSON.prototype.SlicerStyleElementFromJSON = function(oParsed)
	{
		var oSlicerStyleElm = new Asc.CT_slicerStyleElement();

		if (oParsed["type"] != null)
			oSlicerStyleElm.type = From_XML_ST_slicerStyleType(oParsed["type"]);
		if (oParsed["dxfId"] != null)
			oSlicerStyleElm.dxfId = oParsed["dxfId"];

		return oSlicerStyleElm;
	};
	ReaderFromJSON.prototype.ColsFromJSON = function(oParsedCols, oWorksheet)
	{
		var aTempCols = [];
		for (var nCol = 0; nCol < oParsedCols.length; nCol++)
			aTempCols.push(this.ColFromJSON(oParsedCols[nCol], oWorksheet));

		//если есть стиль последней колонки, назначаем его стилем всей таблицы и убираем из колонок
		var oAllCol = null;
		if(aTempCols.length > 0)
		{
			var oLast = aTempCols[aTempCols.length - 1];
			if(AscCommon.gc_nMaxCol == oLast.Max)
			{
				oAllCol = oWorksheet.getAllCol();
				oLast.col.cloneTo(oAllCol);
			}
		}
		
		for (var i = 0; i < aTempCols.length; ++i)
		{
			var elem = aTempCols[i];
			if(null != oAllCol && oAllCol.isEqual(elem.col))
				continue;
			if(elem.col.isUpdateScroll() && elem.Max >= oWorksheet.nColsCount)
				oWorksheet.nColsCount = elem.Max;

			for(var j = elem.Min; j <= elem.Max; j++){
				var oNewCol = new AscCommonExcel.Col(oWorksheet, j - 1);
				//var oOldProps = oNewCol.getWidthProp();
				elem.col.cloneTo(oNewCol);
				oWorksheet.aCols[oNewCol.index] = oNewCol;
				//oNewCol.xfs = null;
				//elem.col.xfs && oNewCol.setStyle(elem.col.xfs);
				//var oNewProps = oNewCol.getWidthProp();
				// if(false == oOldProps.isEqual(oNewProps))
				// 	History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ColProp, oWorksheet.getId(),
				// 		oNewCol._getUpdateRange(),
				// 		new AscCommonExcel.UndoRedoData_IndexSimpleProp(oNewCol.index, false, oOldProps, oNewProps));
			}
		}
	};
	ReaderFromJSON.prototype.ColFromJSON = function(oParsedCol, oWorksheet)
	{
		var oTempCol = {
			Max: null,
			Min: null,
			col: new AscCommonExcel.Col(oWorksheet || this.curWorksheet, 0)
		}

		if (oParsedCol["bestFit"] !== undefined)
			oTempCol.col.BestFit = oParsedCol["bestFit"];
		if (oParsedCol["hidden"] !== undefined)
			oTempCol.col.setHidden(oParsedCol["hidden"]);
		if (oParsedCol["max"] !== undefined)
			oTempCol.Max = oParsedCol["max"];
		if (oParsedCol["max"] !== undefined)
			oTempCol.Min = oParsedCol["min"];
		if (oParsedCol["width"] !== undefined)
			oTempCol.col.width = oParsedCol["width"];
		if (oParsedCol["customWidth"] !== undefined)
			oTempCol.col.CustomWidth = oParsedCol["customWidth"];
		if (oParsedCol["outlineLevel"] !== undefined)
			oTempCol.col.outlineLevel = oParsedCol["outlineLevel"];
		if (oParsedCol["collapsed"] !== undefined)
			oTempCol.col.collapsed = oParsedCol["collapsed"];
		if (oParsedCol["xfs"] != null)
			oTempCol.col.xfs = this.aCellXfs[oParsedCol["xfs"]];

		return oTempCol;
	};
	ReaderFromJSON.prototype.SheetFormatPrFromJSON = function(oParsedPr, oWorksheet)
	{
		if (!oParsedPr)
			return;

		var oAllRow = oWorksheet.getAllRow();

		if (oParsedPr["defaultColWidth"] != undefined)
			oWorksheet.oSheetFormatPr.dDefaultColWidth = oParsedPr["defaultColWidth"];
		if (oParsedPr["baseColWidth"] != undefined)
			oWorksheet.oSheetFormatPr.nBaseColWidth = oParsedPr["baseColWidth"];
		if (oParsedPr["defaultRowHeight"] != undefined)
			oAllRow.setHeight(oParsedPr["defaultRowHeight"]);
		if (oParsedPr["customHeight"] != undefined)
			oAllRow.setCustomHeight(oParsedPr["customHeight"]);
		if (oParsedPr["zeroHeight"] != undefined)
			oAllRow.setHidden(oParsedPr["zeroHeight"]);
		if (oParsedPr["outlineLevelCol"] != undefined)
			oWorksheet.oSheetFormatPr.nOutlineLevelCol = oParsedPr["outlineLevelCol"];
		if (oParsedPr["outlineLevelRow"] != undefined)
			oAllRow.setOutlineLevel(oParsedPr["outlineLevelRow"]);
	};
	ReaderFromJSON.prototype.PageMarginsFromJSON = function(oParsedMarg, oMargins)
	{
		if (!oParsedMarg)
			return;

		if (oParsedMarg["left"] != undefined)
			oMargins.asc_setLeft(oParsedMarg["left"]);
		if (oParsedMarg["top"] != undefined)
			oMargins.asc_setTop(oParsedMarg["top"]);
		if (oParsedMarg["right"] != undefined)
			oMargins.asc_setRight(oParsedMarg["right"]);
		if (oParsedMarg["bottom"] != undefined)
			oMargins.asc_setBottom(oParsedMarg["bottom"]);
		if (oParsedMarg["header"] != undefined)
			oMargins.asc_setHeader(oParsedMarg["header"]);
		if (oParsedMarg["footer"] != undefined)
			oMargins.asc_setFooter(oParsedMarg["footer"]);
	};
	ReaderFromJSON.prototype.PageSetupExcellFromJSON = function(oParsedSetup, oPageSetup)
	{
		if (!oPageSetup)
			return;

		var nOrientType = undefined;
		switch(oParsedSetup["orientation"])
		{
			case "portrait":
				nOrientType = Asc.EPageOrientation.pageorientPortrait;
				break;
			case "landscape":
				nOrientType = Asc.EPageOrientation.pageorientLandscape;
				break;
		}

		if (oParsedSetup["blackAndWhite"] != undefined)
			oPageSetup.blackAndWhite = oParsedSetup["blackAndWhite"];
		if (oParsedSetup["cellComments"] != undefined)
			oPageSetup.cellComments = FromXML_ST_CellComments(oParsedSetup["cellComments"]);
		if (oParsedSetup["copies"] != undefined)
			oPageSetup.copies = oParsedSetup["copies"];
		if (oParsedSetup["draft"] != undefined)
			oPageSetup.draft = oParsedSetup["draft"];
		if (oParsedSetup["errors"] != undefined)
			oPageSetup.errors = FromXML_ST_PrintError(oParsedSetup["errors"]);
		if (oParsedSetup["firstPageNumber"] != undefined)
			oPageSetup.firstPageNumber = oParsedSetup["firstPageNumber"];
		if (oParsedSetup["fitToHeight"] != undefined)
			oPageSetup.fitToHeight = oParsedSetup["fitToHeight"];
		if (oParsedSetup["fitToWidth"] != undefined)
			oPageSetup.fitToWidth = oParsedSetup["fitToWidth"];
		if (oParsedSetup["horizontalDpi"] != undefined)
			oPageSetup.horizontalDpi = oParsedSetup["horizontalDpi"];
		if (nOrientType != undefined)
			oPageSetup.asc_setOrientation(nOrientType);
		if (oParsedSetup["pageOrder"] != undefined)
			oPageSetup.pageOrder = FromXML_ST_PageOrder(oParsedSetup["pageOrder"]);
		if (oParsedSetup["paperSize"] != undefined)
		{
			var item = DocumentPageSize.getSizeById(oParsedSetup["paperSize"]);
			oPageSetup.asc_setWidth(item.w_mm);
			oPageSetup.asc_setHeight(item.h_mm);
		}
		if (oParsedSetup["scale"] != undefined)
			oPageSetup.scale = oParsedSetup["scale"];
		if (oParsedSetup["useFirstPageNumber"] != undefined)
			oPageSetup.useFirstPageNumber = oParsedSetup["useFirstPageNumber"];
		if (oParsedSetup["usePrinterDefaults"] != undefined)
			oPageSetup.usePrinterDefaults = oParsedSetup["usePrinterDefaults"];
		if (oParsedSetup["verticalDpi"] != undefined)
			oPageSetup.verticalDpi = oParsedSetup["verticalDpi"];
	};
	ReaderFromJSON.prototype.PrintOptionsExcelFromJSON = function(oParsedPr, oPrintOptions)
	{
		if (!oParsedPr)
			return;

		if (oParsedPr["gridLines"] != undefined)
			oPrintOptions.asc_setGridLines(oParsedPr["gridLines"]);
		if (oParsedPr["headings"] != undefined)
			oPrintOptions.asc_setHeadings(oParsedPr["headings"]);
	};
	ReaderFromJSON.prototype.HyperlinksExcelFromJSON = function(aParsedLinks, oWorksheet) // to do
	{
		var oHyperlink;
		var aHyperlinks;

		for (var nLink = 0; nLink < aParsedLinks.length; nLink++)
		{
			oHyperlink = this.HyperlinkExcelFromJSON(aParsedLinks[nLink], oWorksheet);
			aHyperlinks = oWorksheet.hyperlinkManager.get(oHyperlink.Ref.bbox);
			//удаляем ссылки с тем же адресом
			for(var i = 0, length = aHyperlinks.all.length; i < length; i++)
			{
				var hyp = aHyperlinks.all[i];
				if(hyp.bbox.isEqual(oHyperlink.Ref.bbox))
					oWorksheet.hyperlinkManager.removeElement(hyp);
			}
			oHyperlink.Ref && oWorksheet.hyperlinkManager.add(oHyperlink.Ref.bbox, oHyperlink);
		}
	};
	ReaderFromJSON.prototype.HyperlinkExcelFromJSON = function(oParsedLink, oWorksheet)
	{
		var oHyperlink = new AscCommonExcel.Hyperlink();
		if (oParsedLink["ref"] != null)
			oHyperlink.Ref = oWorksheet.getRange2(oParsedLink["ref"]);
		if (oParsedLink["display"] != null)
			oHyperlink.Hyperlink = oParsedLink["display"];
		if (oParsedLink["location"] != null)
			oHyperlink.setLocation(oParsedLink["location"]);
		if (oParsedLink["tooltip"] != null)
			oHyperlink.Tooltip = oParsedLink["tooltip"];

		return oHyperlink;
	};
	ReaderFromJSON.prototype.MergeCellsFromJSON = function(aParsedCells, oWorksheet)
	{
		var oRange;
		for (var nMerge = 0; nMerge < aParsedCells.length; nMerge++)
		{
			oRange = oWorksheet.getRange2(aParsedCells[nMerge]);
			oWorksheet.mergeManager.add(oRange.bbox, 1);
		}
	};
	ReaderFromJSON.prototype.DrawingsExcellFromJSON = function(aParsed, oWorksheet)
	{
		for (var nDrawing = 0; nDrawing < aParsed.length; nDrawing++)
			oWorksheet.Drawings.push(this.DrawingExcellFromJSON(aParsed[nDrawing], oWorksheet));
	};
	ReaderFromJSON.prototype.DrawingExcellFromJSON = function(oParsed, oWorksheet)
	{
		var objectRender = new AscFormat.DrawingObjects();
		var oFlags = {from: false, to: false, pos: false, ext: false, editAs: AscCommon.c_oAscCellAnchorType.cellanchorTwoCell};
		var oDrawing = objectRender.createDrawingObject();
		var oFrom = {}, oTo = {}, oPos = {}, oExt = {};

		var oGraphicObj = this.GraphicObjFromJSON(oParsed["graphic"]);
		oDrawing.graphicObject = oGraphicObj;
		oGraphicObj.setDrawingBase(oDrawing);

		if(typeof oGraphicObj.setWorksheet != "undefined")
			oGraphicObj.setWorksheet(oWorksheet);
		oGraphicObj.setDrawingObjects(objectRender);
		//History.Add(new AscFormat.CChangesDrawingObjectsAddToDrawingObjects(oGraphicObj, oWorksheet.Drawings.length));

		if (oParsed["type"] != null)
			oDrawing.Type = FromXML_ST_EditAs(oParsed["type"]);
		if (oParsed["editAs"] != null)
			oFlags.editAs = FromXML_ST_EditAs(oParsed["editAs"]);
		if (oParsed["from"] != null)
		{
			oFlags.from = true;
			this.FromToObjFromJSON(oParsed["from"], oFrom);
		}
		else
			oFrom = oDrawing.from;

		if (oParsed["to"] != null)
		{
			oFlags.to = true;
			this.FromToObjFromJSON(oParsed["to"], oTo);
		}
		else
			oTo = oDrawing.to;
		
		if (oParsed["pos"] != null)
		{
			oPos = {
				x: oParsed["pos"]["x"],
				y: oParsed["pos"]["y"]
			}
		}
		else
			oPos = {
				x: oDrawing.Pos.X,
				y: oDrawing.Pos.Y,
			};

		if (oParsed["ext"] != null)
		{
			oExt = {
				cx: private_EMU2MM(oParsed["ext"]["cx"]),
				cy: private_EMU2MM(oParsed["ext"]["cy"])
			}
		}

		oGraphicObj.setDrawingBaseCoords(oFrom.col, oFrom.colOff, oFrom.row, oFrom.rowOff, oTo.col, oTo.colOff, oTo.row, oTo.rowOff, oPos.x, oPos.y, oExt.cx, oExt.cy);
		
		if (oParsed["clientData"])
			oGraphicObj.setClientData(this.ClientDataFromJSON(oParsed["clientData"]));

		if(false != oFlags.from && false != oFlags.to){
			oDrawing.Type = AscCommon.c_oAscCellAnchorType.cellanchorTwoCell;
			oDrawing.editAs = oFlags.editAs;
		} else if(false != oFlags.from && false != oFlags.ext)
			oDrawing.Type = AscCommon.c_oAscCellAnchorType.cellanchorOneCell;
		else if(false != oFlags.pos && false != oFlags.ext)
			oDrawing.Type = AscCommon.c_oAscCellAnchorType.cellanchorAbsolute;
		
		
		oGraphicObj.setDrawingBaseType && oGraphicObj.setDrawingBaseType(oDrawing.Type);
		oGraphicObj.setDrawingBaseEditAs && oGraphicObj.setDrawingBaseEditAs(oDrawing.editAs);

		if(!oGraphicObj.spPr)
		{
			oGraphicObj.setSpPr(new AscFormat.CSpPr());
			oGraphicObj.spPr.setParent(oGraphicObj);
		}
		if(!oGraphicObj.spPr.xfrm)
		{
			oGraphicObj.spPr.setXfrm(new AscFormat.CXfrm());
			oGraphicObj.spPr.xfrm.setParent(oGraphicObj.spPr);
			oGraphicObj.spPr.xfrm.setOffX(0);
			oGraphicObj.spPr.xfrm.setOffY(0);
			oGraphicObj.spPr.xfrm.setExtX(0);
			oGraphicObj.spPr.xfrm.setExtY(0);
		}

		return oDrawing;
	};
	ReaderFromJSON.prototype.FromToObjFromJSON = function(oParsed, oFromTo) // CCellObjectInfo
	{
		if (oFromTo.col < 0)
			oFromTo.col = 0;
		if (oFromTo.row < 0)
			oFromTo.row = 0;
		
		oFromTo.row = oParsed["row"];
		oFromTo.col = oParsed["col"];
		
		oFromTo.rowOff = private_EMU2MM(oParsed["rowOff"]);
		oFromTo.colOff = private_EMU2MM(oParsed["colOff"]);
	};
	ReaderFromJSON.prototype.ClientDataFromJSON = function(oParsed)
	{
		var oClientData = new AscFormat.CClientData();
		if (oParsed["fLocksWithSheet"] != null)
			oClientData.fLocksWithSheet = oParsed["fLocksWithSheet"];
		if (oParsed["fPrintsWithSheet"] != null)
			oClientData.fPrintsWithSheet = oParsed["fPrintsWithSheet"];

		return oClientData;
	};
	ReaderFromJSON.prototype.AutoFilterFromJSON = function(oParsed)
	{
		var oAutoFilter = new AscCommonExcel.AutoFilter();

		if (oParsed["ref"] != null)
			oAutoFilter.setStringRef(oParsed["ref"]);
		if (oParsed["filterColumn"] != null)
			oAutoFilter.FilterColumns = this.FilterColumnsFromJSON(oParsed["filterColumn"]);
		if (oParsed["sortState"] != null)
			oAutoFilter.SortState = this.SortStateFromJSON(oParsed["sortState"]);

		return oAutoFilter;
	};
	ReaderFromJSON.prototype.FilterColumnsFromJSON = function(aParsed)
	{
		var aResult = [];
		for (var nFilter = 0; nFilter < aParsed.length; nFilter++)
			aResult.push(this.FilterColumnFromJSON(aParsed[nFilter]));

		return aResult;
	};
	ReaderFromJSON.prototype.FilterColumnFromJSON = function(oParsed)
	{
		var oFilterColumn = new AscCommonExcel.FilterColumn();

		if (oParsed["colId"] != null)
			oFilterColumn.ColId = oParsed["colId"];
		if (oParsed["filters"] != null)
		{
			oFilterColumn.Filters = this.FiltersFromJSON(oParsed["filters"]);
			oFilterColumn.Filters.sortDate();
		}
		if (oParsed["customFilters"] != null)
		{
			oFilterColumn.CustomFiltersObj = this.CustomFiltersFromJSON(oParsed["customFilters"]);
		}
		if (oParsed["dynamicFilter"] != null)
		{
			oFilterColumn.DynamicFilter = this.DynamicFilterFromJSON(oParsed["dynamicFilter"]);
		}
		if (oParsed["colorFilter"] != null)
		{
			oFilterColumn.ColorFilter = this.ColorFilterFromJSON(oParsed["colorFilter"]);
		}
		if (oParsed["top10"] != null)
		{
			oFilterColumn.Top10 = this.Top10FromJSON(oParsed["top10"]);
		}
		if (oParsed["showButton"] != null)
		{
			oFilterColumn.ShowButton = oParsed["showButton"];
		}

		return oFilterColumn;
	};
	ReaderFromJSON.prototype.FiltersFromJSON = function(oParsed)
	{
		var oFilters = new AscCommonExcel.Filters();
		var oFilterVal, oDateGroupItem, oAutoFilterDateElem;

		for (var nFilter = 0; nFilter < oParsed["filter"].length; nFilter++)
		{
			oFilterVal = new AscCommonExcel.Filter();
			oFilterVal.Val = oParsed["filter"][nFilter];
			if(null != oFilterVal.Val)
				oFilters.Values[oFilterVal.Val] = 1;
		}

		for (var nItem = 0; nItem < oParsed["dateGroupItem"].length; nItem++)
		{
			oDateGroupItem = this.DateGroupItemFromJSON(oParsed["dateGroupItem"][nItem]);
			oAutoFilterDateElem = new AscCommonExcel.AutoFilterDateElem();
			oAutoFilterDateElem.convertDateGroupItemToRange(oDateGroupItem);
			oFilters.Dates.push(oAutoFilterDateElem);
		}
			
		if (oParsed["blank"] != null)
			oFilters.Blank = oParsed["blank"];

		return oFilters;
	};
	ReaderFromJSON.prototype.DateGroupItemFromJSON = function(oParsed)
	{
		var oDateGroupItem = new AscCommonExcel.DateGroupItem();
		if (oParsed["dateTimeGrouping"] != null)
			oDateGroupItem.DateTimeGrouping = FromXml_ST_DateTimeGrouping(oParsed["dateTimeGrouping"]);
		if (oParsed["day"] != null)
			oDateGroupItem.Day = oParsed["day"];
		if (oParsed["hour"] != null)
			oDateGroupItem.Hour = oParsed["hour"];
		if (oParsed["minute"] != null)
			oDateGroupItem.Minute = oParsed["minute"];
		if (oParsed["month"] != null)
			oDateGroupItem.Month = oParsed["month"];
		if (oParsed["second"] != null)
			oDateGroupItem.Second = oParsed["second"];
		if (oParsed["year"] != null)
			oDateGroupItem.Year = oParsed["year"];

		return oDateGroupItem;
	};
	ReaderFromJSON.prototype.CustomFiltersFromJSON = function(oParsed)
	{
		var oCustomFilters = new Asc.CustomFilters();

		if (oParsed["and"])
			oCustomFilters.And = oParsed["and"];
		if (oParsed["customFilters"].length > 0)
		{
			oCustomFilters.CustomFilters = [];
			for (var nItem = 0; nItem < oParsed["customFilters"].length; nItem++)
				oCustomFilters.CustomFilters.push(this.CustomFilterFromJSON(oParsed["customFilters"][nItem]));
		}
		
		return oCustomFilters;
	};
	ReaderFromJSON.prototype.CustomFilterFromJSON = function(oParsed)
	{
		var oCustomFilter = new Asc.CustomFilter();

		if (oParsed["operator"] != null)
			oCustomFilter.Operator = FromXml_ST_FilterOperator(oParsed["operator"]);
		if (oParsed["val"] != null)
			oCustomFilter.Val = oParsed["val"];

		return oCustomFilter;
	};
	ReaderFromJSON.prototype.DynamicFilterFromJSON = function(oParsed)
	{
		var oFilter = new Asc.DynamicFilter();
		if (oParsed["type"] != null)
			oFilter.Type = FromXml_ST_DynamicFilterType(oParsed["type"]);
		if (oParsed["maxValIso"] != null)
			oFilter.MaxVal = oParsed["maxValIso"];
		if (oParsed["valIso"] != null)
			oFilter.Val = oParsed["valIso"];

		return oFilter;
	};
	ReaderFromJSON.prototype.ColorFilterFromJSON = function(oParsed)
	{
		var oColorFilter = new Asc.ColorFilter();
		if (oParsed["cellColor"] != null)
			oColorFilter.CellColor = oParsed["cellColor"];
		if (oParsed["dxfId"] != null)
			oColorFilter.dxf = this.aDxfs[oParsed["dxfId"]];

		return oColorFilter;
	};
	ReaderFromJSON.prototype.XfsFromJSON = function(oParsed)
	{
		var oXfs = new Asc.OpenXf();

		if (oParsed["applyAlignment"] != null)
            oXfs.ApplyAlignment = oParsed["applyAlignment"];    
		if (oParsed["applyBorder"] != null)
            oXfs.ApplyBorder = oParsed["applyBorder"];
		if (oParsed["applyFill"] != null)
            oXfs.ApplyFill = oParsed["applyFill"];
		if (oParsed["applyFont"] != null)
            oXfs.ApplyFont = oParsed["applyFont"];
		if (oParsed["applyNumberFormat"] != null)
            oXfs.ApplyNumberFormat = oParsed["applyNumberFormat"];
        if (oParsed["borderId"] != null)
            oXfs.borderid = oParsed["borderId"];
        if (oParsed["fillId"] != null)
            oXfs.fillid = oParsed["fillId"];
        if (oParsed["fontId"] != null)
            oXfs.fontid = oParsed["fontId"];
        if (oParsed["numFmtId"] != null)
            oXfs.numid = oParsed["numFmtId"];
		if (oParsed["quotePrefix"] != null)
            oXfs.QuotePrefix = oParsed["quotePrefix"];
		if (oParsed["pivotButton"] != null)
            oXfs.PivotButton = oParsed["pivotButton"];
		if (oParsed["xfId"] != null)
            oXfs.XfId = oParsed["xfId"];
        if (oParsed["alignment"] != null)
            oXfs.align = this.AlignFromJSON(oParsed["alignment"]);
        if (oParsed["applyProtection"] != null)
            oXfs.applyProtection = oParsed["applyProtection"];
		if (oParsed["protection"] != null)
		{
			if (oParsed["protection"]["hidden"] != null)
				oXfs.hidden = oParsed["protection"]["hidden"];
			if (oParsed["protection"]["locked"] != null)
				oXfs.locked = oParsed["protection"]["locked"];
		}
			
		return oXfs;
	};
	ReaderFromJSON.prototype.CellStyleFromJSON = function(oParsed)
	{
		var oCellStyle = new AscCommonExcel.CCellStyle();
		
		if (oParsed["builtinId"] != null)
            oCellStyle.BuiltinId = oParsed["builtinId"];        
        if (oParsed["customBuiltin"] != null)
            oCellStyle.CustomBuiltin = oParsed["customBuiltin"];
        if (oParsed["hidden"] != null)
            oCellStyle.Hidden = oParsed["hidden"];
        if (oParsed["iLevel"] != null)
            oCellStyle.ILevel = oParsed["iLevel"];
        if (oParsed["name"] != null)
            oCellStyle.Name = oParsed["name"];
        if (oParsed["xfId"] != null)
            oCellStyle.XfId = oParsed["xfId"];

		return oCellStyle;
	};
	ReaderFromJSON.prototype.TableStylesFromJSON = function(oParsed, oWbTableStyles, oCustomTableSyles)
	{
		if (oParsed["defaultTableStyle"] != null)
			oWbTableStyles.DefaultTableStyle = oParsed["defaultTableStyle"];
		if (oParsed["defaultPivotStyle"] != null)
			oWbTableStyles.DefaultPivotStyle = oParsed["defaultPivotStyle"];
		if (oParsed["customStyle"] != null)
			this.TableCustomStylesFromJSON(oParsed["customStyle"], oCustomTableSyles);
	};
	ReaderFromJSON.prototype.TableCustomStylesFromJSON = function(aParsed, oCustomTableSyles)
	{
		var oStyle, aElements;
		for (var nElm = 0; nElm < aParsed.length; nElm++)
		{
			aElements = [];
			oStyle = this.TableCustomStyleFromJSON(aParsed[nElm], aElements);
			
			if (oStyle.name != null)
			{
				if (oStyle.displayName === null)
					oStyle.displayName = oStyle.name;

				oCustomTableSyles[oStyle.name] = {style: oStyle, elements: aElements};
			}
		}
	};
	ReaderFromJSON.prototype.TableCustomStyleFromJSON = function(oParsed, aElements)
	{
		var oStyle = new Asc.CTableStyle();

		if (oParsed["name"] != null)
			oStyle.name = oParsed["name"];
		if (oParsed["pivot"] != null)
			oStyle.pivot = oParsed["pivot"];
		if (oParsed["table"] != null)
			oStyle.table = oParsed["table"];
		if (oParsed["elements"] != null)
		{
			for (var nElm = 0; nElm < oParsed["elements"].length; nElm++)
				aElements.push(this.TableCustomStyleElementFromJSON(oParsed["elements"][nElm]));
		}

		return oStyle;
	};
	ReaderFromJSON.prototype.TableCustomStyleElementFromJSON = function(oParsed)
	{
		var oStyleElement = {Type: null, Size: null, DxfId: null};

		if (oParsed["type"] != null)
			oStyleElement.Type = From_XML_ETableStyleType(oParsed["type"]);
		if (oParsed["size"] != null)
			oStyleElement.Size = oParsed["size"];
		if (oParsed["dxfId"] != null)
			oStyleElement.DxfId = oParsed["size"];

		return oStyleElement;
	};
	ReaderFromJSON.prototype.DxfFromJSON = function(oParsed)
	{
		var oDxf = new AscCommonExcel.CellXfs();

		if (oParsed["alignment"] != null)
            oDxf.align = this.AlignFromJSON(oParsed["alignment"]);
        if (oParsed["border"] != null)
            oDxf.border = this.BorderExcellFromJSON(oParsed["border"]);
        if (oParsed["fill"] != null)
		{
			oDxf.fill = this.FillExcellFromJSON(oParsed["fill"]);
			oDxf.fill.fixForDxf();
		}
        if (oParsed["font"] != null)
            oDxf.font = this.FontExcellFromJSON(oParsed["font"]);
        if (oParsed["numFmt"] != null)
            oDxf.num = this.NumFmtExcellFromJSON(oParsed["numFmt"]);

		return oDxf;
	};
	ReaderFromJSON.prototype.AlignFromJSON = function(oParsed)
	{
		var oAlgn = new AscCommonExcel.Align();

		if (oParsed["horizontal"] != null)
			oAlgn.hor = FromXml_ST_HorizontalAlignment(oParsed["horizontal"]);
		if (oParsed["indent"] != null)
			oAlgn.indent = oParsed["indent"];
		if (oParsed["relativeIndent"] != null)
			oAlgn.RelativeIndent = oParsed["relativeIndent"];
		if (oParsed["shrinkToFit"] != null)
			oAlgn.shrink = oParsed["shrinkToFit"];
		if (oParsed["textRotation"] != null)
			oAlgn.angle = oParsed["textRotation"];
		if (oParsed["vertical"] != null)
			oAlgn.ver = FromXml_ST_VerticalAlignment(oParsed["vertical"]);
		if (oParsed["wrapText"] != null)
			oAlgn.wrap = oParsed["wrapText"];
	
		return oAlgn;
	};
	ReaderFromJSON.prototype.BorderExcellFromJSON = function(oParsed)
	{
		var oBorder = new AscCommonExcel.Border();
		if (oParsed["bottom"] != null) {
			oBorder.b = new AscCommonExcel.BorderProp();
			this.BorderPropFromJSON(oParsed["bottom"], oBorder.b);
		}
		if (oParsed["diagonal"] != null) {
			oBorder.d = new AscCommonExcel.BorderProp();
			this.BorderPropFromJSON(oParsed["diagonal"], oBorder.d);
		}
		if (oParsed["end"] != null) {
			oBorder.r = new AscCommonExcel.BorderProp();
			this.BorderPropFromJSON(oParsed["end"], oBorder.r);
		}
		if (oParsed["horizontal"] != null) {
			oBorder.ih = new AscCommonExcel.BorderProp();
			this.BorderPropFromJSON(oParsed["horizontal"], oBorder.ih);
		}
		if (oParsed["start"] != null) {
			oBorder.l = new AscCommonExcel.BorderProp();
			this.BorderPropFromJSON(oParsed["start"], oBorder.l);
		}
		if (oParsed["top"] != null) {
			oBorder.t = new AscCommonExcel.BorderProp();
			this.BorderPropFromJSON(oParsed["top"], oBorder.t);
		}
		if (oParsed["vertical"] != null) {
			oBorder.iv = new AscCommonExcel.BorderProp();
			this.BorderPropFromJSON(oParsed["vertical"], oBorder.iv);
		}
		if (oParsed["diagonalDown"] != null) {
			oBorder.dd = oParsed["diagonalDown"];
		}
		if (oParsed["diagonalUp"] != null) {
			oBorder.du = oParsed["diagonalUp"];
		}
		return oBorder;
	};
	ReaderFromJSON.prototype.BorderPropFromJSON = function(oParsed, oProp)
	{
		if (oParsed["style"] != null)
			oProp.setStyle(From_XML_EBorderStyle(oParsed["style"]));
		if (oParsed["color"] != null)
			oProp.c = this.ColorExcellFromJSON(oParsed["color"]);
	};
	ReaderFromJSON.prototype.ColorExcellFromJSON = function(oParsed)
	{
		if (oParsed["type"] === "themeClr" && oParsed["theme"] != null)
			return AscCommonExcel.g_oColorManager.getThemeColor(oParsed["theme"], oParsed["tint"]);
		else if (oParsed["type"] === "rgbClr" && oParsed["rgb"] != null)
			return new AscCommonExcel.RgbColor(0x00ffffff & oParsed["rgb"]);

		return null;
	};
	ReaderFromJSON.prototype.FillExcellFromJSON = function(oParsed)
	{
		var oFill = new AscCommonExcel.Fill();
		if (oParsed["patternFill"] != null)
			oFill.patternFill = this.PatternFillExcellFromJSON(oParsed["patternFill"]);
		if (oParsed["gradientFill"] != null)
			oFill.gradientFill = this.GradientFillExcellFromJSON(oParsed["gradientFill"]);

		return oFill;
	};
	ReaderFromJSON.prototype.PatternFillExcellFromJSON = function(oParsed)
	{
		var oPattenrFill = new AscCommonExcel.PatternFill();
		if (oParsed["patternType"] != null)
			oPattenrFill.patternType = FromXml_ST_PatternType(oParsed["patternType"]);
		if (oParsed["fgColor"] != null)
			oPattenrFill.fgColor = this.ColorExcellFromJSON(oParsed["fgColor"]);
		if (oParsed["bgColor"] != null)
			oPattenrFill.bgColor = this.ColorExcellFromJSON(oParsed["bgColor"]);

		return oPattenrFill;
	};	
	ReaderFromJSON.prototype.GradientFillExcellFromJSON = function(oParsed)
	{
		var oGradientFill = new AscCommonExcel.GradientFill();

		if (oParsed["type"] != null)
			oGradientFill.type = FromXml_ST_GradientType(oParsed["type"]);
		if (oParsed["left"] != null)
			oGradientFill.left = oParsed["left"];
		if (oParsed["top"] != null)
			oGradientFill.top = oParsed["top"];
		if (oParsed["right"] != null)
			oGradientFill.right = oParsed["right"];
		if (oParsed["bottom"] != null)
			oGradientFill.bottom = oParsed["bottom"];
		if (oParsed["degree"] != null)
			oGradientFill.degree = oParsed["degree"];
		if (oParsed["aGrStop"].length > 0)
		{
			for (var nGrStop = 0; nGrStop < oParsed["aGrStop"].length; nGrStop++)
			oGradientFill.stop.push(this.GradStopEcxellFromJSON(oParsed["aGrStop"][nGrStop]));
		}
		
		return oGradientFill;
	};
	ReaderFromJSON.prototype.GradStopEcxellFromJSON = function(oParsed)
	{
		var oGradientStop = new AscCommonExcel.GradientStop();
		if (oParsed["position"] != null)
			oGradientStop.position = oParsed["position"];
		if (oParsed["color"] != null)
			oGradientStop.color = this.ColorExcellFromJSON(oParsed["color"]);

		return oGradientStop;
	};
	ReaderFromJSON.prototype.FontExcellFromJSON = function(oParsed)
	{
		var oFont = new AscCommonExcel.Font();
		
		if (oParsed["b"] != null)
			oFont.b = oParsed["b"];
		if (oParsed["color"] != null)
			oFont.c = this.ColorExcellFromJSON(oParsed["color"]);
		if (oParsed["i"] != null)
			oFont.i = oParsed["i"];
		if (oParsed["name"] != null)
			oFont.fn = oParsed["name"];
		if (oParsed["strike"] != null)
			oFont.s = oParsed["strike"];
		if (oParsed["sz"] != null)
			oFont.fs = oParsed["sz"];
		if (oParsed["u"] != null)
			oFont.u = FromXML_EUnderline(oParsed["u"]);
		if (oParsed["vertAlign"] != null)
			oFont.va = FromXML_ST_VerticalAlignRun(oParsed["vertAlign"]);
		if (oParsed["scheme"] != null)
			oFont.scheme = FromXML_EFontScheme(oParsed["scheme"]);

		oFont.checkSchemeFont(this.Workbook.theme);

		return oFont;
	};
	ReaderFromJSON.prototype.NumFmtExcellFromJSON = function(oParsed, oNumFmts)
	{
		var oNumFmt = null;
		var oTempFmt = {f: null, id: null};
		if (oParsed["formatCode"] != null)
			oTempFmt.f = oParsed["formatCode"];
		if (oParsed["numFmtId"] != null)
			oTempFmt.id = oParsed["numFmtId"];

		if(null != oTempFmt.id)
			oNumFmt = this.ParseNum(oTempFmt, oNumFmts);

		return oNumFmt;
	};
	ReaderFromJSON.prototype.ParseNum = function(oNum, oNumFmts) // to do
	{
		var oRes = new AscCommonExcel.Num();
		var useNumId = false;
		if (null != oNum && null != oNum.f) {
			oRes.f = oNum.f;
		} else {
			var sStandartNumFormat = AscCommonExcel.aStandartNumFormats[oNum.id];
			if (null != sStandartNumFormat) {
				oRes.f = sStandartNumFormat;
			}
			if (null == oRes.f) {
				oRes.f = "General";
			}
			//format string is more priority then id. so, fill oRes.id only if format is empty
			useNumId = true;
		}
		if ((useNumId || this.useNumId) && AscCommon.canGetFormatByStandardId(oNum.id)) {
			oRes.id = oNum.id;
		}
		var numFormat = AscCommon.oNumFormatCache.get(oRes.f);
		numFormat.checkCultureInfoFontPicker();
		if (null != oNumFmts) {
			oNumFmts[oNum.id] = oRes;
		}
		return oRes;
	};
	ReaderFromJSON.prototype.Top10FromJSON = function(oParsed)
	{
		var oTop10 = new Asc.Top10();

		if (oParsed["filterVal"] != null)
			oTop10.FilterVal = oParsed["filterVal"];
		if (oParsed["percent"] != null)
			oTop10.Percent = oParsed["percent"];
		if (oParsed["Top"] != null)
			oTop10.top = oParsed["top"];
		if (oParsed["val"] != null)
			oTop10.Val = oParsed["val"];
		
		return oTop10;
	};
	ReaderFromJSON.prototype.SortStateFromJSON = function(oParsed, oParent)
	{
		var oSortState = new AscCommonExcel.SortState();

		if (oParsed["ref"] != null)
			oSortState.Ref = AscCommonExcel.g_oRangeCache.getAscRange(oParsed["ref"]);
		if (oParsed["caseSensitive"] != null)
			oSortState.CaseSensitive = oParsed["caseSensitive"];
		if (oParsed["columnSort"] != null)
			oSortState.ColumnSort = oParsed["columnSort"];
		if (oParsed["sortMethod"] != null)
			oSortState.SortMethod = FromXML_ST_SortMethod(oParsed["sortMethod"]);
		if (oParsed["sortCondition"].length > 0)
		{
			oSortState.SortConditions = [];
			for (var nSort = 0; nSort < oParsed["sortCondition"].length; nSort++)
				oSortState.SortConditions.push(this.SortConditionFromJSON(oParsed["sortCondition"][nSort]));
		}

		return oSortState;
	};
	ReaderFromJSON.prototype.SortConditionFromJSON = function(oParsed)
	{
		var oSortCond = new AscCommonExcel.SortCondition();
		if (oParsed["ref"] != null)
			oSortCond.Ref = AscCommonExcel.g_oRangeCache.getAscRange(oParsed["ref"]);
		if (oParsed["sortBy"] != null)
			oSortCond.ConditionSortBy = FromXML_ESortBy(oParsed["sortBy"]);
		if (oParsed["descending"] != null)
			oSortCond.ConditionDescending = oParsed["descending"];
		if (oParsed["dxfId"] != null)
			oSortCond.dxf = this.aDxfs[oParsed["dxfId"]];

		return oSortCond;
	};
	ReaderFromJSON.prototype.TablePartsFromJSON = function(aParsed, oWorksheet)
	{
		var oTable;
		for (var nTable = 0; nTable < aParsed.length; nTable++)
		{
			oTable = this.TablePartFromJSON(aParsed[nTable], oWorksheet);
			if(null != oTable.Ref && null != oTable.DisplayName) {
				oWorksheet.workbook.dependencyFormulas.addTableName(oWorksheet, oTable);
				oTable.buildDependencies();
			}
			
			oWorksheet.TableParts.push(oTable);
		}
		
		//return aResult;
	};
	ReaderFromJSON.prototype.TablePartFromJSON = function(oParsed, oWorksheet)
	{
		var oTable = oWorksheet.createTablePart();

		if (oParsed["ref"] != null)
			oTable.Ref = AscCommonExcel.g_oRangeCache.getAscRange(oParsed["ref"]);
		if (oParsed["headerRowCount"] != null)
			oTable.HeaderRowCount = oParsed["headerRowCount"];
		if (oParsed["totalsRowCount"] != null)
			oTable.TotalsRowCount = oParsed["totalsRowCount"];
		if (oParsed["displayName"] != null)
			oTable.DisplayName = oParsed["displayName"];
		if (oParsed["autoFilter"] != null)
		{
			oTable.AutoFilter = this.AutoFilterFromJSON(oParsed["autoFilter"], false);
			if(!oTable.AutoFilter.Ref) {
				oTable.AutoFilter.Ref = oTable.generateAutoFilterRef();
			}
		}
		if (oParsed["sortState"] != null)
			oTable.SortState = this.SortStateFromJSON(oParsed["sortState"]);
		if (oParsed["tableColumns"].length > 0)
			oTable.TableColumns = this.TableColumnsFromJSON(oParsed["tableColumns"]);
		if (oParsed["tableStyleInfo"] != null)
			oTable.TableStyleInfo = this.TableStyleInfoFromJSON(oParsed["tableStyleInfo"]);
		if (oParsed["altText"] != null)
			oTable.altText = oParsed["altText"];
		if (oParsed["altTextSummary"] != null)
			oTable.altTextSummary = oParsed["altTextSummary"];
		if (oParsed["id"] != null) // to do (похоже надо мапить)
			oTable.id = oParsed["id"];
		if (oParsed["queryTable"] != null)
			oTable.QueryTable = this.QueryTableFromJSON(oParsed["queryTable"]);
		if (oParsed["tableType"] != null)
			oTable.tableType = FromXML_ST_TableType(oParsed["tableType"]);

		//var oSettings = new AscCommonExcel.AddFormatTableOptions();
		//var oAddFormatTableOptionsRange = AscCommonExcel.g_oRangeCache.getAscRange(oParsed["ref"]).clone();
		//oAddFormatTableOptionsRange.r2 -= 1;

		//oSettings.setProperty(0, oAddFormatTableOptionsRange.getName());
		//oSettings.setProperty(1, false);

		// this.curWorksheet.autoFilters._addHistoryObj({Ref: oTable.Ref}, AscCH.historyitem_AutoFilter_Add, {
		// 	activeCells: oTable.Ref,
		// 	styleName: oTable.TableStyleInfo && oTable.TableStyleInfo.Name ? oTable.TableStyleInfo.Name : undefined,
		// 	addFormatTableOptionsObj: oSettings,
		// 	displayName: oTable.DisplayName ? oTable.DisplayName : null,
		// 	tablePart: undefined
		// }, null, oTable.Ref, undefined);

		return oTable;
	};
	ReaderFromJSON.prototype.TableColumnsFromJSON = function(aParsed)
	{
		var aTableCols = [];
		for (var nCol = 0; nCol < aParsed.length; nCol++)
			aTableCols.push(this.TableColumnFromJSON(aParsed[nCol]));

		return aTableCols;
	};
	ReaderFromJSON.prototype.TableColumnFromJSON = function(oParsed)
	{
		var oTableColumn = new AscCommonExcel.TableColumn();

		if (oParsed["name"] != null)
			oTableColumn.Name = oParsed["name"];
		if (oParsed["totalsRowLabel"] != null)
			oTableColumn.TotalsRowLabel = oParsed["totalsRowLabel"];
		if (oParsed["totalsRowFunction"] != null)
			oTableColumn.TotalsRowFunction = FromXML_ETotalsRowFunction(oParsed["totalsRowFunction"]);
		if (oParsed["totalsRowFormula"] != null)
			oTableColumn.TotalsRowFormula = oParsed["totalsRowFormula"]; // to do (need to check)
		if (oParsed["dataDxfId"] != null)
			oTableColumn.dxf = this.aDxfs[oParsed["dataDxfId"]];
		if (oParsed["queryTableFieldId"] != null)
			oTableColumn.queryTableFieldId = oParsed["queryTableFieldId"];
		if (oParsed["uniqueName"] != null)
			oTableColumn.uniqueName = oParsed["uniqueName"];
		if (oParsed["id"] != null)
			oTableColumn.id = oParsed["id"];

		return oTableColumn;
	};
	ReaderFromJSON.prototype.TableStyleInfoFromJSON = function(oParsed)
	{
		var oTableStyleInfo = new AscCommonExcel.TableStyleInfo();
		
		if (oParsed["name"] != null)
			oTableStyleInfo.Name = oParsed["name"];
		if (oParsed["showColumnStripes"] != null)
			oTableStyleInfo.ShowColumnStripes = oParsed["showColumnStripes"];
		if (oParsed["showRowStripes"] != null)
			oTableStyleInfo.ShowRowStripes = oParsed["showRowStripes"];
		if (oParsed["showFirstColumn"] != null)
			oTableStyleInfo.ShowFirstColumn = oParsed["showFirstColumn"];
		if (oParsed["showLastColumn"] != null)
			oTableStyleInfo.ShowLastColumn = oParsed["showLastColumn"];

		return oTableStyleInfo;
	};
	ReaderFromJSON.prototype.QueryTableFromJSON = function(oParsed)
	{
		var oQueryTable = new AscCommonExcel.QueryTable();

		if (oParsed["connectionId"] != null)
			oQueryTable.connectionId = oParsed["connectionId"];
		if (oParsed["name"] != null)
			oQueryTable.name = oParsed["name"];
		if (oParsed["autoFormatId"] != null)
			oQueryTable.autoFormatId = oParsed["autoFormatId"];
		if (oParsed["growShrinkType"] != null)
			oQueryTable.growShrinkType = oParsed["growShrinkType"];
		if (oParsed["adjustColumnWidth"] != null)
			oQueryTable.adjustColumnWidth = oParsed["adjustColumnWidth"];
		if (oParsed["applyAlignmentFormats"] != null)
			oQueryTable.applyAlignmentFormats = oParsed["applyAlignmentFormats"];
		if (oParsed["applyBorderFormats"] != null)
			oQueryTable.applyBorderFormats = oParsed["applyBorderFormats"];
		if (oParsed["applyFontFormats"] != null)
			oQueryTable.applyFontFormats = oParsed["applyFontFormats"];
		if (oParsed["applyNumberFormats"] != null)
			oQueryTable.applyNumberFormats = oParsed["applyNumberFormats"];
		if (oParsed["applyPatternFormats"] != null)
			oQueryTable.ApplyPatternFormats = oParsed["applyPatternFormats"];
		if (oParsed["applyWidthHeightFormats"] != null)
			oQueryTable.applyWidthHeightFormats = oParsed["applyWidthHeightFormats"];
		if (oParsed["backgroundRefresh"] != null)
			oQueryTable.backgroundRefresh = oParsed["backgroundRefresh"];
		if (oParsed["disableEdit"] != null)
			oQueryTable.disableEdit = oParsed["disableEdit"];
		if (oParsed["disableRefresh"] != null)
			oQueryTable.disableRefresh = oParsed["disableRefresh"];
		if (oParsed["fillFormulas"] != null)
			oQueryTable.fillFormulas = oParsed["fillFormulas"];
		if (oParsed["firstBackgroundRefresh"] != null)
			oQueryTable.firstBackgroundRefresh = oParsed["firstBackgroundRefresh"];
		if (oParsed["headers"] != null)
			oQueryTable.headers = oParsed["headers"];
		if (oParsed["intermediate"] != null)
			oQueryTable.intermediate = oParsed["intermediate"];
		if (oParsed["preserveFormatting"] != null)
			oQueryTable.preserveFormatting = oParsed["preserveFormatting"];
		if (oParsed["refreshOnLoad"] != null)
			oQueryTable.refreshOnLoad = oParsed["refreshOnLoad"];
		if (oParsed["removeDataOnSave"] != null)
			oQueryTable.removeDataOnSave = oParsed["removeDataOnSave"];
		if (oParsed["rowNumbers"] != null)
			oQueryTable.rowNumbers = oParsed["rowNumbers"];
		if (oParsed["queryTableRefresh"] != null)
			oQueryTable.queryTableRefresh = this.QueryTableRefreshFromJSON(oParsed["queryTableRefresh"]);

		return oQueryTable;
	};
	ReaderFromJSON.prototype.QueryTableRefreshFromJSON = function(oParsed)
	{
		var oQueryTableRefresh = new AscCommonExcel.QueryTableRefresh();

		if (oParsed["nextId"] != null)
			oQueryTableRefresh.nextId = oParsed["nextId"];
		if (oParsed["minimumVersion"] != null)
			oQueryTableRefresh.minimumVersion = oParsed["minimumVersion"];
		if (oParsed["unboundColumnsLeft"] != null)
			oQueryTableRefresh.unboundColumnsLeft = oParsed["unboundColumnsLeft"];
		if (oParsed["unboundColumnsRight"] != null)
			oQueryTableRefresh.unboundColumnsRight = oParsed["unboundColumnsRight"];
		if (oParsed["fieldIdWrapped"] != null)
			oQueryTableRefresh.fieldIdWrapped = oParsed["fieldIdWrapped"];
		if (oParsed["headersInLastRefresh"] != null)
			oQueryTableRefresh.headersInLastRefresh = oParsed["headersInLastRefresh"];
		if (oParsed["preserveSortFilterLayout"] != null)
			oQueryTableRefresh.preserveSortFilterLayout = oParsed["preserveSortFilterLayout"];
		if (oParsed["sortState"] != null)
			oQueryTableRefresh.sortState = this.SortStateFromJSON(oParsed["sortState"]);
		if (oParsed["queryTableFields"] != null)
			oQueryTableRefresh.queryTableFields = this.QueryTableFieldsFromJSON(oParsed["queryTableFields"]);
		if (oParsed["queryTableDeletedFields"] != null)
			oQueryTableRefresh.queryTableDeletedFields = this.QueryTableDeletedFieldsFromJSON(oParsed["queryTableDeletedFields"]);

		return oQueryTableRefresh;
	};
	ReaderFromJSON.prototype.QueryTableFieldsFromJSON = function(aParsed)
	{
		var aResult = [];
		for (var nFld = 0; nFld < aParsed.length; nFld++)
			aResult.push(this.QueryTableFieldFromJSON(aParsed[nFld]));

		return aResult;
	};
	ReaderFromJSON.prototype.QueryTableFieldFromJSON = function(oParsed)
	{
		var oQueryTableField = new AscCommonExcel.QueryTableField(true);

		if (oParsed["name"] != null)
			oQueryTableField.name = oParsed["name"];
		if (oParsed["id"] != null)
			oQueryTableField.id = oParsed["id"];
		if (oParsed["tableColumnId"] != null)
			oQueryTableField.tableColumnId = oParsed["tableColumnId"];
		if (oParsed["rowNumbers"] != null)
			oQueryTableField.rowNumbers = oParsed["rowNumbers"];
		if (oParsed["fillFormulas"] != null)
			oQueryTableField.fillFormulas = oParsed["fillFormulas"];
		if (oParsed["dataBound"] != null)
			oQueryTableField.dataBound = oParsed["dataBound"];
		if (oParsed["clipped"] != null)
			oQueryTableField.clipped = oParsed["clipped"];

		return oQueryTableField;
	};
	ReaderFromJSON.prototype.QueryTableDeletedFieldsFromJSON = function(aParsed)
	{
		var aResult = [];
		for (var nFld = 0; nFld < aParsed.length; nFld++)
			aResult.push(this.QueryTableDeletedFieldFromJSON(aParsed[nFld]));

		return aResult;
	};
	ReaderFromJSON.prototype.QueryTableDeletedFieldFromJSON = function(oParsed)
	{
		var oQueryTableDeletedField = new AscCommonExcel.QueryTableDeletedField();

		if (oParsed["name"] != null)
			oQueryTableDeletedField.name = oParsed["name"];

		return oQueryTableDeletedField;
	};
	ReaderFromJSON.prototype.CondFormattingFromJSON = function(aParsed, oWorksheet)
	{
		var oConditionalFormatting, oCFRule;
		for (var nCond = 0; nCond < aParsed.length; nCond++)
		{
			oConditionalFormatting = new AscCommonExcel.CConditionalFormatting();

			if (aParsed[nCond]["pivot"] != null)
				oConditionalFormatting.pivot = aParsed[nCond]["pivot"];
			if (aParsed[nCond]["sqref"] != null)
				oConditionalFormatting.setSqRef(aParsed[nCond]["sqref"]);
			
			oCFRule = this.CondFormattingRuleFromJSON(aParsed[nCond]);
			oConditionalFormatting.aRules.push(oCFRule);
			if (oConditionalFormatting.isValid())
			{
				oConditionalFormatting.initRules();
				oWorksheet.addCFRule(oCFRule, true);
			}
		}
	};
	ReaderFromJSON.prototype.CondFormattingRuleFromJSON = function(oParsed)
	{
		var oConditionalFormattingRule = new AscCommonExcel.CConditionalFormattingRule();

		if (oParsed["aboveAverage"] != null)
			oConditionalFormattingRule.aboveAverage = oParsed["aboveAverage"];
		if (oParsed["bottom"] != null)
			oConditionalFormattingRule.bottom = oParsed["bottom"];
		if (oParsed["dxf"] != null)
			oConditionalFormattingRule.dxf = this.DxfFromJSON(oParsed["dxf"]);
		if (oParsed["equalAverage"] != null)
			oConditionalFormattingRule.equalAverage = oParsed["equalAverage"];
		if (oParsed["operator"] != null)
			oConditionalFormattingRule.operator = FromXML_CFOperatorType(oParsed["operator"]);
		if (oParsed["percent"] != null)
			oConditionalFormattingRule.percent = oParsed["percent"];
		if (oParsed["priority"] != null)
			oConditionalFormattingRule.priority = oParsed["priority"];
		if (oParsed["rank"] != null)
			oConditionalFormattingRule.rank = oParsed["rank"];
		if (oParsed["stdDev"] != null)
			oConditionalFormattingRule.stdDev = oParsed["stdDev"];
		if (oParsed["stopIfTrue"] != null)
			oConditionalFormattingRule.stopIfTrue = oParsed["stopIfTrue"];
		if (oParsed["text"] != null)
			oConditionalFormattingRule.text = oParsed["text"];
		if (oParsed["timePeriod"] != null)
			oConditionalFormattingRule.timePeriod = oParsed["timePeriod"];
		if (oParsed["type"] != null)
			oConditionalFormattingRule.type = FromXML_CfRuleType(oParsed["type"]);
		if (oParsed["aboveAverage"] != null)
			oConditionalFormattingRule.aboveAverage = oParsed["aboveAverage"];
		if (oParsed["rules"].length > 0)
		{
			for (var nElm = 0; nElm < oParsed["rules"].length; nElm++)
				oConditionalFormattingRule.aRuleElements.push(this.CondFormRuleElementFromJSON(oParsed["rules"][nElm]));
		}

		return oConditionalFormattingRule;
	};
	ReaderFromJSON.prototype.CondFormRuleElementFromJSON = function(oParsed)
	{
		switch (oParsed["type"])
		{
			case "clrScale":
				return this.ColorScaleFromJSON(oParsed);
			case "dataBar":
				return this.DataBarFromJSON(oParsed);
			case "iconSet":
				return this.IconSetFromJSON(oParsed);
			case "formulaCf":
				return this.FormulaCFFromJSON(oParsed);
		}
	};
	ReaderFromJSON.prototype.ColorScaleFromJSON = function(oParsed)
	{
		var oColorScale = new AscCommonExcel.CColorScale();

		if (oParsed["cfvo"].length > 0)
		{
			for (var nElm = 0; nElm < oParsed["cfvo"].length; nElm++)
				oColorScale.aCFVOs.push(this.CondFmtValObjFromJSON(oParsed["cfvo"][nElm]));
		}
		if (oParsed["color"].length > 0)
		{
			for (var nElm = 0; nElm < oParsed["color"].length; nElm++)
				oColorScale.aColors.push(this.ColorExcellFromJSON(oParsed["color"][nElm]));
		}

		return oColorScale;
	};
	ReaderFromJSON.prototype.CondFmtValObjFromJSON = function(oParsed)
	{
		var oCFVO = new AscCommonExcel.CConditionalFormatValueObject();

		if (oParsed["gte"] != null)
			oCFVO.Gte = oParsed["gte"];
		if (oParsed["type"] != null)
			oCFVO.Type = FromXML_ST_CfvoType(oParsed["type"]);
		if (oParsed["val"] != null)
			oCFVO.Val = oParsed["val"];

		return oCFVO;
	};
	ReaderFromJSON.prototype.DataBarFromJSON = function(oParsed)
	{
		var oDataBar = new AscCommonExcel.CDataBar();

		if (oParsed["maxLength"] != null)
			oDataBar.MaxLength = oParsed["maxLength"];
		if (oParsed["minLength"] != null)
			oDataBar.MinLength = oParsed["minLength"];
		if (oParsed["showValue"] != null)
			oDataBar.ShowValue = oParsed["showValue"];
		if (oParsed["color"] != null)
			oDataBar.Color = this.ColorExcellFromJSON(oParsed["color"]);
		if (oParsed["negativeColor"] != null)
			oDataBar.NegativeColor = this.ColorExcellFromJSON(oParsed["negativeColor"]);
		if (oParsed["borderColor"] != null)
			oDataBar.BorderColor = this.ColorExcellFromJSON(oParsed["borderColor"]);
		if (oParsed["axisColor"] != null)
			oDataBar.AxisColor = this.ColorExcellFromJSON(oParsed["axisColor"]);
		if (!oDataBar.AxisColor)
			oDataBar.AxisColor = new AscCommonExcel.RgbColor(0);
		if (oParsed["negativeBorderColor"] != null)
			oDataBar.NegativeBorderColor = this.ColorExcellFromJSON(oParsed["negativeBorderColor"]);
		if (oParsed["axPos"] != null)
			oDataBar.AxisPosition = FromXML_EDataBarAxisPosition(oParsed["axPos"]);
		if (oParsed["dir"] != null)
			oDataBar.Direction = FromXML_EDataBarDirection(oParsed["dir"]);
		if (oParsed["gradient"] != null)
			oDataBar.Gradient = oParsed["gradient"];
		if (oParsed["negBarClrSameAsPositive"] != null)
			oDataBar.NegativeBarColorSameAsPositive = oParsed["negBarClrSameAsPositive"];
		if (oParsed["negBarBdrClrSameAsPositive"] != null)
			oDataBar.NegativeBarBorderColorSameAsPositive = oParsed["negBarBdrClrSameAsPositive"];

		if (oParsed["cfvo"].length > 0)
		{
			for (var nElm = 0; nElm < oParsed["cfvo"].length; nElm++)
				oDataBar.aCFVOs.push(this.CondFmtValObjFromJSON(oParsed["cfvo"][nElm]));
		}

		return oDataBar;
	};
	ReaderFromJSON.prototype.IconSetFromJSON = function(oParsed)
	{
		var oIconSet = new AscCommonExcel.CIconSet();

		if (oParsed["iconSet"] != null)
			oIconSet.IconSet = FromXML_IconSetType(oParsed["iconSet"]);
		if (oParsed["percent"] != null)
			oIconSet.Percent = oParsed["percent"];
		if (oParsed["reverse"] != null)
			oIconSet.Reverse = oParsed["reverse"];
		if (oParsed["showValue"] != null)
			oIconSet.ShowValue = oParsed["showValue"];
		if (oParsed["cfvo"].length > 0)
		{
			for (var nElm = 0; nElm < oParsed["cfvo"].length; nElm++)
				oIconSet.aCFVOs.push(this.CondFmtValObjFromJSON(oParsed["cfvo"][nElm]));
		}
		if (oParsed["cfIcon"].length > 0)
		{
			for (var nElm = 0; nElm < oParsed["cfIcon"].length; nElm++)
				oIconSet.aIconSets.push(this.CondFmtIconSetFromJSON(oParsed["cfIcon"][nElm]));
		}

		return oIconSet;
	};
	ReaderFromJSON.prototype.CondFmtIconSetFromJSON = function(oParsed)
	{
		var oCFIS = new AscCommonExcel.CConditionalFormatIconSet();

		if (oParsed["iconSet"] != null)
			oCFIS.IconSet = FromXML_IconSetType(oParsed["iconSet"]);
		if (oParsed["iconId"] != null)
			oCFIS.IconId = oParsed["iconId"];

		return oCFIS;
	};
	ReaderFromJSON.prototype.FormulaCFFromJSON = function(oParsed)
	{
		var oFormula = new AscCommonExcel.CFormulaCF();
		if (oParsed["formula"] != null)
			oFormula.Text = oParsed["formula"];
		
		return oFormula;
	};
	ReaderFromJSON.prototype.SheetViewsFromJSON = function(oParsed, oWorksheet)
	{
		// read only first sheet view
		var aResult = [];
		oParsed[0] != null && aResult.push(this.SheetViewFromJSON(oParsed[0], oWorksheet));

		return aResult;
	};
	ReaderFromJSON.prototype.SheetViewFromJSON = function(oParsed, oWorksheet)
	{
		var oSheetView = new AscCommonExcel.asc_CSheetViewSettings();

		if (oParsed["showGridLines"] != null)
			oSheetView.showGridLines = oParsed["showGridLines"];
		if (oParsed["showRowColHeaders"] != null)
			oSheetView.showRowColHeaders = oParsed["showRowColHeaders"];
		if (oParsed["showZeros"] != null)
			oSheetView.showZeros = oParsed["showZeros"];
		if (oParsed["topLeftCell"] != null)
		{
			var _topLeftCell = AscCommonExcel.g_oRangeCache.getAscRange(oParsed["topLeftCell"]);
			if (_topLeftCell)
				oSheetView.topLeftCell = new Asc.Range(_topLeftCell.c1, _topLeftCell.r1, _topLeftCell.c1, _topLeftCell.r1);
		}
		if (oParsed["zoomScale"] != null)
			oSheetView.zoomScale = oParsed["zoomScale"];
		if (oParsed["pane"] != null)
		{
			oSheetView.pane = this.PaneFromJSON(oParsed["pane"]);
			oSheetView.pane.init();
		}
		if (oParsed["selection"] != null)
		{
			oWorksheet.selectionRange.clean();
			this.SelectionRangeFromJSON(oParsed["selection"], oWorksheet.selectionRange);
			oWorksheet.selectionRange.update();
		}

		return oSheetView;
	};
	ReaderFromJSON.prototype.PaneFromJSON = function(oParsed)
	{
		var oPane = new AscCommonExcel.asc_CPane();

		// if (oParsed["activePane"] != null)
		// 	oPane.activePane = oParsed["activePane"];
		if (oParsed["state"] != null)
			oPane.state = oParsed["state"];
		if (oParsed["topLeftCell"] != null)
			oPane.topLeftCell = oParsed["topLeftCell"];
		if (oParsed["xSplit"] != null)
			oPane.xSplit = oParsed["xSplit"];
		if (oParsed["ySplit"] != null)
			oPane.ySplit = oParsed["ySplit"];

		oPane.init();
		// var oData = new AscCommonExcel.UndoRedoData_FromTo(new AscCommonExcel.UndoRedoData_FrozenBBox(new Asc.Range(0, 0, 0, 0)), new AscCommonExcel.UndoRedoData_FrozenBBox(new Asc.Range(oPane.xSplit, oPane.ySplit, oPane.xSplit, oPane.ySplit)), null);
		// History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ChangeFrozenCell,
		// 	this.curWorksheet.getId(), null, oData);
		return oPane;
	};
	ReaderFromJSON.prototype.SelectionRangeFromJSON = function(oParsed, oSelectionRange)
	{
		if (oParsed["activeCell"] != null)
		{
			var activeCell = AscCommonExcel.g_oRangeCache.getAscRange(oParsed["activeCell"]);
			if (activeCell)
				oSelectionRange.activeCell = new AscCommon.CellBase(activeCell.r1, activeCell.c1);
		}
		if (oParsed["activeCellId"] != null)
			oSelectionRange.activeCellId = oParsed["activeCellId"];
		if (oParsed["sqref"] != null)
		{
			var selectionNew = AscCommonExcel.g_oRangeCache.getRangesFromSqRef(oParsed["sqref"]);
			if (selectionNew.length > 0)
				oSelectionRange.ranges = selectionNew;
		}
	};
	ReaderFromJSON.prototype.SheetPrFromJSON = function(oParsed)
	{
		var oSheetPr = new AscCommonExcel.asc_CSheetPr();

		if (oParsed["codeName"] != null)
			oSheetPr.CodeName = oParsed["codeName"];
		if (oParsed["enableFormatConditionsCalculation"] != null)
			oSheetPr.EnableFormatConditionsCalculation = oParsed["enableFormatConditionsCalculation"];
		if (oParsed["filterMode"] != null)
			oSheetPr.FilterMode = oParsed["filterMode"];
		if (oParsed["published"] != null)
			oSheetPr.Published = oParsed["published"];
		if (oParsed["syncHorizontal"] != null)
			oSheetPr.SyncHorizontal = oParsed["syncHorizontal"];
		if (oParsed["syncRef"] != null)
			oSheetPr.SyncRef = oParsed["syncRef"];
		if (oParsed["syncVertical"] != null)
			oSheetPr.SyncVertical = oParsed["syncVertical"];
		if (oParsed["transitionEntry"] != null)
			oSheetPr.TransitionEntry = oParsed["transitionEntry"];
		if (oParsed["transitionEvaluation"] != null)
			oSheetPr.TransitionEvaluation = oParsed["transitionEvaluation"];
		if (oParsed["tabColor"] != null)
			oSheetPr.TabColor = this.ColorExcellFromJSON(oParsed["tabColor"]);
		// pageSetUpPr
		if (oParsed["pageSetUpPr"]["autoPageBreaks"] != null)
			oSheetPr.AutoPageBreaks = oParsed["pageSetUpPr"]["autoPageBreaks"];
		if (oParsed["pageSetUpPr"]["fitToPage"] != null)
			oSheetPr.FitToPage = oParsed["pageSetUpPr"]["fitToPage"];
		// outlinePr
		if (oParsed["outlinePr"]["applyStyles"] != null)
			oSheetPr.ApplyStyles = oParsed["outlinePr"]["applyStyles"];
		if (oParsed["outlinePr"]["showOutlineSymbols"] != null)
			oSheetPr.ShowOutlineSymbols = oParsed["outlinePr"]["showOutlineSymbols"];
		if (oParsed["outlinePr"]["summaryBelow"] != null)
			oSheetPr.SummaryBelow = oParsed["outlinePr"]["summaryBelow"];
		if (oParsed["outlinePr"]["summaryRight"] != null)
			oSheetPr.SummaryRight = oParsed["outlinePr"]["summaryRight"];

		return oSheetPr;
	};
	ReaderFromJSON.prototype.SparklineGroupsFromJSON = function(aParsed)
	{
		var aResult = [];
		for (var nElm = 0; nElm < aParsed.length; nElm++)
			aResult.push(this.SparklineGroupFromJSON(aParsed[nElm]));
			
		return aResult;
	};
	ReaderFromJSON.prototype.SparklineGroupFromJSON = function(oParsed)
	{
		var oSparklineGroup = new AscCommonExcel.sparklineGroup(true);

		if (oParsed["manualMax"] != null)
			oSparklineGroup.manualMax = oParsed["manualMax"];
		if (oParsed["manualMin"] != null)
			oSparklineGroup.manualMin = oParsed["manualMin"];
		if (oParsed["lineWeight"] != null)
			oSparklineGroup.lineWeight = oParsed["lineWeight"];
		if (oParsed["type"] != null)
			oSparklineGroup.type = FromXML_ST_SparklineType(oParsed["type"]);
		if (oParsed["dateAxis"] != null)
			oSparklineGroup.dateAxis = oParsed["dateAxis"];
		if (oParsed["displayEmptyCellsAs"] != null)
			oSparklineGroup.displayEmptyCellsAs = FromXML_ST_DispBlanksAs(oParsed["displayEmptyCellsAs"]);
		if (oParsed["markers"] != null)
			oSparklineGroup.markers = oParsed["markers"];
		if (oParsed["high"] != null)
			oSparklineGroup.high = oParsed["high"];
		if (oParsed["low"] != null)
			oSparklineGroup.low = oParsed["low"];
		if (oParsed["first"] != null)
			oSparklineGroup.first = oParsed["first"];
		if (oParsed["last"] != null)
			oSparklineGroup.last = oParsed["last"];
		if (oParsed["negative"] != null)
			oSparklineGroup.negative = oParsed["negative"];
		if (oParsed["displayXAxis"] != null)
			oSparklineGroup.displayXAxis = oParsed["displayXAxis"];
		if (oParsed["displayHidden"] != null)
			oSparklineGroup.displayHidden = oParsed["displayHidden"];
		if (oParsed["minAxisType"] != null)
			oSparklineGroup.minAxisType = FromXML_ST_SparklineAxisMinMax(oParsed["minAxisType"]);
		if (oParsed["maxAxisType"] != null)
			oSparklineGroup.maxAxisType = FromXML_ST_SparklineAxisMinMax(oParsed["maxAxisType"]);
		if (oParsed["rightToLeft"] != null)
			oSparklineGroup.rightToLeft = oParsed["rightToLeft"];
		if (oParsed["colorSeries"] != null)
			oSparklineGroup.colorSeries = this.ColorExcellFromJSON(oParsed["colorSeries"]);
		if (oParsed["colorNegative"] != null)
			oSparklineGroup.colorNegative = this.ColorExcellFromJSON(oParsed["colorNegative"]);
		if (oParsed["colorAxis"] != null)
			oSparklineGroup.colorAxis = this.ColorExcellFromJSON(oParsed["colorAxis"]);
		if (oParsed["colorMarkers"] != null)
			oSparklineGroup.colorMarkers = this.ColorExcellFromJSON(oParsed["colorMarkers"]);
		if (oParsed["colorFirst"] != null)
			oSparklineGroup.colorFirst = this.ColorExcellFromJSON(oParsed["colorFirst"]);
		if (oParsed["colorLast"] != null)
			oSparklineGroup.colorLast = this.ColorExcellFromJSON(oParsed["colorLast"]);
		if (oParsed["colorHigh"] != null)
			oSparklineGroup.colorHigh = this.ColorExcellFromJSON(oParsed["colorHigh"]);
		if (oParsed["colorLow"] != null)
			oSparklineGroup.colorLow = this.ColorExcellFromJSON(oParsed["colorLow"]);
		if (oParsed["f"] != null)
			oSparklineGroup.f = oParsed["f"];
		if (oParsed["sparklines"].length > 0)
		{
			for (var nElm = 0; nElm < oParsed["sparklines"].length; nElm++)
				oSparklineGroup.arrSparklines.push(this.SparkLineFromJSON(oParsed["sparklines"][nElm]));
		}
		
		return oSparklineGroup;
	};
	ReaderFromJSON.prototype.SparkLineFromJSON = function(oParsed)
	{
		var oSparkline = new AscCommonExcel.sparkline();

		if (oParsed["f"] != null)
			oSparkline.setF(oParsed["f"]);
		if (oParsed["sqRef"] != null)
			oSparkline.setSqRef(oParsed["sqRef"]);

		return oSparkline;
	};
	ReaderFromJSON.prototype.HdrFtrExcellFromJSON = function(oParsed, oHdrFtr)
	{
		if (oParsed["alignWithMargins"] != null)
			oHdrFtr.setAlignWithMargins(oParsed["alignWithMargins"]);
		if (oParsed["differentFirst"] != null)
			oHdrFtr.setDifferentFirst(oParsed["differentFirst"]);
		if (oParsed["differentOddEven"] != null)
			oHdrFtr.setDifferentOddEven(oParsed["differentOddEven"]);
		if (oParsed["scaleWithDoc"] != null)
			oHdrFtr.setScaleWithDoc(oParsed["scaleWithDoc"]);
		if (oParsed["evenFooter"] != null)
			oHdrFtr.setEvenFooter(oParsed["evenFooter"]);
		if (oParsed["evenHeader"] != null)
			oHdrFtr.setEvenHeader(oParsed["evenHeader"]);
		if (oParsed["firstFooter"] != null)
			oHdrFtr.setFirstFooter(oParsed["firstFooter"]);
		if (oParsed["firstHeader"] != null)
			oHdrFtr.setFirstHeader(oParsed["firstHeader"]);
		if (oParsed["oddFooter"] != null)
			oHdrFtr.setOddFooter(oParsed["oddFooter"]);
		if (oParsed["oddHeader"] != null)
			oHdrFtr.setOddHeader(oParsed["oddHeader"]);
	};
	ReaderFromJSON.prototype.DataValidationsFromJSON = function(oParsed)
	{
		var oDataValidation = new AscCommonExcel.CDataValidations();

		if (oParsed["disablePrompts"] != null)
			oDataValidation.disablePrompts = oParsed["disablePrompts"];
		if (oParsed["xWindow"] != null)
			oDataValidation.xWindow = oParsed["xWindow"];
		if (oParsed["yWindow"] != null)
			oDataValidation.yWindow = oParsed["yWindow"];
		if (oParsed["dataValidation"].length > 0)
		{
			for (var nElm = 0; nElm < oParsed["dataValidation"].length; nElm++)
				oDataValidation.elems.push(this.DataValidationFromJSON(oParsed["dataValidation"][nElm]));
		}
		
		return oDataValidation;
	};
	ReaderFromJSON.prototype.DataValidationFromJSON = function(oParsed)
	{
		var oDataValidation = new AscCommonExcel.CDataValidation();

		if (oParsed["allowBlank"] != null)
			oDataValidation.allowBlank = oParsed["allowBlank"];
		if (oParsed["type"] != null)
			oDataValidation.type = FromXML_ST_DataValidationType(oParsed["type"]);
		if (oParsed["error"] != null)
			oDataValidation.error = oParsed["error"];
		if (oParsed["errorTitle"] != null)
			oDataValidation.errorTitle = oParsed["errorTitle"];
		if (oParsed["errorStyle"] != null)
			oDataValidation.errorStyle = FromXML_ST_DataValidationErrorStyle(oParsed["errorStyle"]);
		if (oParsed["imeMode"] != null)
			oDataValidation.imeMode = FromXML_ST_DataValidationImeMode(oParsed["imeMode"]);
		if (oParsed["operator"] != null)
			oDataValidation.operator = FromXML_ST_DataValidationOperator(oParsed["operator"]);
		if (oParsed["prompt"] != null)
			oDataValidation.prompt = oParsed["prompt"];
		if (oParsed["promptTitle"] != null)
			oDataValidation.promptTitle = oParsed["promptTitle"];
		if (oParsed["showDropDown"] != null)
			oDataValidation.showDropDown = oParsed["showDropDown"];
		if (oParsed["showErrorMessage"] != null)
			oDataValidation.showErrorMessage = oParsed["showErrorMessage"];
		if (oParsed["showInputMessage"] != null)
			oDataValidation.showInputMessage = oParsed["showInputMessage"];
		if (oParsed["sqref"] != null)
			oDataValidation.setSqRef(oParsed["sqref"]);
		if (oParsed["formula1"] != null)
			oDataValidation.formula1 = new Asc.CDataFormula(oParsed["formula1"]);
		if (oParsed["formula2"] != null)
			oDataValidation.formula2 = new Asc.CDataFormula(oParsed["formula2"]);

		return oDataValidation;
	};
	ReaderFromJSON.prototype.PivotTablesFromJSON = function(aParsed, oWorksheet)
	{
		var oPivotTable;
		for (var nTable = 0; nTable < aParsed.length; nTable++)
		{
			oPivotTable = this.PivotTableFromJSON(aParsed[nTable]);
			oWorksheet.insertPivotTable(oPivotTable, true);
		}
	};
	ReaderFromJSON.prototype.PivotTableFromJSON = function(oParsed)
	{
		var oPivotTable = new Asc.CT_pivotTableDefinition(true);

		if (oParsed["name"] != null)
			oPivotTable.name = oParsed["name"];
		if (oParsed["cacheId"] != null)
		{
			oPivotTable.cacheId = oParsed["cacheId"];
			for (var key in this.pivotCaches)
			{
				if (this.pivotCaches[key].cacheSource && this.pivotCaches[key].cacheSource.connectionId === oPivotTable.cacheId)
					oPivotTable.cacheDefinition = this.pivotCaches[key];
			}
		}
		if (oParsed["dataOnRows"] != null)
			oPivotTable.dataOnRows = oParsed["dataOnRows"];
		if (oParsed["dataPosition"] != null)
			oPivotTable.dataPosition = oParsed["dataPosition"];
		if (oParsed["autoFormatId"] != null)
			oPivotTable.autoFormatId = oParsed["autoFormatId"];
		if (oParsed["applyNumberFormats"] != null)
			oPivotTable.applyNumberFormats = oParsed["applyNumberFormats"];
		if (oParsed["applyBorderFormats"] != null)
			oPivotTable.applyBorderFormats = oParsed["applyBorderFormats"];
		if (oParsed["applyFontFormats"] != null)
			oPivotTable.applyFontFormats = oParsed["applyFontFormats"];
		if (oParsed["applyPatternFormats"] != null)
			oPivotTable.applyPatternFormats = oParsed["applyPatternFormats"];
		if (oParsed["applyAlignmentFormats"] != null)
			oPivotTable.applyAlignmentFormats = oParsed["applyAlignmentFormats"];
		if (oParsed["applyWidthHeightFormats"] != null)
			oPivotTable.applyWidthHeightFormats = oParsed["applyWidthHeightFormats"];
		if (oParsed["dataCaption"] != null)
			oPivotTable.dataCaption = oParsed["dataCaption"];
		if (oParsed["grandTotalCaption"] != null)
			oPivotTable.grandTotalCaption = oParsed["grandTotalCaption"];
		if (oParsed["errorCaption"] != null)
			oPivotTable.errorCaption = oParsed["errorCaption"];
		if (oParsed["showError"] != null)
			oPivotTable.showError = oParsed["showError"];
		if (oParsed["missingCaption"] != null)
			oPivotTable.missingCaption = oParsed["missingCaption"];
		if (oParsed["showMissing"] != null)
			oPivotTable.showMissing = oParsed["showMissing"];
		if (oParsed["pageStyle"] != null)
			oPivotTable.pageStyle = oParsed["pageStyle"];
		if (oParsed["pivotTableStyle"] != null)
			oPivotTable.pivotTableStyle = oParsed["pivotTableStyle"];
		if (oParsed["vacatedStyle"] != null)
			oPivotTable.vacatedStyle = oParsed["vacatedStyle"];
		if (oParsed["tag"] != null)
			oPivotTable.tag = oParsed["tag"];
		if (oParsed["updatedVersion"] != null)
			oPivotTable.updatedVersion = oParsed["updatedVersion"];
		if (oParsed["minRefreshableVersion"] != null)
			oPivotTable.minRefreshableVersion = oParsed["minRefreshableVersion"];
		if (oParsed["asteriskTotals"] != null)
			oPivotTable.asteriskTotals = oParsed["asteriskTotals"];
		if (oParsed["showItems"] != null)
			oPivotTable.showItems = oParsed["showItems"];
		if (oParsed["editData"] != null)
			oPivotTable.editData = oParsed["editData"];
		if (oParsed["disableFieldList"] != null)
			oPivotTable.disableFieldList = oParsed["disableFieldList"];
		if (oParsed["showCalcMbrs"] != null)
			oPivotTable.showCalcMbrs = oParsed["showCalcMbrs"];
		if (oParsed["visualTotals"] != null)
			oPivotTable.visualTotals = oParsed["visualTotals"];
		if (oParsed["showMultipleLabel"] != null)
			oPivotTable.showMultipleLabel = oParsed["showMultipleLabel"];
		if (oParsed["showDataDropDown"] != null)
			oPivotTable.showDataDropDown = oParsed["showDataDropDown"];
		if (oParsed["showDrill"] != null)
			oPivotTable.showDrill = oParsed["showDrill"];
		if (oParsed["printDrill"] != null)
			oPivotTable.printDrill = oParsed["printDrill"];
		if (oParsed["showMemberPropertyTips"] != null)
			oPivotTable.showMemberPropertyTips = oParsed["showMemberPropertyTips"];
		if (oParsed["showDataTips"] != null)
			oPivotTable.showDataTips = oParsed["showDataTips"];
		if (oParsed["enableWizard"] != null)
			oPivotTable.enableWizard = oParsed["enableWizard"];
		if (oParsed["enableDrill"] != null)
			oPivotTable.enableDrill = oParsed["enableDrill"];
		if (oParsed["enableFieldProperties"] != null)
			oPivotTable.enableFieldProperties = oParsed["enableFieldProperties"];
		if (oParsed["preserveFormatting"] != null)
			oPivotTable.preserveFormatting = oParsed["preserveFormatting"];
		if (oParsed["useAutoFormatting"] != null)
			oPivotTable.useAutoFormatting = oParsed["useAutoFormatting"];
		if (oParsed["pageWrap"] != null)
			oPivotTable.pageWrap = oParsed["pageWrap"];
		if (oParsed["pageOverThenDown"] != null)
			oPivotTable.pageOverThenDown = oParsed["pageOverThenDown"];
		if (oParsed["subtotalHiddenItems"] != null)
			oPivotTable.subtotalHiddenItems = oParsed["subtotalHiddenItems"];
		if (oParsed["rowGrandTotals"] != null)
			oPivotTable.rowGrandTotals = oParsed["rowGrandTotals"];
		if (oParsed["colGrandTotals"] != null)
			oPivotTable.colGrandTotals = oParsed["colGrandTotals"];
		if (oParsed["fieldPrintTitles"] != null)
			oPivotTable.fieldPrintTitles = oParsed["fieldPrintTitles"];
		if (oParsed["itemPrintTitles"] != null)
			oPivotTable.itemPrintTitles = oParsed["itemPrintTitles"];
		if (oParsed["mergeItem"] != null)
			oPivotTable.mergeItem = oParsed["mergeItem"];
		if (oParsed["showDropZones"] != null)
			oPivotTable.showDropZones = oParsed["showDropZones"];
		if (oParsed["createdVersion"] != null)
			oPivotTable.createdVersion = oParsed["createdVersion"];
		if (oParsed["indent"] != null)
			oPivotTable.indent = oParsed["indent"];
		if (oParsed["showEmptyRow"] != null)
			oPivotTable.showEmptyRow = oParsed["showEmptyRow"];
		if (oParsed["showEmptyCol"] != null)
			oPivotTable.showEmptyCol = oParsed["showEmptyCol"];
		if (oParsed["showHeaders"] != null)
			oPivotTable.showHeaders = oParsed["showHeaders"];
		if (oParsed["compact"] != null)
			oPivotTable.compact = oParsed["compact"];
		if (oParsed["outline"] != null)
			oPivotTable.outline = oParsed["outline"];
		if (oParsed["outlineData"] != null)
			oPivotTable.outlineData = oParsed["outlineData"];
		if (oParsed["compactData"] != null)
			oPivotTable.compactData = oParsed["compactData"];
		if (oParsed["published"] != null)
			oPivotTable.published = oParsed["published"];
		if (oParsed["gridDropZones"] != null)
			oPivotTable.gridDropZones = oParsed["gridDropZones"];
		if (oParsed["immersive"] != null)
			oPivotTable.immersive = oParsed["immersive"];
		if (oParsed["multipleFieldFilters"] != null)
			oPivotTable.multipleFieldFilters = oParsed["multipleFieldFilters"];
		if (oParsed["chartFormat"] != null)
			oPivotTable.chartFormat = oParsed["chartFormat"];
		if (oParsed["rowHeaderCaption"] != null)
			oPivotTable.rowHeaderCaption = oParsed["rowHeaderCaption"];
		if (oParsed["colHeaderCaption"] != null)
			oPivotTable.colHeaderCaption = oParsed["colHeaderCaption"];
		if (oParsed["fieldListSortAscending"] != null)
			oPivotTable.fieldListSortAscending = oParsed["fieldListSortAscending"];
		if (oParsed["mdxSubqueries"] != null)
			oPivotTable.mdxSubqueries = oParsed["mdxSubqueries"];
		if (oParsed["customListSort"] != null)
			oPivotTable.customListSort = oParsed["customListSort"];
		if (oParsed["location"] != null)
			oPivotTable.location = this.LocationFromJSON(oParsed["location"]);
		if (oParsed["pivotFields"] != null)
			oPivotTable.pivotFields = this.PivotFieldsFromJSON(oParsed["pivotFields"]);
		if (oParsed["rowFields"] != null)
			oPivotTable.rowFields = this.RowFieldsFromJSON(oParsed["rowFields"]);
		if (oParsed["rowItems"] != null)
			oPivotTable.rowItems = this.RowItemsFromJSON(oParsed["rowItems"]);
		if (oParsed["colFields"] != null)
			oPivotTable.colFields = this.ColFieldsFromJSON(oParsed["colFields"]);
		if (oParsed["colItems"] != null)
			oPivotTable.colItems = this.ColItemsFromJSON(oParsed["colItems"]);
		if (oParsed["pageFields"] != null)
			oPivotTable.pageFields = this.PageFieldsFromJSON(oParsed["pageFields"]);
		if (oParsed["dataFields"] != null)
			oPivotTable.dataFields = this.DataFieldsFromJSON(oParsed["dataFields"]);
		if (oParsed["formats"] != null)
			oPivotTable.formats = this.FormatsFromJSON(oParsed["formats"]);
		if (oParsed["conditionalFormats"] != null)
			oPivotTable.conditionalFormats = this.ConditionalFormatsFromJSON(oParsed["conditionalFormats"]);
		if (oParsed["chartFormats"] != null)
			oPivotTable.chartFormats = this.ChartFormatsFromJSON(oParsed["chartFormats"]);
		if (oParsed["pivotHierarchies"] != null)
			oPivotTable.pivotHierarchies = this.PivotHierarchiesFromJSON(oParsed["pivotHierarchies"]);
		if (oParsed["pivotTableStyleInfo"] != null)
			oPivotTable.pivotTableStyleInfo = this.PivotTableStyleInfoFromJSON(oParsed["pivotTableStyleInfo"]);
		if (oParsed["filters"] != null)
			oPivotTable.filters = this.PivotFiltersFromJSON(oParsed["filters"]);
		if (oParsed["rowHierarchiesUsage"] != null)
			oPivotTable.rowHierarchiesUsage = this.RowHierarchiesUsageFromJSON(oParsed["rowHierarchiesUsage"]);
		if (oParsed["colHierarchiesUsage"] != null)
			oPivotTable.colHierarchiesUsage = this.ColHierarchiesUsageFromJSON(oParsed["colHierarchiesUsage"]);
		if (oParsed["pivotTableDefinitionX14"] != null)
			oPivotTable.pivotTableDefinitionX14 = this.PivotTableDefinitionX14FromJSON(oParsed["pivotTableDefinitionX14"]);

		return oPivotTable;
	};
	ReaderFromJSON.prototype.LocationFromJSON = function(oParsed)
	{
		var oLocation = new CT_Location();

		if (oParsed["ref"] != null)
			oLocation.ref = AscCommonExcel.g_oRangeCache.getAscRange(oParsed["ref"]);
		if (oParsed["firstHeaderRow"] != null)
			oLocation.firstHeaderRow = oParsed["firstHeaderRow"];
		if (oParsed["firstDataRow"] != null)
			oLocation.firstDataRow = oParsed["firstDataRow"];
		if (oParsed["firstDataCol"] != null)
			oLocation.firstDataCol = oParsed["firstDataCol"];
		if (oParsed["rowPageCount"] != null)
			oLocation.rowPageCount = oParsed["rowPageCount"];
		if (oParsed["colPageCount"] != null)
			oLocation.colPageCount = oParsed["colPageCount"];

		return oLocation;
	};
	ReaderFromJSON.prototype.PivotFieldsFromJSON = function(oParsed)
	{
		var oPivotFields = new CT_PivotFields();

		for (var nField = 0; nField < oParsed["pivotField"].length; nField++)
			oPivotFields.pivotField.push(this.PivotFieldFromJSON(oParsed["pivotField"][nField]));

		return oPivotFields;
	};
	ReaderFromJSON.prototype.PivotFieldFromJSON = function(oParsed)
	{
		var oPivotField = new CT_PivotField(true);

		if (oParsed["name"] != null)
			oPivotField.name = oParsed["name"];
		if (oParsed["axis"] != null)
			oPivotField.axis = FromXml_ST_Axis(oParsed["axis"]);
		if (oParsed["dataField"] != null)
			oPivotField.dataField = oParsed["dataField"];
		if (oParsed["subtotalCaption"] != null)
			oPivotField.subtotalCaption = oParsed["subtotalCaption"];
		if (oParsed["showDropDowns"] != null)
			oPivotField.showDropDowns = oParsed["showDropDowns"];
		if (oParsed["hiddenLevel"] != null)
			oPivotField.hiddenLevel = oParsed["hiddenLevel"];
		if (oParsed["uniqueMemberProperty"] != null)
			oPivotField.uniqueMemberProperty = oParsed["uniqueMemberProperty"];
		if (oParsed["compact"] != null)
			oPivotField.compact = oParsed["compact"];
		if (oParsed["allDrilled"] != null)
			oPivotField.allDrilled = oParsed["allDrilled"];
		if (oParsed["numFmtId"] != null)
			oPivotField.numFmtId = oParsed["numFmtId"];
		if (oParsed["outline"] != null)
			oPivotField.outline = oParsed["outline"];
		if (oParsed["subtotalTop"] != null)
			oPivotField.subtotalTop = oParsed["subtotalTop"];
		if (oParsed["dragToRow"] != null)
			oPivotField.dragToRow = oParsed["dragToRow"];
		if (oParsed["dragToCol"] != null)
			oPivotField.dragToCol = oParsed["dragToCol"];
		if (oParsed["multipleItemSelectionAllowed"] != null)
			oPivotField.multipleItemSelectionAllowed = oParsed["multipleItemSelectionAllowed"];
		if (oParsed["dragToPage"] != null)
			oPivotField.dragToPage = oParsed["dragToPage"];
		if (oParsed["dragToData"] != null)
			oPivotField.dragToData = oParsed["dragToData"];
		if (oParsed["dragOff"] != null)
			oPivotField.dragOff = oParsed["dragOff"];
		if (oParsed["showAll"] != null)
			oPivotField.showAll = oParsed["showAll"];
		if (oParsed["insertBlankRow"] != null)
			oPivotField.insertBlankRow = oParsed["insertBlankRow"];
		if (oParsed["serverField"] != null)
			oPivotField.serverField = oParsed["serverField"];
		if (oParsed["insertPageBreak"] != null)
			oPivotField.insertPageBreak = oParsed["insertPageBreak"];
		if (oParsed["autoShow"] != null)
			oPivotField.autoShow = oParsed["autoShow"];
		if (oParsed["topAutoShow"] != null)
			oPivotField.topAutoShow = oParsed["topAutoShow"];
		if (oParsed["hideNewItems"] != null)
			oPivotField.hideNewItems = oParsed["hideNewItems"];
		if (oParsed["measureFilter"] != null)
			oPivotField.measureFilter = oParsed["measureFilter"];
		if (oParsed["includeNewItemsInFilter"] != null)
			oPivotField.includeNewItemsInFilter = oParsed["includeNewItemsInFilter"];
		if (oParsed["itemPageCount"] != null)
			oPivotField.itemPageCount = oParsed["itemPageCount"];
		if (oParsed["sortType"] != null)
			oPivotField.sortType = FromXml_ST_FieldSortType(oParsed["sortType"]);
		if (oParsed["dataSourceSort"] != null)
			oPivotField.dataSourceSort = oParsed["dataSourceSort"];
		if (oParsed["nonAutoSortDefault"] != null)
			oPivotField.nonAutoSortDefault = oParsed["nonAutoSortDefault"];
		if (oParsed["rankBy"] != null)
			oPivotField.rankBy = oParsed["rankBy"];
		if (oParsed["defaultSubtotal"] != null)
			oPivotField.defaultSubtotal = oParsed["defaultSubtotal"];
		if (oParsed["sumSubtotal"] != null)
			oPivotField.sumSubtotal = oParsed["sumSubtotal"];
		if (oParsed["countASubtotal"] != null)
			oPivotField.countASubtotal = oParsed["countASubtotal"];
		if (oParsed["avgSubtotal"] != null)
			oPivotField.avgSubtotal = oParsed["avgSubtotal"];
		if (oParsed["maxSubtotal"] != null)
			oPivotField.maxSubtotal = oParsed["maxSubtotal"];
		if (oParsed["minSubtotal"] != null)
			oPivotField.minSubtotal = oParsed["minSubtotal"];
		if (oParsed["productSubtotal"] != null)
			oPivotField.productSubtotal = oParsed["productSubtotal"];
		if (oParsed["countSubtotal"] != null)
			oPivotField.countSubtotal = oParsed["countSubtotal"];
		if (oParsed["stdDevSubtotal"] != null)
			oPivotField.stdDevSubtotal = oParsed["stdDevSubtotal"];
		if (oParsed["stdDevPSubtotal"] != null)
			oPivotField.stdDevPSubtotal = oParsed["stdDevPSubtotal"];
		if (oParsed["varSubtotal"] != null)
			oPivotField.varSubtotal = oParsed["varSubtotal"];
		if (oParsed["varPSubtotal"] != null)
			oPivotField.varPSubtotal = oParsed["varPSubtotal"];
		if (oParsed["showPropCell"] != null)
			oPivotField.showPropCell = oParsed["showPropCell"];
		if (oParsed["showPropTip"] != null)
			oPivotField.showPropTip = oParsed["showPropTip"];
		if (oParsed["showPropAsCaption"] != null)
			oPivotField.showPropAsCaption = oParsed["showPropAsCaption"];
		if (oParsed["defaultAttributeDrillState"] != null)
			oPivotField.defaultAttributeDrillState = oParsed["defaultAttributeDrillState"];
		if (oParsed["pivotFieldX14"] != null)
			oPivotField.pivotFieldX14 = this.PivotFieldX14FromJSON(oParsed["pivotFieldX14"]);
		if (oParsed["items"] != null)
			oPivotField.items = this.PivotFieldItemsFromJSON(oParsed["items"]);
		if (oParsed["autoSortScope"] != null)
			oPivotField.autoSortScope = this.AutoSortScopeFromJSON(oParsed["autoSortScope"]);

		return oPivotField;
	};
	ReaderFromJSON.prototype.PivotFieldItemsFromJSON = function(aParsed)
	{
		var oItems = new CT_Items();

		for (var nItem = 0; nItem < aParsed["item"].length; nItem++)
			oItems.item.push(this.PivotFieldItemFromJSON(aParsed["item"][nItem]));

		return oItems;
	};
	ReaderFromJSON.prototype.PivotFieldItemFromJSON = function(oParsed)
	{
		var oItem = new CT_Item();

		if (oParsed["n"] != null)
			oItem.n = oParsed["n"];
		if (oParsed["t"] != null)
			oItem.t = FromXml_ST_ItemType(oParsed["t"]);
		if (oParsed["h"] != null)
			oItem.h = oParsed["h"];
		if (oParsed["s"] != null)
			oItem.s = oParsed["s"];
		if (oParsed["sd"] != null)
			oItem.sd = oParsed["sd"];
		if (oParsed["f"] != null)
			oItem.f = oParsed["f"];
		if (oParsed["m"] != null)
			oItem.m = oParsed["m"];
		if (oParsed["c"] != null)
			oItem.c = oParsed["c"];
		if (oParsed["x"] != null)
			oItem.x = oParsed["x"];
		if (oParsed["d"] != null)
			oItem.d = oParsed["d"];
		if (oParsed["e"] != null)
			oItem.e = oParsed["e"];

		return oItem;
	};
	ReaderFromJSON.prototype.AutoSortScopeFromJSON = function(oParsed)
	{
		var oAutoSortScope = new CT_AutoSortScope();

		if (oParsed["pivotArea"] != null)
			oAutoSortScope.pivotArea = this.PivotAreaFromJSON(oParsed["pivotArea"]);

		return oAutoSortScope;
	};
	ReaderFromJSON.prototype.PivotAreaFromJSON = function(oParsed)
	{
		var oPivotArea = new CT_PivotArea();

		if (oParsed["field"] != null)
			oPivotArea.field = oParsed["field"];
		if (oParsed["type"] != null)
			oPivotArea.type = FromXml_ST_PivotAreaType(oParsed["type"]);
		if (oParsed["dataOnly"] != null)
			oPivotArea.dataOnly = oParsed["dataOnly"];
		if (oParsed["labelOnly"] != null)
			oPivotArea.labelOnly = oParsed["labelOnly"];
		if (oParsed["grandRow"] != null)
			oPivotArea.grandRow = oParsed["grandRow"];
		if (oParsed["grandCol"] != null)
			oPivotArea.grandCol = oParsed["grandCol"];
		if (oParsed["cacheIndex"] != null)
			oPivotArea.cacheIndex = oParsed["cacheIndex"];
		if (oParsed["outline"] != null)
			oPivotArea.outline = oParsed["outline"];
		if (oParsed["offset"] != null)
			oPivotArea.offset = oParsed["offset"];
		if (oParsed["collapsedLevelsAreSubtotals"] != null)
			oPivotArea.collapsedLevelsAreSubtotals = oParsed["collapsedLevelsAreSubtotals"];
		if (oParsed["axis"] != null)
			oPivotArea.axis = FromXml_ST_Axis(oParsed["axis"]);
		if (oParsed["fieldPosition"] != null)
			oPivotArea.fieldPosition = oParsed["fieldPosition"];
		if (oParsed["references"] != null)
			oPivotArea.references = this.PivotAreaRefsFromJSON(oParsed["references"]);
		if (oParsed["extLst"] != null)
			oPivotArea.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oPivotArea;
	};
	ReaderFromJSON.prototype.PivotAreaRefsFromJSON = function(aParsed)
	{
		var oPivotAreaRefs = new CT_PivotAreaReferences();

		for (var nRef = 0; nRef < aParsed.length; nRef++)
			oPivotAreaRefs.reference.push(this.PivotAreaRefFromJSON(aParsed[nRef]));

		return oPivotAreaRefs;
	};
	ReaderFromJSON.prototype.PivotAreaRefFromJSON = function(oParsed)
	{
		var oPivotAreaRef = new CT_PivotAreaReference();
		
		if (oParsed["field"] != null)
			oPivotAreaRef.field = oParsed["field"];
		if (oParsed["selected"] != null)
			oPivotAreaRef.selected = oParsed["selected"];
		if (oParsed["byPosition"] != null)
			oPivotAreaRef.byPosition = oParsed["byPosition"];
		if (oParsed["relative"] != null)
			oPivotAreaRef.relative = oParsed["relative"];
		if (oParsed["defaultSubtotal"] != null)
			oPivotAreaRef.defaultSubtotal = oParsed["defaultSubtotal"];
		if (oParsed["sumSubtotal"] != null)
			oPivotAreaRef.sumSubtotal = oParsed["sumSubtotal"];
		if (oParsed["countASubtotal"] != null)
			oPivotAreaRef.countASubtotal = oParsed["countASubtotal"];
		if (oParsed["avgSubtotal"] != null)
			oPivotAreaRef.avgSubtotal = oParsed["avgSubtotal"];
		if (oParsed["maxSubtotal"] != null)
			oPivotAreaRef.maxSubtotal = oParsed["maxSubtotal"];
		if (oParsed["minSubtotal"] != null)
			oPivotAreaRef.minSubtotal = oParsed["minSubtotal"];
		if (oParsed["productSubtotal"] != null)
			oPivotAreaRef.productSubtotal = oParsed["productSubtotal"];
		if (oParsed["countSubtotal"] != null)
			oPivotAreaRef.countSubtotal = oParsed["countSubtotal"];
		if (oParsed["stdDevSubtotal"] != null)
			oPivotAreaRef.stdDevSubtotal = oParsed["stdDevSubtotal"];
		if (oParsed["stdDevPSubtotal"] != null)
			oPivotAreaRef.stdDevPSubtotal = oParsed["stdDevPSubtotal"];
		if (oParsed["varSubtotal"] != null)
			oPivotAreaRef.varSubtotal = oParsed["varSubtotal"];
		if (oParsed["varPSubtotal"] != null)
			oPivotAreaRef.varPSubtotal = oParsed["varPSubtotal"];
		if (oParsed["x"].length > 0)
		{
			for (var nElm = 0; nElm < oParsed["x"].length; nElm++)
				oPivotAreaRef.x.push(this.IndexFromJSON(oParsed["x"][nElm]));
		}
		if (oParsed["extLst"] != null)
			oPivotAreaRef.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oPivotAreaRef;
	};
	ReaderFromJSON.prototype.IndexFromJSON = function(oParsed)
	{
		var oIndex = new CT_Index();
		oIndex.v = oParsed["v"];

		return oIndex;
	};
	ReaderFromJSON.prototype.ExtensionListFromJSON = function(oParsed)
	{
		var oExtLst = new CT_ExtensionList();

		for (var nExt = 0; nExt < oParsed["ext"].length; nExt++)
			oExtLst.ext.push(this.ExtensionFromJSON(oParsed["ext"].ext[nExt]));

		return oExtLst;
	};
	ReaderFromJSON.prototype.ExtensionFromJSON = function(oParsed)
	{
		var oExt = new CT_Extension();

		if (oParsed["uri"] != null)
			oExt.uri = oParsed["uri"];
		
		if (oParsed["elem"] != null)
		{
			switch (oParsed["elem"]["type"])
			{
				case "x14:pivotField":
					oExt.elem = this.PivotFieldX14FromJSON(oParsed["elem"]);
					break;
				case "x14:pivotCacheDefinition":
					oExt.elem = this.PivotCacheDefinitionX14FromJSON(oParsed["elem"]);
					break;
				case "x14:pivotTableDefinition":
					oExt.elem = this.PivotTableDefinitionX14FromJSON(oParsed["elem"]);
					break;
			}
		}

		return oExt;
	};
	ReaderFromJSON.prototype.PivotFieldX14FromJSON = function(oParsed)
	{
		var oField = new CT_PivotFieldX14();

		if (oParsed["fillDownLabels"] != null)
			oField.fillDownLabels = oParsed["fillDownLabels"];
		if (oParsed["ignore"] != null)
			oField.ignore = oParsed["ignore"];

		return oField;
	};
	ReaderFromJSON.prototype.PivotCacheDefinitionX14FromJSON = function(oParsed)
	{
		var oDefinition = new CT_PivotCacheDefinitionX14();

		if (oParsed["slicerData"] != null)
			oDefinition.slicerData = oParsed["slicerData"];
		if (oParsed["pivotCacheId"] != null)
			oDefinition.pivotCacheId = oParsed["pivotCacheId"];
		if (oParsed["supportSubqueryNonVisual"] != null)
			oDefinition.supportSubqueryNonVisual = oParsed["supportSubqueryNonVisual"];
		if (oParsed["supportSubqueryCalcMem"] != null)
			oDefinition.supportSubqueryCalcMem = oParsed["supportSubqueryCalcMem"];
		if (oParsed["supportAddCalcMems"] != null)
			oDefinition.supportAddCalcMems = oParsed["supportAddCalcMems"];

		return oDefinition;
	};
	ReaderFromJSON.prototype.PivotTableDefinitionX14FromJSON = function(oParsed)
	{
		var oDefinition = new CT_pivotTableDefinitionX14();

		if (oParsed["fillDownLabelsDefault"] != null)
			oDefinition.fillDownLabelsDefault = oParsed["fillDownLabelsDefault"];
		if (oParsed["visualTotalsForSets"] != null)
			oDefinition.visualTotalsForSets = oParsed["visualTotalsForSets"];
		if (oParsed["calculatedMembersInFilters"] != null)
			oDefinition.calculatedMembersInFilters = oParsed["calculatedMembersInFilters"];
		if (oParsed["altText"] != null)
			oDefinition.altText = oParsed["altText"];
		if (oParsed["altTextSummary"] != null)
			oDefinition.altTextSummary = oParsed["altTextSummary"];
		if (oParsed["enableEdit"] != null)
			oDefinition.enableEdit = oParsed["enableEdit"];
		if (oParsed["autoApply"] != null)
			oDefinition.autoApply = oParsed["autoApply"];
		if (oParsed["allocationMethod"] != null)
			oDefinition.allocationMethod = FromXml_ST_AllocationMethod(oParsed["allocationMethod"]);
		if (oParsed["weightExpression"] != null)
			oDefinition.weightExpression = oParsed["weightExpression"];
		if (oParsed["hideValuesRow"] != null)
			oDefinition.hideValuesRow = oParsed["hideValuesRow"];

		return oDefinition;
	};
	ReaderFromJSON.prototype.RowFieldsFromJSON = function(oParsed)
	{
		var oRowFields = new CT_RowFields();

		for (var nElm = 0; nElm < oParsed["field"].length; nElm++)
			oRowFields.field.push(this.FieldFromJSON(oParsed["field"][nElm]));

		return oRowFields;
	};
	ReaderFromJSON.prototype.FieldFromJSON = function(oParsed)
	{
		var oField = new CT_Field();
		if (oParsed["x"] != null)
			oField.x = oParsed["x"];

		return oField;
	};
	ReaderFromJSON.prototype.RowItemsFromJSON = function(oParsed)
	{
		var oRowItems = new CT_rowItems();

		for (var nElm = 0; nElm < oParsed["i"].length; nElm++)
			oRowItems.i.push(this.IFromJSON(oParsed["i"][nElm]));

		return oRowItems;
	};
	ReaderFromJSON.prototype.IFromJSON = function(oParsed)
	{
		var oI = new CT_I();

		if (oParsed["t"] != null)
			oI.t = FromXml_ST_ItemType(oParsed["t"]);
		if (oParsed["r"] != null)
			oI.r = oParsed["r"];
		if (oParsed["i"] != null)
			oI.i = oParsed["i"];
		if (oParsed["x"].length > 0)
		{
			for (var nElm = 0; nElm < oParsed["x"].length; nElm++)
				oI.x.push(this.XFromJSON(oParsed["x"][nElm]));
		}

		return oI;
	};
	ReaderFromJSON.prototype.XFromJSON = function(oParsed)
	{
		var oX = new CT_X();

		if (oParsed["v"] != null)
			oX.v = oParsed["v"];

		return oX;
	};
	ReaderFromJSON.prototype.ColFieldsFromJSON = function(oParsed)
	{
		var oColFields = new CT_ColFields();

		for (var nElm = 0; nElm < oParsed["field"].length; nElm++)
			oColFields.field.push(this.FieldFromJSON(oParsed["field"][nElm]));

		return oColFields;
	};
	ReaderFromJSON.prototype.ColItemsFromJSON = function(oParsed)
	{
		var oColItems = new CT_colItems();

		for (var nElm = 0; nElm < oParsed["i"].length; nElm++)
			oColItems.i.push(this.IFromJSON(oParsed["i"][nElm]));

		return oColItems;
	};
	ReaderFromJSON.prototype.PageFieldsFromJSON = function(oParsed)
	{
		var oPageFields = new CT_PageFields();

		for (var nElm = 0; nElm < oParsed["pageField"].length; nElm++)
			oPageFields.pageField.push(this.PageFieldFromJSON(oParsed["pageField"][nElm]));

		return oPageFields;
	};
	ReaderFromJSON.prototype.PageFieldFromJSON = function(oParsed)
	{
		var oPageField = new CT_PageField();
		
		if (oParsed["fld"] != null)
			oPageField.fld = oParsed["fld"];
		if (oParsed["item"] != null)
			oPageField.item = oParsed["item"];
		if (oParsed["hier"] != null)
			oPageField.hier = oParsed["hier"];
		if (oParsed["name"] != null)
			oPageField.name = oParsed["name"];
		if (oParsed["cap"] != null)
			oPageField.cap = oParsed["cap"];
		if (oParsed["extLst"] != null)
			oPageField.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oPageField;
	};
	ReaderFromJSON.prototype.DataFieldsFromJSON = function(oParsed)
	{
		var oDataFields = new CT_DataFields();
		
		for (var nElm = 0; nElm < oParsed["dataField"].length; nElm++)
			oDataFields.dataField.push(this.DataFieldFromJSON(oParsed["dataField"][nElm]));

		return oDataFields;
	};
	ReaderFromJSON.prototype.DataFieldFromJSON = function(oParsed)
	{
		var oDataField = new CT_DataField(true);

		if (oParsed["name"] != null)
			oDataField.name = oParsed["name"];
		if (oParsed["fld"] != null)
			oDataField.fld = oParsed["fld"];
		if (oParsed["subtotal"] != null)
			oDataField.subtotal = FromXml_ST_DataConsolidateFunction(oParsed["subtotal"]);
		if (oParsed["showDataAs"] != null)
			oDataField.showDataAs = FromXml_ST_ShowDataAs(oParsed["showDataAs"]);
		if (oParsed["baseField"] != null)
			oDataField.baseField = oParsed["baseField"];
		if (oParsed["baseItem"] != null)
			oDataField.baseItem = oParsed["baseItem"];
		if (oParsed["numFmtId"] != null)
			oDataField.numFmtId = oParsed["numFmtId"];
		if (oParsed["extLst"] != null)
			oDataField.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oDataField;
	};
	ReaderFromJSON.prototype.FormatsFromJSON = function(oParsed)
	{
		var oFormats = new CT_Formats();

		for (var nElm = 0; nElm < oParsed["format"].length; nElm++)
			oFormats.format.push(this.FormatFromJSON(oParsed["format"][nElm]));

		return oFormats;
	};
	ReaderFromJSON.prototype.FormatFromJSON = function(oParsed)
	{
		var oFormat = new CT_Format();

		if (oParsed["action"] != null)
			oFormat.action = FromXml_ST_FormatAction(oParsed["action"]);
		if (oParsed["pivotArea"] != null)
			oFormat.pivotArea = this.PivotAreaFromJSON(oParsed["pivotArea"]);
		if (oParsed["extLst"] != null)
			oFormat.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oFormat;
	};
	ReaderFromJSON.prototype.ConditionalFormatsFromJSON = function(oParsed)
	{
		var oCondFormats = new CT_ConditionalFormats();

		for (var nElm = 0; nElm < oParsed["conditionalFormat"].length; nElm++)
			oCondFormats.conditionalFormat.push(this.ConditionalFormatFromJSON(oParsed["conditionalFormat"][nElm]));
	
		return oCondFormats;
	};
	ReaderFromJSON.prototype.ConditionalFormatFromJSON = function(oParsed)
	{
		var oCondFormat = new CT_ConditionalFormat();

		if (oParsed["scope"] != null)
			oCondFormat.scope = FromXml_ST_Scope(oParsed["scope"]);
		if (oParsed["type"] != null)
			oCondFormat.type = FromXml_ST_Type(oParsed["type"]);
		if (oParsed["priority"] != null)
			oCondFormat.priority = oParsed["priority"];
		if (oParsed["pivotAreas"] != null)
			oCondFormat.pivotAreas = this.PivotAreasFromJSON(oParsed["pivotAreas"]);
		if (oParsed["extLst"] != null)
			oCondFormat.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oCondFormat;
	};
	ReaderFromJSON.prototype.PivotAreasFromJSON = function(oParsed)
	{
		var oPivotAreas = new CT_PivotAreas();

		for (var nElm = 0; nElm < oParsed["pivotArea"].length; nElm++)
			oPivotAreas.pivotArea.push(this.PivotAreaFromJSON(oParsed["pivotArea"][nElm]));
	
		return oPivotAreas;
	};
	ReaderFromJSON.prototype.ChartFormatsFromJSON = function(oParsed)
	{
		var oChartFormats = new CT_ChartFormats();

		for (var nElm = 0; nElm < oParsed["chartFormat"].length; nElm++)
			oChartFormats.chartFormat.push(this.ChartFormatFromJSON(oParsed["chartFormat"][nElm]));
	
		return oChartFormats;
	};
	ReaderFromJSON.prototype.ChartFormatFromJSON = function(oParsed)
	{
		var oChartFormat = new CT_ChartFormat();

		if (oParsed["chart"] != null)
			oChartFormat.chart = oParsed["chart"];
		if (oParsed["format"] != null)
			oChartFormat.format = oParsed["format"];
		if (oParsed["series"] != null)
			oChartFormat.series = oParsed["series"];
		if (oParsed["pivotArea"] != null)
			oChartFormat.pivotArea = this.PivotAreaFromJSON(oParsed["pivotArea"]);

		return oChartFormat;
	};
	ReaderFromJSON.prototype.PivotHierarchiesFromJSON = function(oParsed)
	{
		var oPivotHiers = new CT_PivotHierarchies();

		for (var nElm = 0; nElm < oParsed["pivotHierarchy"].length; nElm++)
			oPivotHiers.pivotHierarchy.push(this.PivotHierarchyFromJSON(oParsed["pivotHierarchy"][nElm]));
		
		return oPivotHiers;
	};
	ReaderFromJSON.prototype.PivotHierarchyFromJSON = function(oParsed)
	{
		var oPivotHier = new CT_PivotHierarchy();

		if (oParsed["outline"] != null)
			oPivotHier.outline = oParsed["outline"];
		if (oParsed["multipleItemSelectionAllowed"] != null)
			oPivotHier.multipleItemSelectionAllowed = oParsed["multipleItemSelectionAllowed"];
		if (oParsed["subtotalTop"] != null)
			oPivotHier.subtotalTop = oParsed["subtotalTop"];
		if (oParsed["showInFieldList"] != null)
			oPivotHier.showInFieldList = oParsed["showInFieldList"];
		if (oParsed["dragToRow"] != null)
			oPivotHier.dragToRow = oParsed["dragToRow"];
		if (oParsed["dragToCol"] != null)
			oPivotHier.dragToCol = oParsed["dragToCol"];
		if (oParsed["dragToPage"] != null)
			oPivotHier.dragToPage = oParsed["dragToPage"];
		if (oParsed["dragToData"] != null)
			oPivotHier.dragToData = oParsed["dragToData"];
		if (oParsed["dragOff"] != null)
			oPivotHier.dragOff = oParsed["dragOff"];
		if (oParsed["includeNewItemsInFilter"] != null)
			oPivotHier.includeNewItemsInFilter = oParsed["includeNewItemsInFilter"];
		if (oParsed["caption"] != null)
			oPivotHier.caption = oParsed["caption"];
		if (oParsed["mps"] != null)
			oPivotHier.mps = this.MemberPropsFromJSON(oParsed["mps"]);
		if (oParsed["members"] != null)
			oPivotHier.members = this.MembersFromJSON(oParsed["members"]);
		if (oParsed["extLst"] != null)
			oPivotHier.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oPivotHier;
	};
	ReaderFromJSON.prototype.MemberPropsFromJSON = function(oParsed)
	{
		var oMemberProps = new CT_MemberProperties();

		for (var nElm = 0; nElm < oParsed["mp"].length; nElm++)
			oMemberProps.mp.push(this.MemberPropFromJSON(oParsed["mp"][nElm]));
	
		return oMemberProps;
	};
	ReaderFromJSON.prototype.MemberPropFromJSON = function(oParsed)
	{
		var oMemberProp = new CT_MemberProperty();

		if (oParsed["name"] != null)
			oMemberProp.name = oParsed["name"];
		if (oParsed["showCell"] != null)
			oMemberProp.showCell = oParsed["showCell"];
		if (oParsed["showTip"] != null)
			oMemberProp.showTip = oParsed["showTip"];
		if (oParsed["showAsCaption"] != null)
			oMemberProp.showAsCaption = oParsed["showAsCaption"];
		if (oParsed["nameLen"] != null)
			oMemberProp.nameLen = oParsed["nameLen"];
		if (oParsed["pPos"] != null)
			oMemberProp.pPos = oParsed["pPos"];
		if (oParsed["pLen"] != null)
			oMemberProp.pLen = oParsed["pLen"];
		if (oParsed["level"] != null)
			oMemberProp.level = oParsed["level"];
		if (oParsed["field"] != null)
			oMemberProp.field = oParsed["field"];

		return oMemberProp;
	};
	ReaderFromJSON.prototype.MembersFromJSON = function(oParsed)
	{
		var oMembers = new CT_Members();

		for (var nElm = 0; nElm < oParsed["member"].length; nElm++)
			oMembers.member.push(this.MemberFromJSON(oParsed["member"][nElm]));

		if (oParsed["level"] != null)
			oMembers.level = oParsed["level"];

		return oMembers;
	};
	ReaderFromJSON.prototype.MemberFromJSON = function(oParsed)
	{
		var oMember = new CT_Member();

		if (oParsed["name"] != null)
			oMember.name = oParsed["name"];

		return oMember;
	};
	ReaderFromJSON.prototype.PivotTableStyleInfoFromJSON = function(oParsed)
	{
		var oStyleInfo = new Asc.CT_PivotTableStyle();

		if (oParsed["name"] != null)
			oStyleInfo.name = oParsed["name"];
		if (oParsed["showRowHeaders"] != null)
			oStyleInfo.showRowHeaders = oParsed["showRowHeaders"];
		if (oParsed["showColHeaders"] != null)
			oStyleInfo.showColHeaders = oParsed["showColHeaders"];
		if (oParsed["showRowStripes"] != null)
			oStyleInfo.showRowStripes = oParsed["showRowStripes"];
		if (oParsed["showColStripes"] != null)
			oStyleInfo.showColStripes = oParsed["showColStripes"];
		if (oParsed["showLastColumn"] != null)
			oStyleInfo.showLastColumn = oParsed["showLastColumn"];

		return oStyleInfo;
	};
	ReaderFromJSON.prototype.PivotFiltersFromJSON = function(oParsed)
	{
		var oFilters = new CT_PivotFilters();

		for (var nElm = 0; nElm < oParsed["filter"].length; nElm++)
			oFilters.filter.push(this.PivotFilterFromJSON(oParsed["filter"][nElm]));

		return oFilters;
	};
	ReaderFromJSON.prototype.PivotFilterFromJSON = function(oParsed)
	{
		var oFilter = new CT_PivotFilter();

		if (oParsed["fld"] != null)
			oFilter.fld = oParsed["fld"];
		if (oParsed["mpFld"] != null)
			oFilter.mpFld = oParsed["mpFld"];
		if (oParsed["type"] != null)
			oFilter.type = FromXml_ST_PivotFilterType(oParsed["type"]);
		if (oParsed["evalOrder"] != null)
			oFilter.evalOrder = oParsed["evalOrder"];
		//if (oParsed["id"] != null)
		//	oFilter.id = oParsed["id"];
		if (oParsed["iMeasureHier"] != null)
			oFilter.iMeasureHier = oParsed["iMeasureHier"];
		if (oParsed["iMeasureFld"] != null)
			oFilter.iMeasureFld = oParsed["iMeasureFld"];
		if (oParsed["name"] != null)
			oFilter.name = oParsed["name"];
		if (oParsed["description"] != null)
			oFilter.description = oParsed["description"];
		if (oParsed["stringValue1"] != null)
			oFilter.stringValue1 = oParsed["stringValue1"];
		if (oParsed["stringValue2"] != null)
			oFilter.stringValue2 = oParsed["stringValue2"];
		if (oParsed["autoFilter"] != null)
			oFilter.autoFilter = this.AutoFilterFromJSON(oParsed["autoFilter"]);
		if (oParsed["extLst"] != null)
			oFilter.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oFilter;
	};
	ReaderFromJSON.prototype.RowHierarchiesUsageFromJSON = function(oParsed)
	{
		var oUsages = new CT_RowHierarchiesUsage();

		for (var nElm = 0; nElm < oParsed["rowHierarchyUsage"].length; nElm++)
			oUsages.rowHierarchyUsage.push(this.HierarchyUsageFromJSON(oParsed["rowHierarchyUsage"][nElm]));

		return oUsages;
	};
	ReaderFromJSON.prototype.HierarchyUsageFromJSON = function(oParsed)
	{
		var oUsage = new CT_HierarchyUsage();

		if (oParsed["hierarchyUsage"] != null)
			oUsage.hierarchyUsage = oParsed["hierarchyUsage"];

		return oUsage;
	};
	ReaderFromJSON.prototype.ColHierarchiesUsageFromJSON = function(oParsed)
	{
		var oUsages = new CT_ColHierarchiesUsage();

		for (var nElm = 0; nElm < oParsed["colHierarchyUsage"].length; nElm++)
			oUsages.colHierarchyUsage.push(this.HierarchyUsageFromJSON(oParsed["colHierarchyUsage"][nElm]));

		return oUsages;
	};
	ReaderFromJSON.prototype.SlicersFromJSON = function(aParsed, oWorksheet)
	{
		var oSlicers = new Asc.CT_slicers();
		oSlicers.slicer = oWorksheet.aSlicers;

		for (var nElm = 0; nElm < aParsed.length; nElm++)
			oSlicers.slicer.push(this.SlicerFromJSON(aParsed[nElm], oWorksheet));

		// тут отсюда здесь
		//return oSlicers;
	};
	ReaderFromJSON.prototype.SlicerFromJSON = function(oParsed, oWorksheet)
	{
		var oSlicer = new Asc.CT_slicer(oWorksheet);

		if (oParsed["name"] != null)
			oSlicer.name = oParsed["name"];
		//if (oParsed["uid"] != null)
		//	oSlicer.uid = oParsed["uid"];
		if (oParsed["cacheDefinition"] != null)
			oSlicer.cacheDefinition = this.slicerCaches[oParsed["cacheDefinition"]] || this.slicerCachesExt[oParsed["cacheDefinition"]];
		if (oParsed["caption"] != null)
			oSlicer.caption = oParsed["caption"];
		if (oParsed["startItem"] != null)
			oSlicer.startItem = oParsed["startItem"];
		if (oParsed["columnCount"] != null)
			oSlicer.columnCount = oParsed["columnCount"];
		if (oParsed["showCaption"] != null)
			oSlicer.showCaption = oParsed["showCaption"];
		if (oParsed["level"] != null)
			oSlicer.level = oParsed["level"];
		if (oParsed["style"] != null)
			oSlicer.style = oParsed["style"];
		if (oParsed["lockedPosition"] != null)
			oSlicer.lockedPosition = oParsed["lockedPosition"];
		if (oParsed["rowHeight"] != null)
			oSlicer.rowHeight = oParsed["rowHeight"];

		return oSlicer;
	};
	ReaderFromJSON.prototype.SlicerCacheDefinitionFromJSON = function(oParsed)
	{
		var oCache = new Asc.CT_slicerCacheDefinition(this.Workbook);

		if (oParsed["name"] != null)
			oCache.name = oParsed["name"];
		//if (oParsed["uid"] != null)
		//	oCache.uid = oParsed["uid"];
		if (oParsed["sourceName"] != null)
			oCache.sourceName = oParsed["sourceName"];
		if (oParsed["pivotTables"] != null)
			oCache.pivotTables = this.SlicerCachePivotTablesFromJSON(oParsed["pivotTables"]);
		if (oParsed["data"] != null)
			oCache.data = this.SlicerCacheDataFromJSON(oParsed["data"]);
		if (oParsed["tableSlicerCache"] != null)
			oCache.tableSlicerCache = this.TableSlicerCacheFromJSON(oParsed["tableSlicerCache"]);
		if (oParsed["slicerCacheHideItemsWithNoData"] != null)
			oCache.slicerCacheHideItemsWithNoData = this.SlicerCacheHideNoDataFromJSON(oParsed["slicerCacheHideItemsWithNoData"]);

		return oCache;
	};
	ReaderFromJSON.prototype.SlicerCachePivotTablesFromJSON = function(aParsed)
	{
		var aResult = [];
		for (var nTable = 0; nTable < aParsed.length; nTable++)
			aResult.push(this.SlicerCachePivotTableFromJSON(aParsed[nTable]));

		return aResult;
	};
	ReaderFromJSON.prototype.SlicerCachePivotTableFromJSON = function(oParsed)
	{
		var oTable = new Asc.CT_slicerCachePivotTable();
		
		if (oParsed["name"] != null)
			oTable.name = oParsed["name"];
		//if (oParsed["tabIdOpen"] != null)
		//	oTable.tabIdOpen = oParsed["tabIdOpen"];
		if (oParsed["sheetId"] != null)
			oTable.sheetId = oParsed["sheetId"];
		return oTable;
	};
	ReaderFromJSON.prototype.SlicerCacheDataFromJSON = function(oParsed)
	{
		var oData = new Asc.CT_slicerCacheData();

		if (oParsed["olap"] != null)
			oData.olap = this.OlapSlicerCacheFromJSON(oParsed["olap"]);

		if (oParsed["tabular"] != null)
			oData.tabular = this.TabularSlicerCacheFromJSON(oParsed["tabular"]);
	
		return oData;
	};
	ReaderFromJSON.prototype.OlapSlicerCacheFromJSON = function(oParsed)
	{
		var oOlap = new Asc.CT_olapSlicerCache();

		if (oParsed["levels"].length > 0)
		{
			for (var nElm = 0; nElm < oParsed["levels"].length; nElm++)
				oOlap.levels.push(this.OlapSlicerCacheLevelDataFromJSON(oParsed["levels"][nElm]));
		}
		if (oParsed["selections"].length > 0)
		{
			for (var nElm = 0; nElm < oParsed["selections"].length; nElm++)
				oOlap.selections.push(this.OlapSlicerCacheSelectionFromJSON(oParsed["selections"][nElm]));
		}
		// to do (возможно стоит писать в какой-то общий массив, после оттуда зачитывать)
		if (oParsed["pivotCacheDefinition"] != null)
			oOlap.pivotCacheDefinition = this.PivotCacheDefinitionFromJSON(oParsed["pivotCacheDefinition"]);

		return oOlap;
	};
	ReaderFromJSON.prototype.OlapSlicerCacheLevelDataFromJSON = function(oParsed)
	{
		var oLvlData = new Asc.CT_olapSlicerCacheLevelData();

		if (oParsed["uniqueName"] != null)
			oLvlData.uniqueName = oParsed["uniqueName"];
		if (oParsed["sourceCaption"] != null)
			oLvlData.sourceCaption = oParsed["sourceCaption"];
		if (oParsed["count"] != null)
			oLvlData.count = oParsed["count"];
		if (oParsed["sortOrder"] != null)
			oLvlData.sortOrder = FromXML_ST_olapSlicerCacheSortOrder(oParsed["sortOrder"]);
		if (oParsed["crossFilter"] != null)
			oLvlData.crossFilter = FromXML_ST_slicerCacheCrossFilter(oParsed["crossFilter"]);
		if (oParsed["ranges"] != null)
			oLvlData.ranges = this.OlapSlicerCacheRangesFromJSON(oParsed["ranges"]);

		return oLvlData;		
	};
	ReaderFromJSON.prototype.OlapSlicerCacheRangesFromJSON = function(aParsed)
	{
		var aResult = [];
		for (var nRange = 0; nRange < aParsed.length; nRange++)
			aResult.push(this.OlapSlicerCacheRangeFromJSON(aParsed[nRange]));

		return aResult;
	};
	ReaderFromJSON.prototype.OlapSlicerCacheRangeFromJSON = function(oParsed)
	{
		var oRange = new Asc.CT_olapSlicerCacheRange();

		if (oParsed["startItem"] != null)
			oRange.startItem = oParsed["startItem"];
		if (oParsed["i"].length > 0)
		{
			for (var nElm = 0; nElm < oParsed["i"].length; nElm++)
				oRange.i.push(this.OlapSlicerCacheItemFromJSON(oParsed["i"][nElm]));
		}
		
		return oRange;
	};
	ReaderFromJSON.prototype.OlapSlicerCacheItemFromJSON = function(oParsed)
	{
		var oItem = new Asc.CT_olapSlicerCacheItem();

		if (oParsed["n"] != null)
			oItem.n = oParsed["n"];
		if (oParsed["c"] != null)
			oItem.c = oParsed["c"];
		if (oParsed["nd"] != null)
			oItem.nd = oParsed["nd"];
		if (oParsed["p"] != null)
			oItem.p = this.OlapSlicerCacheItemParentsFromJSON(oParsed["p"]);

		return oItem;
	};
	ReaderFromJSON.prototype.OlapSlicerCacheItemParentsFromJSON = function(aParsed)
	{
		var aResult = [];
		for (var nIndex = 0; nIndex < aParsed.length; nIndex++)
			aResult.push(this.OlapSlicerCacheItemParentFromJSON(aParsed[nIndex]));

		return aResult;
	};
	ReaderFromJSON.prototype.OlapSlicerCacheItemParentFromJSON = function(oParsed)
	{
		var oParent = new Asc.CT_olapSlicerCacheItemParent();

		if (oParsed["n"] != null)
			oParent.n = oParsed["n"];

		return oParent;
	};
	ReaderFromJSON.prototype.OlapSlicerCacheSelectionFromJSON = function(oParsed)
	{
		var oSelection = new Asc.CT_olapSlicerCacheSelection();

		if (oParsed["n"] != null)
			oSelection.n = oParsed["n"];
		if (oParsed["p"] != null)
			oSelection.p = this.OlapSlicerCacheItemParentsFromJSON(oParsed["p"]);

		return oSelection;
	};
	ReaderFromJSON.prototype.PivotCacheDefinitionFromJSON = function(oParsed)
	{
		var oPivotCahceDef = new Asc.CT_PivotCacheDefinition();

		// attributes
		// if (oParsed["id"] != null)
		// 	oPivotCahceDef.id = oParsed["id"];
		if (oParsed["invalid"] != null)
			oPivotCahceDef.invalid = oParsed["invalid"];
		if (oParsed["saveData"] != null)
			oPivotCahceDef.saveData = oParsed["saveData"];
		if (oParsed["refreshOnLoad"] != null)
			oPivotCahceDef.refreshOnLoad = oParsed["refreshOnLoad"];
		if (oParsed["optimizeMemory"] != null)
			oPivotCahceDef.optimizeMemory = oParsed["optimizeMemory"];
		if (oParsed["enableRefresh"] != null)
			oPivotCahceDef.enableRefresh = oParsed["enableRefresh"];
		if (oParsed["refreshedBy"] != null)
			oPivotCahceDef.refreshedBy = oParsed["refreshedBy"];
		if (oParsed["refreshedDate"] != null)
			oPivotCahceDef.refreshedDate = oParsed["refreshedDate"];
		if (oParsed["backgroundQuery"] != null)
			oPivotCahceDef.backgroundQuery = oParsed["backgroundQuery"];
		if (oParsed["missingItemsLimit"] != null)
			oPivotCahceDef.missingItemsLimit = oParsed["missingItemsLimit"];
		if (oParsed["createdVersion"] != null)
			oPivotCahceDef.createdVersion = oParsed["createdVersion"];
		if (oParsed["refreshedVersion"] != null)
			oPivotCahceDef.refreshedVersion = oParsed["refreshedVersion"];
		if (oParsed["minRefreshableVersion"] != null)
			oPivotCahceDef.minRefreshableVersion = oParsed["minRefreshableVersion"];
		if (oParsed["upgradeOnRefresh"] != null)
			oPivotCahceDef.upgradeOnRefresh = oParsed["upgradeOnRefresh"];
		if (oParsed["tupleCache"] != null)
			oPivotCahceDef.tupleCache = oParsed["tupleCache"];
		if (oParsed["supportSubquery"] != null)
			oPivotCahceDef.supportSubquery = oParsed["supportSubquery"];
		if (oParsed["supportAdvancedDrill"] != null)
			oPivotCahceDef.supportAdvancedDrill = oParsed["supportAdvancedDrill"];
		// members
		if (oParsed["cacheRecords"] != null)
			oPivotCahceDef.cacheRecords = this.PivotCacheRecordsFromJSON(oParsed["cacheRecords"]);
		if (oParsed["cacheSource"] != null)
			oPivotCahceDef.cacheSource = this.CacheSourceFromJSON(oParsed["cacheSource"]);
		if (oParsed["cacheFields"] != null)
			oPivotCahceDef.cacheFields = this.CacheFieldsFromJSON(oParsed["cacheFields"]);
		if (oParsed["cacheHierarchies"] != null)
			oPivotCahceDef.cacheHierarchies = this.CacheHierarchiesFromJSON(oParsed["cacheHierarchies"]);
		if (oParsed["kpis"] != null)
			oPivotCahceDef.kpis = this.PCDKPIsFromJSON(oParsed["kpis"]);
		if (oParsed["tupleCache"] != null)
			oPivotCahceDef.tupleCache = this.TupleCacheFromJSON(oParsed["tupleCache"]);
		if (oParsed["calculatedItems"] != null)
			oPivotCahceDef.calculatedItems = this.CalculatedItemsFromJSON(oParsed["calculatedItems"]);
		if (oParsed["calculatedMembers"] != null)
			oPivotCahceDef.calculatedMembers = this.CalculatedMembersFromJSON(oParsed["calculatedMembers"]);
		if (oParsed["dimensions"] != null)
			oPivotCahceDef.dimensions = this.DimensionsFromJSON(oParsed["dimensions"]);
		if (oParsed["measureGroups"] != null)
			oPivotCahceDef.measureGroups = this.MeasureGroupsFromJSON(oParsed["measureGroups"]);
		if (oParsed["maps"] != null)
			oPivotCahceDef.maps = this.MeasureDimensionMapsFromJSON(oParsed["maps"]);
			// ext
		if (oParsed["pivotCacheDefinitionX14"] != null)
			oPivotCahceDef.pivotCacheDefinitionX14 = this.PivotCacheDefinitionX14FromJSON(oParsed["pivotCacheDefinitionX14"]); // to do

		return oPivotCahceDef;
	};
	ReaderFromJSON.prototype.PivotCacheRecordsFromJSON = function(oParsed)
	{
		let oRecords = new Asc.CT_PivotCacheRecords();
		
		if (oParsed["records"].length)
		{
			for (let i = 0; i < oParsed["records"][0].length; i++)
				oRecords._cols.push(new Asc.PivotRecords());
		}

		for (let i = 0; i < oParsed["records"].length; i++) {
			for (let j = 0; j < oParsed["records"][i].length; j++) {
				this.PivotRecordsFromJSON([oParsed["records"][i][j]], oRecords._cols[j]);
			}
		}

		oRecords.startCount = oParsed["count"];

		if (oParsed["extLst"])
			oRecords.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oRecords;
	};
	ReaderFromJSON.prototype.PivotCachesFromJSON = function(oParsed)
	{
		var oPivotCaches = {};

		for (var key in oParsed)
			oPivotCaches[key] = this.PivotCacheDefinitionFromJSON(oParsed[key]);

		return oPivotCaches;
	};
	ReaderFromJSON.prototype.SlicerCachesFromJSON = function(oParsed)
	{
		var oSlicerCaches = {};

		for (var key in oParsed)
			oSlicerCaches[key] = this.SlicerCacheDefinitionFromJSON(oParsed[key]);

		return oSlicerCaches;
	};
	ReaderFromJSON.prototype.CacheSourceFromJSON = function(oParsed)
	{
		var oSource = new CT_CacheSource();

		if (oParsed["type"] != null)
			oSource.type = FromXml_ST_SourceType(oParsed["type"]);
		if (oParsed["connectionId"] != null)
			oSource.connectionId = oParsed["connectionId"];
		if (oParsed["consolidation"] != null)
			oSource.consolidation = this.ConsolidationFromJSON(oParsed["consolidation"]);
		if (oParsed["extLst"] != null)
			oSource.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);
		if (oParsed["worksheetSource"] != null)
			oSource.worksheetSource = this.WorksheetSourceFromJSON(oParsed["worksheetSource"]);

		return oSource;
	};
	ReaderFromJSON.prototype.ConsolidationFromJSON = function(oParsed)
	{
		var oConsolidation = new CT_Consolidation();

		if (oParsed["autoPage"] != null)
			oConsolidation.autoPage = oParsed["autoPage"];
		if (oParsed["pages"] != null)
			oConsolidation.pages = this.PagesFromJSON(oParsed["pages"]);
		if (oParsed["rangeSets"] != null)
			oConsolidation.rangeSets = this.RangeSetsFromJSON(oParsed["rangeSets"]);

		return oConsolidation;
	};
	ReaderFromJSON.prototype.PagesFromJSON = function(oParsed)
	{
		var oPages = new CT_Pages();

		for (var nElm = 0; nElm < oParsed["page"].length; nElm++)
			oPages.page.push(this.PageFromJSON(oParsed["page"][nElm]));

		return oPages;
	};
	ReaderFromJSON.prototype.PageFromJSON = function(oParsed)
	{
		var oPage = new CT_PCDSCPage();
		for (var nElm = 0; nElm < oParsed["pageItem"].length; nElm++)
			oPage.pageItem.push(this.PageItemFromJSON(oParsed["pageItem"][nElm]));

		return oPage;
	};
	ReaderFromJSON.prototype.PageItemFromJSON = function(oParsed)
	{
		var oPageItem = new CT_PageItem();
		if (oParsed["name"] != null)
			oPageItem.name = oParsed["name"];

		return oPageItem;
	};
	ReaderFromJSON.prototype.RangeSetsFromJSON = function(oParsed)
	{
		var oSets = new CT_RangeSets();

		for (var nElm = 0; nElm < oParsed["rangeSet"].length; nElm++)
			oSets.rangeSet.push(this.RangeSetFromJSON(oParsed["rangeSet"][nElm]));

		return oSets;
	};
	ReaderFromJSON.prototype.RangeSetFromJSON = function(oParsed)
	{
		var oSet = new CT_RangeSet();

		if (oParsed["i1"] != null)
			oSet.i1 = oParsed["i1"];
		if (oParsed["i2"] != null)
			oSet.i2 = oParsed["i2"];
		if (oParsed["i3"] != null)
			oSet.i3 = oParsed["i3"];
		if (oParsed["i4"] != null)
			oSet.i4 = oParsed["i4"];
		if (oParsed["ref"] != null)
			oSet.ref = oParsed["ref"];
		if (oParsed["name"] != null)
			oSet.name = oParsed["name"];
		if (oParsed["sheet"] != null)
			oSet.sheet = oParsed["sheet"];
	
		return oSet;
	};
	ReaderFromJSON.prototype.WorksheetSourceFromJSON = function(oParsed)
	{
		var oWSource = new Asc.CT_WorksheetSource();

		if (oParsed["ref"] != null)
			oWSource.ref = oParsed["ref"];
		if (oParsed["name"] != null)
			oWSource.name = oParsed["name"];
		if (oParsed["sheet"] != null)
			oWSource.sheet = oParsed["sheet"];

		return oWSource;
	};
	ReaderFromJSON.prototype.CacheFieldsFromJSON = function(oParsed)
	{
		var oFields = new CT_CacheFields();

		for (var nElm = 0; nElm < oParsed["cacheField"].length; nElm++)
			oFields.cacheField.push(this.CacheFieldFromJSON(oParsed["cacheField"][nElm]));

		return oFields;
	};
	ReaderFromJSON.prototype.CacheFieldFromJSON = function(oParsed)
	{
		var oField = new Asc.CT_CacheField();

		if (oParsed["name"] != null)
			oField.name = oParsed["name"];
		if (oParsed["caption"] != null)
			oField.caption = oParsed["caption"];
		if (oParsed["propertyName"] != null)
			oField.propertyName = oParsed["propertyName"];
		if (oParsed["serverField"] != null)
			oField.serverField = oParsed["serverField"];
		if (oParsed["uniqueList"] != null)
			oField.uniqueList = oParsed["uniqueList"];
		if (oParsed["numFmtId"] != null)
			oField.numFmtId = oParsed["numFmtId"];
		if (oParsed["formula"] != null)
			oField.formula = oParsed["formula"];
		if (oParsed["sqlType"] != null)
			oField.sqlType = oParsed["sqlType"];
		if (oParsed["hierarchy"] != null)
			oField.hierarchy = oParsed["hierarchy"];
		if (oParsed["level"] != null)
			oField.level = oParsed["level"];
		if (oParsed["databaseField"] != null)
			oField.databaseField = oParsed["databaseField"];
		if (oParsed["mappingCount"] != null)
			oField.mappingCount = oParsed["mappingCount"];
		if (oParsed["memberPropertyField"] != null)
			oField.memberPropertyField = oParsed["memberPropertyField"];
		if (oParsed["sharedItems"] != null)
			oField.sharedItems = this.SharedItemsFromJSON(oParsed["sharedItems"]);
		if (oParsed["fieldGroup"] != null)
			oField.fieldGroup = this.FieldGroupFromJSON(oParsed["fieldGroup"]);
		if (oParsed["mpMap"] != null)
		{
			for (var nElm = 0; nElm < oParsed["mpMap"].length; nElm++)
				oField.mpMap.push(this.XFromJSON(oParsed["mpMap"][nElm]));
		}
		if (oParsed["extLst"] != null)
			oField.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oField;
	};
	ReaderFromJSON.prototype.SharedItemsFromJSON = function(oParsed)
	{
		var oItems = new CT_SharedItems();

		if (oParsed["containsSemiMixedTypes"] != null)
			oItems.containsSemiMixedTypes = oParsed["containsSemiMixedTypes"];
		if (oParsed["containsNonDate"] != null)
			oItems.containsNonDate = oParsed["containsNonDate"];
		if (oParsed["containsDate"] != null)
			oItems.containsDate = oParsed["containsDate"];
		if (oParsed["containsString"] != null)
			oItems.containsString = oParsed["containsString"];
		if (oParsed["containsBlank"] != null)
			oItems.containsBlank = oParsed["containsBlank"];
		if (oParsed["containsMixedTypes"] != null)
			oItems.containsMixedTypes = oParsed["containsMixedTypes"];
		if (oParsed["containsNumber"] != null)
			oItems.containsNumber = oParsed["containsNumber"];
		if (oParsed["containsInteger"] != null)
			oItems.containsInteger = oParsed["containsInteger"];
		if (oParsed["minValue"] != null)
			oItems.minValue = oParsed["minValue"];
		if (oParsed["maxValue"] != null)
			oItems.maxValue = oParsed["maxValue"];
		if (oParsed["minDate"] != null)
			oItems.minDate = Asc.cDate.prototype.fromISO8601(oParsed["minDate"]);
		if (oParsed["maxDate"] != null)
			oItems.maxDate = Asc.cDate.prototype.fromISO8601(oParsed["maxDate"]);
		if (oParsed["items"].length > 0)
			this.PivotRecordsFromJSON(oParsed["items"], oItems.Items);

		return oItems;
	};
	ReaderFromJSON.prototype.PivotRecordsFromJSON = function(aParsed, oPivotRecords)
	{
		for (var nElm = 0; nElm < aParsed.length; nElm++)
		{
			switch (aParsed[nElm]["type"])
			{
				case "b":
					this.PivotRecordFromJSON(aParsed[nElm], oPivotRecords._curBoolean);
					oPivotRecords.addBool(oPivotRecords._curBoolean.v, oPivotRecords._curBoolean);
					break;
				case "d":
					this.PivotRecordFromJSON(aParsed[nElm], oPivotRecords._curDateTime);
					oPivotRecords.addDate(oPivotRecords._curDateTime.v, oPivotRecords._curDateTime);
					break;
				case "e":
					this.PivotRecordFromJSON(aParsed[nElm], oPivotRecords._curError);
					oPivotRecords.addError(oPivotRecords._curError.v, oPivotRecords._curError);
					break;
				case "m":
					this.PivotRecordFromJSON(aParsed[nElm], oPivotRecords._curMissing);
					oPivotRecords.addMissing(oPivotRecords._curMissing.v, oPivotRecords._curMissing);
					break;
				case "n":
					this.PivotRecordFromJSON(aParsed[nElm], oPivotRecords._curNumber);
					oPivotRecords.addNumber(oPivotRecords._curNumber.v, oPivotRecords._curNumber);
					break;
				case "s":
					this.PivotRecordFromJSON(aParsed[nElm], oPivotRecords._curString);
					oPivotRecords.addString(oPivotRecords._curString.v, oPivotRecords._curString);
					break;
				case "i":
					this.PivotRecordFromJSON(aParsed[nElm], oPivotRecords._curIndex);
					oPivotRecords.addIndex(oPivotRecords._curIndex.v, oPivotRecords._curIndex);
					break;
			}
		}
	};
	ReaderFromJSON.prototype.PivotRecordFromJSON = function(oParsed, oPivotRecord)
	{
		// attributes
		if (oParsed["v"] != null)
			oPivotRecord.v = oParsed["v"];
		if (oParsed["u"] != null)
			oPivotRecord.u = oParsed["u"];
		if (oParsed["f"] != null)
			oPivotRecord.f = oParsed["f"];
		if (oParsed["c"] != null)
			oPivotRecord.c = oParsed["c"];
		if (oParsed["cp"] != null)
			oPivotRecord.cp = oParsed["cp"];
		if (oParsed["in"] != null)
			oPivotRecord.in = oParsed["in"];
		if (oParsed["bc"] != null)
			oPivotRecord.bc = oParsed["bc"];
		if (oParsed["fc"] != null)
			oPivotRecord.fc = oParsed["fc"];
		if (oParsed["i"] != null)
			oPivotRecord.i = oParsed["i"];
		if (oParsed["un"] != null)
			oPivotRecord.un = oParsed["un"];
		if (oParsed["st"] != null)
			oPivotRecord.st = oParsed["st"];
		if (oParsed["b"] != null)
			oPivotRecord.b = oParsed["b"];
		// members
		if (oParsed["tpls"] != null)
		{
			for (var nElm = 0; nElm < oParsed["tpls"].length; nElm++)
				oPivotRecord.tpls.push(this.TuplesFromJSON(oParsed["tpls"][nElm]));
		}
		if (oParsed["x"] != null)
		{
			for (var nElm = 0; nElm < oParsed["x"].length; nElm++)
				oPivotRecord.x.push(this.XFromJSON(oParsed["x"][nElm]));
		}

	};
	ReaderFromJSON.prototype.TuplesFromJSON = function(oParsed)
	{
		var oTuples = new CT_Tuples();

		if (oParsed["c"] != null)
			oTuples.c = oParsed["c"];

		for (var nElm = 0; nElm < oParsed["tpl"].length; nElm++)
			oTuples.tpl.push(this.TupleFromJSON(oParsed["tpl"][nElm]));

		return oTuples;
	};
	ReaderFromJSON.prototype.TupleFromJSON = function(oParsed)
	{
		var oTuple = new CT_Tuple();

		if (oParsed["fld"] != null)
			oTuple.fld = oParsed["fld"];
		if (oParsed["hier"] != null)
			oTuple.hier = oParsed["hier"];
		if (oParsed["item"] != null)
			oTuple.item = oParsed["item"];

		return oTuple;
	};
	ReaderFromJSON.prototype.FieldGroupFromJSON = function(oParsed)
	{
		var oGroup = new CT_FieldGroup();

		if (oParsed["par"] != null)
            oGroup.par = oParsed["par"];
        if (oParsed["base"] != null)
            oGroup.base = oParsed["base"];
        if (oParsed["rangePr"] != null)
            oGroup.rangePr = this.RangePrFromJSON(oParsed["rangePr"]);
        if (oParsed["discretePr"] != null)
            oGroup.discretePr = this.DiscretePrFromJSON(oParsed["discretePr"]);
        if (oParsed["groupItems"] != null)
            oGroup.groupItems = this.SharedItemsFromJSON(oParsed["groupItems"]);

		return oGroup;
	};
	ReaderFromJSON.prototype.RangePrFromJSON = function(oParsed)
	{
		var oPr = new CT_RangePr();
		
		if (oParsed["autoStart"] != null)
            oPr.autoStart = oParsed["autoStart"];
        if (oParsed["autoEnd"] != null)
            oPr.autoEnd = oParsed["autoEnd"];
        if (oParsed["groupBy"] != null)
            oPr.groupBy = FromXml_ST_GroupBy(oParsed["groupBy"]);
        if (oParsed["startNum"] != null)
            oPr.startNum = oParsed["startNum"];
        if (oParsed["endNum"] != null)
            oPr.endNum = oParsed["endNum"];
        if (oParsed["startDate"] != null)
            oPr.startDate = Asc.cDate.prototype.fromISO8601(oParsed["startDate"]);
        if (oParsed["endDate"] != null)
            oPr.endDate = Asc.cDate.prototype.fromISO8601(oParsed["endDate"]);
        if (oParsed["groupInterval"] != null)
            oPr.groupInterval = oParsed["groupInterval"];

		return oPr;
	};
	ReaderFromJSON.prototype.DiscretePrFromJSON = function(oParsed)
	{
		var oPr = new CT_DiscretePr();

		for (var nElm = 0; nElm < oParsed["x"].length; nElm++)
			oPr.x.push(this.IndexFromJSON(oParsed["x"][nElm]));

		return oPr;
	};
	ReaderFromJSON.prototype.CacheHierarchiesFromJSON = function(oParsed)
	{
		var oCacheHiers = new CT_CacheHierarchies();

		for (var nElm = 0; nElm < oParsed["cacheHierarchy"].length; nElm++)
			oCacheHiers.cacheHierarchy.push(this.CacheHierarchyFromJSON(oParsed["cacheHierarchy"][nElm]));

		return oCacheHiers;
	};
	ReaderFromJSON.prototype.CacheHierarchyFromJSON = function(oParsed)
	{
		var oCacheHier = new CT_CacheHierarchy();

		if (oParsed["uniqueName"] != null)
            oCacheHier.uniqueName = oParsed["uniqueName"];    
        if (oParsed["caption"] != null)
            oCacheHier.caption = oParsed["caption"];
        if (oParsed["measure"] != null)
            oCacheHier.measure = oParsed["measure"];
        if (oParsed["set"] != null)
            oCacheHier.set = oParsed["set"];
        if (oParsed["parentSet"] != null)
            oCacheHier.parentSet = oParsed["parentSet"];      
        if (oParsed["iconSet"] != null)
            oCacheHier.iconSet = oParsed["iconSet"];
        if (oParsed["attribute"] != null)
            oCacheHier.attribute = oParsed["attribute"];      
        if (oParsed["time"] != null)
            oCacheHier.time = oParsed["time"];
        if (oParsed["keyAttribute"] != null)
            oCacheHier.keyAttribute = oParsed["keyAttribute"];
        if (oParsed["defaultMemberUniqueName"] != null)
            oCacheHier.defaultMemberUniqueName = oParsed["defaultMemberUniqueName"];
        if (oParsed["allUniqueName"] != null)
            oCacheHier.allUniqueName = oParsed["allUniqueName"];
        if (oParsed["allCaption"] != null)
            oCacheHier.allCaption = oParsed["allCaption"];
        if (oParsed["dimensionUniqueName"] != null)
            oCacheHier.dimensionUniqueName = oParsed["dimensionUniqueName"];
        if (oParsed["displayFolder"] != null)
            oCacheHier.displayFolder = oParsed["displayFolder"];
        if (oParsed["measureGroup"] != null)
            oCacheHier.measureGroup = oParsed["measureGroup"];
        if (oParsed["measures"] != null)
            oCacheHier.measures = oParsed["measures"];
        if (oParsed["count"] != null)
            oCacheHier.count = oParsed["count"];
        if (oParsed["oneField"] != null)
            oCacheHier.oneField = oParsed["oneField"];
        if (oParsed["memberValueDatatype"] != null)
            oCacheHier.memberValueDatatype = oParsed["memberValueDatatype"];
        if (oParsed["unbalanced"] != null)
            oCacheHier.unbalanced = oParsed["unbalanced"];
        if (oParsed["unbalancedGroup"] != null)
            oCacheHier.unbalancedGroup = oParsed["unbalancedGroup"];
        if (oParsed["hidden"] != null)
            oCacheHier.hidden = oParsed["hidden"];
        if (oParsed["fieldsUsage"] != null)
            oCacheHier.fieldsUsage = this.FieldsUsageFromJSON(oParsed["fieldsUsage"]);
        if (oParsed["groupLevels"] != null)
            oCacheHier.groupLevels = this.GroupLevelsFromJSON(oParsed["groupLevels"]);
        if (oParsed["extLst"] != null)
            oCacheHier.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oCacheHier;
	};
	ReaderFromJSON.prototype.FieldsUsageFromJSON = function(oParsed)
	{
		var oFields = new CT_FieldsUsage();

		for (var nElm = 0; nElm < oParsed["fieldUsage"].length; nElm++)
			oFields.fieldUsage.push(this.FieldUsageFromJSON(oParsed["fieldUsage"][nElm]));

		return oFields;
	};
	ReaderFromJSON.prototype.FieldUsageFromJSON = function(oParsed)
	{
		var oField = new CT_FieldUsage();

		if (oParsed["x"] != null)
			oField.x = oParsed["x"];

		return oField;
	};
	ReaderFromJSON.prototype.GroupLevelsFromJSON = function(oParsed)
	{
		var oLevels = new CT_GroupLevels();

		for (var nElm = 0; nElm < oParsed["groupLevel"].length; nElm++)
			oLevels.groupLevel.push(this.GroupLevelFromJSON(oParsed["groupLevel"][nElm]));

		return oLevels;
	};
	ReaderFromJSON.prototype.GroupLevelFromJSON = function(oParsed)
	{
		var oLevel = new CT_GroupLevel();

		if (oParsed["uniqueName"] != null)
            oLevel.uniqueName = oParsed["uniqueName"];
        if (oParsed["caption"] != null)
            oLevel.caption = oParsed["caption"];
        if (oParsed["user"] != null)
            oLevel.user = oParsed["user"];
        if (oParsed["customRollUp"] != null)
            oLevel.customRollUp = oParsed["customRollUp"];
        if (oParsed["groups"] != null)
            oLevel.groups = this.GroupsFromJSON(oParsed["groups"]);
        if (oParsed["extLst"] != null)
            oLevel.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oLevel;
	};
	ReaderFromJSON.prototype.GroupsFromJSON = function(oParsed)
	{
		var oGroups = new CT_Groups();

		for (var nElm = 0; nElm < oParsed["group"].length; nElm++)
			oGroups.group.push(this.GroupFromJSON(oParsed["group"][nElm]));

		return oGroups;
	};
	ReaderFromJSON.prototype.GroupFromJSON = function(oParsed)
	{
		var oGroup = new CT_LevelGroup();

		if (oParsed["name"] != null)
            oGroup.name = oParsed["name"];
        if (oParsed["uniqueName"] != null)
            oGroup.uniqueName = oParsed["uniqueName"];    
        if (oParsed["caption"] != null)
            oGroup.caption = oParsed["caption"];
        if (oParsed["uniqueParent"] != null)
            oGroup.uniqueParent = oParsed["uniqueParent"];
        if (oParsed["id"] != null)
            oGroup.id = oParsed["id"];
        if (oParsed["groupMembers"] != null)
            oGroup.groupMembers = this.GroupMembersFromJSON(oParsed["groupMembers"]);

		return oGroup;
	};
	ReaderFromJSON.prototype.GroupMembersFromJSON = function(oParsed)
	{
		var oMembers = new CT_GroupMembers();

		for (var nElm = 0; nElm < oParsed["groupMember"].length; nElm++)
			oMembers.groupMember.push(this.GroupMemberFromJSON(oParsed["groupMember"][nElm]));

		return oMembers;
	};
	ReaderFromJSON.prototype.GroupMemberFromJSON = function(oParsed)
	{
		var oMember = new CT_GroupMember();

		if (oParsed["uniqueName"] != null)
			oMember.uniqueName = oParsed["uniqueName"];
		if (oParsed["group"] != null)
			oMember.group = oParsed["group"];
		
		return oMember;
	};
	ReaderFromJSON.prototype.PCDKPIsFromJSON = function(oParsed)
	{
		var oKPIs = new CT_PCDKPIs();

		for (var nElm = 0; nElm < oParsed["kpi"].length; nElm++)
			oKPIs.kpi.push(this.KPIFromJSON(oParsed["kpi"][nElm]));

		return oKPIs;
	};
	ReaderFromJSON.prototype.KPIFromJSON = function(oParsed)
	{
		var oKPI = new CT_PCDKPI();

		if (oParsed["uniqueName"] != null)
            oKPI.uniqueName = oParsed["uniqueName"];
        if (oParsed["caption"] != null)
            oKPI.caption = oParsed["caption"];      
        if (oParsed["displayFolder"] != null)
            oKPI.displayFolder = oParsed["displayFolder"];
        if (oParsed["measureGroup"] != null)
            oKPI.measureGroup = oParsed["measureGroup"];
        if (oParsed["parent"] != null)
            oKPI.parent = oParsed["parent"];
        if (oParsed["value"] != null)
            oKPI.value = oParsed["value"];
        if (oParsed["goal"] != null)
            oKPI.goal = oParsed["goal"];
        if (oParsed["status"] != null)
            oKPI.status = oParsed["status"];
        if (oParsed["trend"] != null)
            oKPI.trend = oParsed["trend"];
        if (oParsed["weight"] != null)
            oKPI.weight = oParsed["weight"];
        if (oParsed["time"] != null)
            oKPI.time = oParsed["time"];

		return oKPI;
	};
	ReaderFromJSON.prototype.TupleCacheFromJSON = function(oParsed)
	{
		var oCache = new CT_TupleCache();

		if (oParsed["entries"] != null)
            oCache.entries = this.PCDSDTCEntriesFromJSON(oParsed["entries"]);
        if (oParsed["sets"] != null)
            oCache.sets = this.SetsFromJSON(oParsed["sets"]);
        if (oParsed["queryCache"] != null)
            oCache.queryCache = this.QueryCacheFromJSON(oParsed["queryCache"]);
        if (oParsed["serverFormats"] != null)
            oCache.serverFormats = this.ServerFormatsFromJSON(oParsed["serverFormats"]);
        if (oParsed["extLst"] != null)
            oCache.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oCache;
	};
	ReaderFromJSON.prototype.PCDSDTCEntriesFromJSON = function(oParsed)
	{
		var oEntries = new CT_PCDSDTCEntries();

		for (var nElm = 0; nElm < oParsed["items"].length; nElm++)
			oEntries.Items.push(this.PCDSDTCItemFromJSON(oParsed["items"][nElm]));

		return oEntries;
	};
	ReaderFromJSON.prototype.PCDSDTCItemFromJSON = function(oParsed)
	{
		var oItem;
		switch (oParsed["type"])
		{
			case "e":
				oItem = new CT_Error();
				this.PivotRecordFromJSON(oParsed, oItem);
				break;
			case "m":
				oItem = new CT_Error();
				this.PivotRecordFromJSON(oParsed, oItem);
				break;
			case "n":
				oItem = new CT_Error();
				this.PivotRecordFromJSON(oParsed, oItem);
				break;
			case "se":
				oItem = new CT_Error();
				this.PivotRecordFromJSON(oParsed, oItem);
				break;
		}

		return oItem;
	};
	ReaderFromJSON.prototype.SetsFromJSON = function(oParsed)
	{
		var oSets = new CT_Sets();

		for (var nElm = 0; nElm < oParsed["set"].length; nElm++)
			oSets.set.push(this.SetFromJSON(oParsed["set"][nElm]));

		return oSets;
	};
	ReaderFromJSON.prototype.SetFromJSON = function(oParsed)
	{
		var oSet = new CT_Set();

		if (oParsed["count"] != null)
            oSet.count = oParsed["count"];
        if (oParsed["maxRank"] != null)
            oSet.maxRank = oParsed["maxRank"];
        if (oParsed["setDefinition"] != null)
            oSet.setDefinition = oParsed["setDefinition"];
        if (oParsed["sortType"] != null)
            oSet.sortType = FromXml_ST_SortType(oParsed["sortType"]);
        if (oParsed["queryFailed"] != null)
            oSet.queryFailed = oParsed["queryFailed"];
        if (oParsed["tpls"] != null)
		{
			for (var nElm = 0; nElm < oParsed["tpls"].length; nElm++)
				oSet.tpls.push(this.TuplesFromJSON(oParsed["tpls"][nElm]));
		}
        if (oParsed["sortByTuple"] != null)
            oSet.sortByTuple = oParsed["sortByTuple"];

		return oSet;
	};
	ReaderFromJSON.prototype.QueryCacheFromJSON = function(oParsed)
	{
		var oQuery = new CT_QueryCache();

		for (var nElm = 0; nElm < oParsed["query"].length; nElm++)
			oQuery.query.push(this.QueryFromJSON(oParsed["query"][nElm]));

		return oQuery;
	};
	ReaderFromJSON.prototype.QueryFromJSON = function(oParsed)
	{
		var oQuery = new CT_Query();

		if (oParsed["mdx"] != null)
            oQuery.mdx = oParsed["mdx"];
        if (oParsed["tpls"] != null)
            oQuery.tpls = this.TuplesFromJSON(oParsed["tpls"]);

		return oQuery;
	};
	ReaderFromJSON.prototype.ServerFormatsFromJSON = function(oParsed)
	{
		var oFormats = new CT_ServerFormats();

		for (var nElm = 0; nElm < oParsed["serverFormat"].length; nElm++)
			oFormats.serverFormat.push(this.ServerFormatFromJSON(oParsed["serverFormat"][nElm]));

		return oFormats;
	};
	ReaderFromJSON.prototype.ServerFormatFromJSON = function(oParsed)
	{
		var oFormat = new CT_ServerFormat();

		if (oParsed["culture"] != null)
            oFormat.culture = oParsed["culture"];
		if (oParsed["format"] != null)
            oFormat.format = oParsed["format"];

		return oFormat;
	};
	ReaderFromJSON.prototype.CalculatedItemsFromJSON = function(oParsed)
	{
		var oItems = new CT_CalculatedItems();

		for (var nElm = 0; nElm < oParsed["calculatedItem"].length; nElm++)
			oItems.calculatedItem.push(this.CalculatedItemFromJSON(oParsed["calculatedItem"][nElm]));

		return oItems;
	};
	ReaderFromJSON.prototype.CalculatedItemFromJSON = function(oParsed)
	{
		var oItem = new CT_CalculatedItem();

		if (oParsed["field"] != null)
            oItem.field = oParsed["field"];
        if (oParsed["formula"] != null)
            oItem.formula = oParsed["formula"];
        if (oParsed["pivotArea"] != null)
            oItem.pivotArea = this.PivotAreaFromJSON(oParsed["pivotArea"]);
        if (oParsed["extLst"] != null)
            oItem.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oItem;
	};
	ReaderFromJSON.prototype.CalculatedMembersFromJSON = function(oParsed)
	{
		var oMembers = new CT_CalculatedMembers();

		for (var nElm = 0; nElm < oParsed["calculatedMember"].length; nElm++)
			oMembers.calculatedMember.push(this.CalculatedMemberFromJSON(oParsed["calculatedMember"][nElm]));

		return oMembers;
	};
	ReaderFromJSON.prototype.CalculatedMemberFromJSON = function(oParsed)
	{
		var oMember = new CT_CalculatedMember();
	
		if (oParsed["name"] != null)
            oMember.name = oParsed["name"];
        if (oParsed["mdx"] != null)
            oMember.mdx = oParsed["mdx"];
        if (oParsed["memberName"] != null)
            oMember.memberName = oParsed["memberName"];
        if (oParsed["hierarchy"] != null)
            oMember.hierarchy = oParsed["hierarchy"];  
        if (oParsed["parent"] != null)
            oMember.parent = oParsed["parent"];        
        if (oParsed["solveOrder"] != null)
            oMember.solveOrder = oParsed["solveOrder"];
        if (oParsed["set"] != null)
            oMember.set = oParsed["set"];
        if (oParsed["extLst"] != null)
            oMember.extLst = this.ExtensionListFromJSON(oParsed["extLst"]);

		return oMember;
	};
	ReaderFromJSON.prototype.DimensionsFromJSON = function(oParsed)
	{
		var oDimensions = new CT_Dimensions();

		for (var nElm = 0; nElm < oParsed["dimension"].length; nElm++)
			oDimensions.dimension.push(this.DimensionFromJSON(oParsed["dimension"][nElm]));

		return oDimensions;
	};
	ReaderFromJSON.prototype.DimensionFromJSON = function(oParsed)
	{
		var oDimension = new CT_PivotDimension();

		if (oParsed["measure"] != null)
            oDimension.measure = oParsed["measure"];      
        if (oParsed["name"] != null)
            oDimension.name = oParsed["name"];
        if (oParsed["uniqueName"] != null)
            oDimension.uniqueName = oParsed["uniqueName"];
        if (oParsed["caption"] != null)
            oDimension.caption = oParsed["caption"];

		return oDimension;
	};
	ReaderFromJSON.prototype.MeasureGroupsFromJSON = function(oParsed)
	{
		var oGroups = new CT_MeasureGroups();

		for (var nElm = 0; nElm < oParsed["measureGroup"].length; nElm++)
			oGroups.measureGroup.push(this.MeasureGroupFromJSON(oParsed["measureGroup"][nElm]));

		return oGroups;
	};
	ReaderFromJSON.prototype.MeasureGroupFromJSON = function(oParsed)
	{
		var oGroup = new CT_MeasureGroup();

		if (oParsed["name"] != null)
			oGroup.name = oParsed["name"];
		if (oParsed["caption"] != null)
			oGroup.caption = oParsed["caption"];

		return oGroup;
	};
	ReaderFromJSON.prototype.MeasureDimensionMapsFromJSON = function(oParsed)
	{
		var oMaps = new CT_MeasureDimensionMaps();

		for (var nElm = 0; nElm < oParsed["map"].length; nElm++)
			oMaps.map.push(this.MeasureDimensionMapFromJSON(oParsed["map"][nElm]));

		return oMaps;
	};
	ReaderFromJSON.prototype.MeasureDimensionMapFromJSON = function(oParsed)
	{
		var oMap = new CT_MeasureDimensionMap();

		if (oParsed["measureGroup"] != null)
			oMap.measureGroup = oParsed["measureGroup"];
		if (oParsed["dimension"] != null)
			oMap.dimension = oParsed["dimension"];

		return oMap;
	};
	ReaderFromJSON.prototype.TabularSlicerCacheFromJSON = function(oParsed)
	{
		var oCache = new Asc.CT_tabularSlicerCache();

		if (oParsed["pivotCacheId"] != null)
            oCache.pivotCacheId = oParsed["pivotCacheId"];
        if (oParsed["pivotCacheDefinition"] != null)
            oCache.pivotCacheDefinition = this.PivotCacheDefinitionFromJSON(oParsed["pivotCacheDefinition"]);
        if (oParsed["sortOrder"] != null)
            oCache.sortOrder = FromXML_ST_tabularSlicerCacheSortOrder(oParsed["sortOrder"]);
        if (oParsed["customListSort"] != null)
            oCache.customListSort = oParsed["customListSort"];
        if (oParsed["showMissing"] != null)
            oCache.showMissing = oParsed["showMissing"];
        if (oParsed["crossFilter"] != null)
            oCache.crossFilter = FromXML_ST_slicerCacheCrossFilter(oParsed["crossFilter"]);
        if (oParsed["items"] != null)
            oCache.items = this.TabularSlicerCacheItemsFromJSON(oParsed["items"]);

		return oCache;
	};
	ReaderFromJSON.prototype.TabularSlicerCacheItemsFromJSON = function(aParsed)
	{
		var aResult = [];
		for (var nElm = 0; nElm < aParsed.length; nElm++)
			aResult.push(this.TabularSlicerCacheItemFromJSON(aParsed[nElm]));

		return aResult;
	};
	ReaderFromJSON.prototype.TabularSlicerCacheItemFromJSON = function(oParsed)
	{
		var oItem = new Asc.CT_tabularSlicerCacheItem();

		if (oParsed["x"] != null)
            oItem.x = oParsed["x"];
		if (oParsed["s"] != null)
            oItem.s = oParsed["s"];
		if (oParsed["nd"] != null)
            oItem.nd = oParsed["nd"];

		return oItem;
	};
	ReaderFromJSON.prototype.TableSlicerCacheFromJSON = function(oParsed)
	{
		var oCache = new Asc.CT_tableSlicerCache();

		if (oParsed["tableId"] != null)
            oCache.tableId = oParsed["tableId"];
        if (oParsed["tableIdOpen"] != null)
            oCache.tableIdOpen = oParsed["tableIdOpen"];      
        if (oParsed["column"] != null)
            oCache.column = oParsed["column"];
        if (oParsed["columnOpen"] != null)
            oCache.columnOpen = oParsed["columnOpen"];        
        if (oParsed["sortOrder"] != null)
            oCache.sortOrder = FromXML_ST_tabularSlicerCacheSortOrder(oParsed["sortOrder"]);
        if (oParsed["customListSort"] != null)
            oCache.customListSort = oParsed["customListSort"];
        if (oParsed["crossFilter"] != null)
            oCache.crossFilter = FromXML_ST_slicerCacheCrossFilter(oParsed["crossFilter"]);

		return oCache;
	};
	ReaderFromJSON.prototype.SlicerCacheHideNoDataFromJSON = function(oParsed)
	{
		var oCache = new Asc.CT_slicerCacheHideNoData();

		if (oParsed["count"] != null)
			oCache.count = oParsed["count"];
			
		for (var nElm = 0; nElm < oParsed["slicerCacheOlapLevelName"].length; nElm++)
			oCache.slicerCacheOlapLevelName.push(this.SlicerCacheOlapLevelNameFromJSON(oParsed["slicerCacheOlapLevelName"][nElm]));

		return oCache;
	};
	ReaderFromJSON.prototype.SlicerCacheOlapLevelNameFromJSON = function(oParsed)
	{
		var oLvlName = new Asc.CT_slicerCacheOlapLevelName();

		if (oParsed["uniqueName"] != null)
			oLvlName.uniqueName = oParsed["uniqueName"];
		if (oParsed["count"] != null)
			oLvlName.count = oParsed["count"];

		return oLvlName;
	};
	ReaderFromJSON.prototype.NamedSheetViewsFromJSON = function(aParsed, oWorksheet)
	{
		var oNamedSheetViews = Asc.CT_NamedSheetViews ? new Asc.CT_NamedSheetViews() : null;
		if (!oNamedSheetViews)
			return;

		oWorksheet.aNamedSheetViews = oNamedSheetViews.namedSheetView;
		for (var nView = 0; nView < aParsed.length; nView++)
			oNamedSheetViews.namedSheetView.push(this.NamedSheetViewFromJSON(aParsed[nView]));
	};
	ReaderFromJSON.prototype.NamedSheetViewFromJSON = function(oParsed)
	{
		var oView = new Asc.CT_NamedSheetView();

		if (oParsed["name"] != null)
            oView.name = oParsed["name"];
        if (oParsed["id"] != null)
            oView.id = oParsed["id"];
        if (oParsed["nsvFilters"] != null)
		{
			for (var nElm = 0; nElm < oParsed["nsvFilters"].length; nElm++)
				oView.nsvFilters.push(this.NsvFilterFromJSON(oParsed["nsvFilters"][nElm]));
		}

		return oView;
	};
	ReaderFromJSON.prototype.NsvFilterFromJSON = function(oParsed)
	{
		var oFilter = new Asc.CT_NsvFilter();

		if (oParsed["filterId"] != null)
            oFilter.filterId = oParsed["filterId"];
        if (oParsed["ref"] != null)
            oFilter.ref = AscCommonExcel.g_oRangeCache.getAscRange(oParsed["ref"]);
        if (oParsed["tableIdOpen"] != null)
            oFilter.tableIdOpen = oParsed["tableIdOpen"];    
        //if (oParsed["tableId"] != null)
        //    oFilter.tableId = oParsed["tableId"];
        if (oParsed["columnsFilter"] != null)
		{
			for (var nElm = 0; nElm < oParsed["columnsFilter"].length; nElm++)
				oFilter.columnsFilter.push(this.FilterColumnFromJSON(oParsed["columnsFilter"][nElm]));
		}
        if (oParsed["sortRules"] != null)
            oFilter.sortRules = this.SortRulesFromJSON(oParsed["sortRules"]);

		return oFilter;
	};
	ReaderFromJSON.prototype.SortRulesFromJSON = function(oParsed)
	{
		var aResult = [];
		for (var nElm = 0; nElm < oParsed["sortRule"].length; nElm++)
			aResult.push(this.SortRuleFromJSON(oParsed["sortRule"][nElm]));

		return aResult;
	};
	ReaderFromJSON.prototype.SortRuleFromJSON = function(oParsed)
	{
		var oRule = new Asc.CT_SortRule();

		if (oParsed["colId"] != null)
            oRule.colId = oParsed["colId"];
        if (oParsed["id"] != null)
            oRule.id = oParsed["id"];
        if (oParsed["sortCondition"] != null)
            oRule.sortCondition = this.SortConditionFromJSON(oParsed["sortCondition"]); // to do

		return oRule;
	};
	ReaderFromJSON.prototype.SheetProtectionFromJSON = function(oParsed, oWorksheet)
	{
		var oSheetProtection = Asc.CSheetProtection ? new Asc.CSheetProtection(oWorksheet) : null;
		if (!oSheetProtection)
			return;

		oWorksheet.sheetProtection = oSheetProtection;

		if (oParsed["algorithmName"] != null)
            oSheetProtection.algorithmName = From_XML_CryptoAlgorithmName(oParsed["algorithmName"]);
        if (oParsed["autoFilter"] != null)
            oSheetProtection.autoFilter = oParsed["autoFilter"];      
        if (oParsed["deleteColumns"] != null)
            oSheetProtection.deleteColumns = oParsed["deleteColumns"];
        if (oParsed["deleteRows"] != null)
            oSheetProtection.deleteRows = oParsed["deleteRows"];      
        if (oParsed["formatCells"] != null)
            oSheetProtection.formatCells = oParsed["formatCells"];    
        if (oParsed["formatColumns"] != null)
            oSheetProtection.formatColumns = oParsed["formatColumns"];
        if (oParsed["formatRows"] != null)
            oSheetProtection.formatRows = oParsed["formatRows"];      
        if (oParsed["hashValue"] != null)
            oSheetProtection.hashValue = oParsed["hashValue"];        
        if (oParsed["insertColumns"] != null)
            oSheetProtection.insertColumns = oParsed["insertColumns"];
        if (oParsed["insertHyperlinks"] != null)
            oSheetProtection.insertHyperlinks = oParsed["insertHyperlinks"];
        if (oParsed["insertRows"] != null)
            oSheetProtection.insertRows = oParsed["insertRows"];
        if (oParsed["objects"] != null)
            oSheetProtection.objects = oParsed["objects"];
        if (oParsed["pivotTables"] != null)
            oSheetProtection.pivotTables = oParsed["pivotTables"];
        if (oParsed["saltValue"] != null)
            oSheetProtection.saltValue = oParsed["saltValue"];
        if (oParsed["scenarios"] != null)
            oSheetProtection.scenarios = oParsed["scenarios"];
        if (oParsed["selectLockedCells"] != null)
            oSheetProtection.selectLockedCells = oParsed["selectLockedCells"];
        if (oParsed["selectUnlockedCells"] != null)
            oSheetProtection.selectUnlockedCells = oParsed["selectUnlockedCells"];
        if (oParsed["sheet"] != null)
            oSheetProtection.sheet = oParsed["sheet"];
        if (oParsed["sort"] != null)
            oSheetProtection.sort = oParsed["sort"];
        if (oParsed["spinCount"] != null)
            oSheetProtection.spinCount = oParsed["spinCount"];
        if (oParsed["password"] != null)
            oSheetProtection.password = oParsed["password"];
        if (oParsed["content"] != null)
            oSheetProtection.content = oParsed["content"];
	};
	ReaderFromJSON.prototype.ProtectedRangesFromJSON = function(aParsed)
	{
		var aResult = [];

		for (var nElm = 0; nElm < aParsed.length; nElm++)
			aResult.push(this.ProtectedRangeFromJSON(aParsed[nElm]));

		return aResult;
	};
	ReaderFromJSON.prototype.ProtectedRangeFromJSON = function(oParsed)
	{
		var oRange = Asc.CProtectedRange ? new Asc.CProtectedRange() : null;
		if (!oRange)
			return oRange;

		if (oParsed["algorithmName"] != null)
            oRange.algorithmName = From_XML_CryptoAlgorithmName(oParsed["algorithmName"]);
        if (oParsed["spinCount"] != null)
            oRange.spinCount = oParsed["spinCount"];        
        if (oParsed["hashValue"] != null)
            oRange.hashValue = oParsed["hashValue"];        
        if (oParsed["saltValue"] != null)
            oRange.saltValue = oParsed["saltValue"];        
        if (oParsed["name"] != null)
            oRange.name = oParsed["name"];
        if (oParsed["sqref"] != null)
            oRange.sqref = AscCommonExcel.g_oRangeCache.getRangesFromSqRef(oParsed["sqref"]);
		if (oParsed["securityDescriptor"] != null)
            oRange.securityDescriptor = oParsed["securityDescriptor"];

		return oRange;
	};
	ReaderFromJSON.prototype.CommentsFromJSON = function(aParsed, oWorksheet)
	{
		var oCommentCoords, oCommentData;
		for (var nElm = 0; nElm < aParsed.length; nElm++)
		{
			oCommentCoords = this.CommentCoordsFromJSON(aParsed[nElm]["coord"]);
			oCommentData = this.CommentDataFromJSON(aParsed[nElm]["data"], oWorksheet);
			oCommentData.asc_putRow(oCommentCoords.nRow);
			oCommentData.asc_putCol(oCommentCoords.nCol);
			oWorksheet.aComments.push(oCommentData);
		}
	};
	ReaderFromJSON.prototype.CommentCoordsFromJSON = function(oParsed)
	{
		var oCommentCoords = new AscCommonExcel.asc_CCommentCoords();

		if (oParsed["row"] != null)
            oCommentCoords.nRow = oParsed["row"];
        if (oParsed["col"] != null)
            oCommentCoords.nCol = oParsed["col"];
        if (oParsed["left"] != null)
            oCommentCoords.nLeft = oParsed["left"];
        if (oParsed["leftOff"] != null)
            oCommentCoords.nLeftOffset = oParsed["leftOff"];    
        if (oParsed["topOff"] != null)
            oCommentCoords.nTopOffset = oParsed["topOff"];      
        if (oParsed["right"] != null)
            oCommentCoords.nRight = oParsed["right"];        
        if (oParsed["rightOff"] != null)
            oCommentCoords.nRightOffset = oParsed["rightOff"];  
        if (oParsed["bottom"] != null)
            oCommentCoords.nBottom = oParsed["bottom"];      
        if (oParsed["bottomOff"] != null)
            oCommentCoords.nBottomOffset = oParsed["bottomOff"];
        if (oParsed["leftMM"] != null)
            oCommentCoords.dLeftMM = oParsed["leftMM"];
        if (oParsed["topMM"] != null)
            oCommentCoords.dTopMM = oParsed["topMM"];
        if (oParsed["widthMM"] != null)
            oCommentCoords.dWidthMM = oParsed["widthMM"];
        if (oParsed["heightMM"] != null)
            oCommentCoords.dHeightMM = oParsed["heightMM"];
        if (oParsed["moveWithCells"] != null)
            oCommentCoords.bMoveWithCells = oParsed["moveWithCells"];
        if (oParsed["sizeWithCells"] != null)
            oCommentCoords.bSizeWithCells = oParsed["sizeWithCells"];
	
		return oCommentCoords;
	};
	ReaderFromJSON.prototype.CommentDataFromJSON = function(oParsed, oWorksheet)
	{
		var oCommentData = new Asc.asc_CCommentData();
		oCommentData.asc_putDocumentFlag(false);

		if (oParsed["text"] != null)
			oCommentData.asc_putText(oParsed["text"]);
		if (oParsed["dt"] != null)
		{
			var dateMs = AscCommon.getTimeISO8601(oParsed["dt"]);
			if(!isNaN(dateMs))
				oCommentData.asc_putOnlyOfficeTime(dateMs + "");
		}
		if (oParsed["id"] != null)
			oCommentData.asc_putGuid(oParsed["id"]);
		// if (oParsed["id"] != null)
		// 	oCommentData.asc_putGuid(AscCommon.CreateGUID());
		if (oParsed["displayName"] != null)
			oCommentData.asc_putUserName(oParsed["displayName"]);
		if (oParsed["userId"] != null)
			oCommentData.asc_putUserId(oParsed["userId"]);
		if (oParsed["providerId"] != null)
			oCommentData.asc_putProviderId(oParsed["providerId"]);
		if (oParsed["done"] != null)
			oCommentData.asc_putSolved(oParsed["done"]);
		if (oParsed["replies"] != null)
		{
			for (var nElm = 0; nElm < oParsed["replies"].length; nElm++)
				oCommentData.asc_addReply(this.CommentDataFromJSON(oParsed["replies"][nElm]));
		}

		oCommentData.asc_putDocumentFlag(false);

		oCommentData.wsId = oWorksheet.Id;
		oCommentData.nId = "sheet" + oCommentData.wsId + "_" + (oWorksheet.aComments.length + 1);

		return oCommentData;
	};
	ReaderFromJSON.prototype.SheetDataFromJSON = function(aParsed, oWorksheet)
	{
		let oThis = this;
		var oParsedRow, oParsedCell, oCellAddress;
		var oTempRow, oTempCell;

		var tmp = {
			pos: null, len: null, bNoBuildDep: false, ws: oWorksheet, row: null,
			cell: null, formula: null, sharedFormulas: {},
			prevFormulas: {}, siFormulas: {}, prevRow: -1, prevCol: -1, formulaArray: []
		};

		for (var nRow = 0; nRow < aParsed.length; nRow++)
		{
			oParsedRow = aParsed[nRow];
			oTempRow = new AscCommonExcel.Row(oWorksheet);
			oTempRow.loadContent(oParsedRow["r"] - 1);
			oWorksheet._initRow(oTempRow, oParsedRow["r"] - 1);
			tmp.row = oTempRow;

			if (oParsedRow["collapsed"] != null && oParsedRow["collapsed"] !== false)
				oTempRow.setCollapsed(oParsedRow["collapsed"]);
			if (oParsedRow["customHeight"] != null && oParsedRow["customHeight"] !== false)
				oTempRow.setCustomHeight(oParsedRow["customHeight"]);
			if (oParsedRow["hidden"] != null && oParsedRow["hidden"] !== false)
				oTempRow.setHidden(oParsedRow["hidden"]);
			if (oParsedRow["ht"] != null && oParsedRow["ht"] !== false)
				oTempRow.setHeight(oParsedRow["ht"]);
			if (oParsedRow["outlineLevel"] != null && oParsedRow["outlineLevel"] !== false)
				oTempRow.setOutlineLevel(oParsedRow["outlineLevel"]);
			if (oParsedRow["s"] != null && oParsedRow["s"] !== 0)
				oTempRow.setStyle(this.aCellXfs[oParsedRow["s"]]);
			
			oTempRow.saveContent(true);

			for (var nCell = 0; nCell < oParsedRow["cell"].length; nCell++)
			{
				oParsedCell = oParsedRow["cell"][nCell];
				oTempCell = new AscCommonExcel.Cell(oWorksheet);
				oCellAddress = AscCommon.g_oCellAddressUtils.getCellAddress(oParsedCell["r"]);
				oTempCell.loadContent(oParsedRow["r"] - 1, oCellAddress.col - 1);
				oWorksheet._initCell(oTempCell, oParsedRow["r"] - 1, oCellAddress.col - 1);

				if (oParsedCell["f"] != null)
				{
					tmp.cell = oTempCell;
					tmp.formula = FormulaFromJSON(oParsedCell["f"]);
					this.InitOpenManager.setFormulaOpen(tmp);
				}
				// multiText
				else if (Array.isArray(oParsedCell["v"]))
				{
					let aMultiText = [];
					for (let Index = 0; Index < oParsedCell["v"].length; Index++)
						aMultiText.push(MultiTextElemFromJSON(oParsedCell["v"][Index]));
				
					oTempCell.setValueMultiTextInternal(aMultiText);
				}
				else if (oParsedCell["v"] != null)
				{
					switch (oParsedCell["t"])
					{
						case "s":
							oTempCell.setValueTextInternal(oParsedCell["v"]);
							break;
						case "e":
							oTempCell.setValueTextInternal(oParsedCell["v"]);
							break;
						case "n":
						case "b":
						default:
							oTempCell.setValueNumberInternal(oParsedCell["v"]);
							break;
					}
				}
				if (oParsedCell["s"] != null && oParsedCell["s"] !== 0)
					oTempCell.setStyle(this.aCellXfs[oParsedCell["s"]]);

				oTempCell.type = FromXml_ST_CellValueType(oParsedCell["t"]);
				oTempCell.saveContent(true);
			}
		}

		// for(var j = 0; j < tmp.formulaArray.length; j++) {
		// 	var curFormula = tmp.formulaArray[j];
		// 	var ref = curFormula.ref;
		// 	if(ref) {
		// 		var rangeFormulaArray = tmp.ws.getRange3(ref.r1, ref.c1, ref.r2, ref.c2);
		// 		rangeFormulaArray._foreach(function(cell){
		// 			cell.setFormulaInternal(curFormula);
		// 			if (curFormula.ca || cell.isNullTextString()) {
		// 				tmp.ws.workbook.dependencyFormulas.addToChangedCell(cell);
		// 			}
		// 		});
		// 	}
		// }
		// for (var nCol in tmp.prevFormulas) {
		// 	if (tmp.prevFormulas.hasOwnProperty(nCol)) {
		// 		var prevFormula = tmp.prevFormulas[nCol];
		// 		if (!tmp.siFormulas[prevFormula.parsed.getListenerId()]) {
		// 			prevFormula.parsed.buildDependencies();
		// 		}
		// 	}
		// }
		// for (var listenerId in tmp.siFormulas) {
		// 	if (tmp.siFormulas.hasOwnProperty(listenerId)) {
		// 		tmp.siFormulas[listenerId].buildDependencies();
		// 	}
		// }

		function FormulaFromJSON(oParsed)
		{
			let oFormula = new AscCommonExcel.OpenFormula();
			oFormula.t = FromXml_ST_CellFormulaType(oParsed["t"]);
			oFormula.v = oParsed["v"];
			if (oParsed["ref"] != null)
				oFormula.ref = oParsed["ref"];
			if (oParsed["ca"] != null)
				oFormula.ca = oParsed["ca"];
			if (oParsed["si"] != null)
				oFormula.si = oParsed["si"];

			return oFormula;
		}

		function MultiTextElemFromJSON(oParsed)
		{
			let oElem = new AscCommonExcel.CMultiTextElem();
			oElem.text = oParsed["text"];
			oElem.format = oThis.FontExcellFromJSON(oParsed["format"]);

			return oElem;
		}
	};

	function FromXml_ST_FilterOperator(val) {
		var res = -1;
		if ("equal" === val) {
			res = Asc.c_oAscCustomAutoFilter.equals;
		} else if ("lessThan" === val) {
			res = Asc.c_oAscCustomAutoFilter.isLessThan;
		} else if ("lessThanOrEqual" === val) {
			res = Asc.c_oAscCustomAutoFilter.isLessThanOrEqualTo;
		} else if ("notEqual" === val) {
			res = Asc.c_oAscCustomAutoFilter.doesNotEqual;
		} else if ("greaterThanOrEqual" === val) {
			res = Asc.c_oAscCustomAutoFilter.isGreaterThanOrEqualTo;
		} else if ("greaterThan" === val) {
			res = Asc.c_oAscCustomAutoFilter.isGreaterThan;
		}
		return res;
	}
	function ToXml_ST_FilterOperator(val) {
		var res = "";
		if (Asc.c_oAscCustomAutoFilter.equals === val) {
			res = "equal";
		} else if (Asc.c_oAscCustomAutoFilter.isLessThan === val) {
			res = "lessThan";
		} else if (Asc.c_oAscCustomAutoFilter.isLessThanOrEqualTo === val) {
			res = "lessThanOrEqual";
		} else if (Asc.c_oAscCustomAutoFilter.doesNotEqual === val) {
			res = "notEqual";
		} else if (Asc.c_oAscCustomAutoFilter.isGreaterThanOrEqualTo === val) {
			res = "greaterThanOrEqual";
		} else if (Asc.c_oAscCustomAutoFilter.isGreaterThan === val) {
			res = "greaterThan";
		}
		return res;
	}

	function FromXml_ST_DynamicFilterType(val) {
		var res = -1;
		if ("null" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nullType;
		} else if ("aboveAverage" === val) {
			res = Asc.c_oAscDynamicAutoFilter.aboveAverage;
		} else if ("belowAverage" === val) {
			res = Asc.c_oAscDynamicAutoFilter.belowAverage;
		} else if ("tomorrow" === val) {
			res = Asc.c_oAscDynamicAutoFilter.tomorrow;
		} else if ("today" === val) {
			res = Asc.c_oAscDynamicAutoFilter.today;
		} else if ("yesterday" === val) {
			res = Asc.c_oAscDynamicAutoFilter.yesterday;
		} else if ("nextWeek" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nextWeek;
		} else if ("thisWeek" === val) {
			res = Asc.c_oAscDynamicAutoFilter.thisWeek;
		} else if ("lastWeek" === val) {
			res = Asc.c_oAscDynamicAutoFilter.lastWeek;
		} else if ("nextMonth" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nextMonth;
		} else if ("thisMonth" === val) {
			res = Asc.c_oAscDynamicAutoFilter.thisMonth;
		} else if ("lastMonth" === val) {
			res = Asc.c_oAscDynamicAutoFilter.lastMonth;
		} else if ("nextQuarter" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nextQuarter;
		} else if ("thisQuarter" === val) {
			res = Asc.c_oAscDynamicAutoFilter.thisQuarter;
		} else if ("lastQuarter" === val) {
			res = Asc.c_oAscDynamicAutoFilter.lastQuarter;
		} else if ("nextYear" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nextYear;
		} else if ("thisYear" === val) {
			res = Asc.c_oAscDynamicAutoFilter.thisYear;
		} else if ("lastYear" === val) {
			res = Asc.c_oAscDynamicAutoFilter.lastYear;
		} else if ("yearToDate" === val) {
			res = Asc.c_oAscDynamicAutoFilter.yearToDate;
		} else if ("Q1" === val) {
			res = Asc.c_oAscDynamicAutoFilter.q1;
		} else if ("Q2" === val) {
			res = Asc.c_oAscDynamicAutoFilter.q2;
		} else if ("Q3" === val) {
			res = Asc.c_oAscDynamicAutoFilter.q3;
		} else if ("Q4" === val) {
			res = Asc.c_oAscDynamicAutoFilter.q4;
		} else if ("M1" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m1;
		} else if ("M2" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m2;
		} else if ("M3" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m3;
		} else if ("M4" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m4;
		} else if ("M5" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m5;
		} else if ("M6" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m6;
		} else if ("M7" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m7;
		} else if ("M8" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m8;
		} else if ("M9" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m9;
		} else if ("M10" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m10;
		} else if ("M11" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m11;
		} else if ("M12" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m12;
		}
		return res;
	}
	function ToXml_ST_DynamicFilterType(val) {
		var res = "";
		if (Asc.c_oAscDynamicAutoFilter.nullType === val){
			res = "null";
		} else if (Asc.c_oAscDynamicAutoFilter.aboveAverage === val) {
			res = "aboveAverage";
		} else if (Asc.c_oAscDynamicAutoFilter.belowAverage === val) {
			res = "belowAverage";
		} else if (Asc.c_oAscDynamicAutoFilter.tomorrow === val) {
			res = "tomorrow";
		} else if (Asc.c_oAscDynamicAutoFilter.today === val) {
			res = "today";
		} else if (Asc.c_oAscDynamicAutoFilter.yesterday === val) {
			res = "yesterday";
		} else if (Asc.c_oAscDynamicAutoFilter.nextWeek === val) {
			res = "nextWeek";
		} else if (Asc.c_oAscDynamicAutoFilter.thisWeek === val) {
			res = "thisWeek";
		} else if (Asc.c_oAscDynamicAutoFilter.lastWeek === val) {
			res = "lastWeek";
		} else if (Asc.c_oAscDynamicAutoFilter.nextMonth === val) {
			res = "nextMonth";
		} else if (Asc.c_oAscDynamicAutoFilter.thisMonth === val) {
			res = "thisMonth";
		} else if (Asc.c_oAscDynamicAutoFilter.lastMonth === val) {
			res = "lastMonth";
		} else if (Asc.c_oAscDynamicAutoFilter.nextQuarter === val) {
			res = "nextQuarter";
		} else if (Asc.c_oAscDynamicAutoFilter.thisQuarter === val) {
			res = "thisQuarter";
		} else if (Asc.c_oAscDynamicAutoFilter.lastQuarter === val) {
			res = "lastQuarter";
		} else if (Asc.c_oAscDynamicAutoFilter.nextYear === val) {
			res = "nextYear";
		} else if (Asc.c_oAscDynamicAutoFilter.thisYear === val) {
			res = "thisYear";
		} else if (Asc.c_oAscDynamicAutoFilter.lastYear === val) {
			res = "lastYear";
		} else if (Asc.c_oAscDynamicAutoFilter.yearToDate === val) {
			res = "yearToDate";
		} else if (Asc.c_oAscDynamicAutoFilter.q1 === val) {
			res = "Q1";
		} else if (Asc.c_oAscDynamicAutoFilter.q2 === val) {
			res = "Q2";
		} else if (Asc.c_oAscDynamicAutoFilter.q3 === val) {
			res = "Q3";
		} else if (Asc.c_oAscDynamicAutoFilter.q4 === val) {
			res = "Q4";
		} else if (Asc.c_oAscDynamicAutoFilter.m1 === val) {
			res = "M1";
		} else if (Asc.c_oAscDynamicAutoFilter.m2 === val) {
			res = "M2";
		} else if (Asc.c_oAscDynamicAutoFilter.m3 === val) {
			res = "M3";
		} else if (Asc.c_oAscDynamicAutoFilter.m4 === val) {
			res = "M4";
		} else if (Asc.c_oAscDynamicAutoFilter.m5 === val) {
			res = "M5";
		} else if (Asc.c_oAscDynamicAutoFilter.m6 === val) {
			res = "M6";
		} else if (Asc.c_oAscDynamicAutoFilter.m7 === val) {
			res = "M7";
		} else if (Asc.c_oAscDynamicAutoFilter.m8 === val) {
			res = "M8";
		} else if (Asc.c_oAscDynamicAutoFilter.m9 === val) {
			res = "M9";
		} else if (Asc.c_oAscDynamicAutoFilter.m10 === val) {
			res = "M10";
		} else if (Asc.c_oAscDynamicAutoFilter.m11 === val) {
			res = "M11";
		} else if (Asc.c_oAscDynamicAutoFilter.m12 === val) {
			res = "M12";
		}
		return res;
	}

	function FromXml_ST_DateTimeGrouping(val) {
		var res = -1;
		if ("year" === val) {
			res = Asc.EDateTimeGroup.datetimegroupYear;
		} else if ("month" === val) {
			res = Asc.EDateTimeGroup.datetimegroupMonth;
		} else if ("day" === val) {
			res = Asc.EDateTimeGroup.datetimegroupDay;
		} else if ("hour" === val) {
			res = Asc.EDateTimeGroup.datetimegroupHour;
		} else if ("minute" === val) {
			res = Asc.EDateTimeGroup.datetimegroupMinute;
		} else if ("second" === val) {
			res = Asc.EDateTimeGroup.datetimegroupSecond;
		}
		return res;
	}
	function ToXml_ST_DateTimeGrouping(val) {
		var res = "";
		if (Asc.EDateTimeGroup.datetimegroupYear === val) {
			res = "year";
		} else if (Asc.EDateTimeGroup.datetimegroupMonth === val) {
			res = "month";
		} else if (Asc.EDateTimeGroup.datetimegroupDay === val) {
			res = "day";
		} else if (Asc.EDateTimeGroup.datetimegroupHour === val) {
			res = "hour";
		} else if (Asc.EDateTimeGroup.datetimegroupMinute === val) {
			res = "minute";
		} else if (Asc.EDateTimeGroup.datetimegroupSecond === val) {
			res = "second";
		}
		return res;
	}
	
	function ToXml_ST_HorizontalAlignment(val) {
		var res = "";
		switch (val)
		{
			case -1:
				res = "general";
				break;
			case AscCommon.align_Left:
				res = "left";
				break;
			case AscCommon.align_Center:
				res = "center";
				break;
			case AscCommon.align_Right:
				res = "right";
				break;
			case AscCommon.align_Justify:
				res = "justify";
				break;
			case AscCommon.align_CenterContinuous:
				res = "centerContinuous";
				break;
		}
		return res;
	}
	function FromXml_ST_HorizontalAlignment(val) {
		var res = -1;
		if ("general" === val) {
			res = -1;
		} else if ("left" === val) {
			res = AscCommon.align_Left;
		} else if ("center" === val) {
			res = AscCommon.align_Center;
		} else if ("right" === val) {
			res = AscCommon.align_Right;
		} else if ("fill" === val) {
			res = AscCommon.align_Justify;
		} else if ("justify" === val) {
			res = AscCommon.align_Justify;
		} else if ("centerContinuous" === val) {
			res = AscCommon.align_Center;
		} else if ("distributed" === val) {
			res = AscCommon.align_Justify;
		}
		return res;
	}

	function ToXml_ST_VerticalAlignment(val) {
		var res = "";
		switch (val)
		{
			case Asc.c_oAscVAlign.Top:
				res = "top";
				break;
			case Asc.c_oAscVAlign.Center:
				res = "center";
				break;
			case Asc.c_oAscVAlign.Bottom:
				res = "bottom";
				break;
			case Asc.c_oAscVAlign.Just:
				res = "justify";
				break;
			case Asc.c_oAscVAlign.Dist:
				res = "distributed";
				break;
		}
		return res;
	}
	function FromXml_ST_VerticalAlignment(val) {
		var res = -1;
		if ("top" === val) {
			res = Asc.c_oAscVAlign.Top;
		} else if ("center" === val) {
			res = Asc.c_oAscVAlign.Center;
		} else if ("bottom" === val) {
			res = Asc.c_oAscVAlign.Bottom;
		} else if ("justify" === val) {
			res = Asc.c_oAscVAlign.Just;
		} else if ("distributed" === val) {
			res = Asc.c_oAscVAlign.Dist;
		}
		return res;
	}

	function ToXML_ST_CfvoType(nType)
	{
		var sType = "";
		switch (nType)
		{
			case AscCommonExcel.ECfvoType.Formula:
				sType = "formula";
				break;
			case AscCommonExcel.ECfvoType.Maximum:
				sType = "max";
				break;
			case AscCommonExcel.ECfvoType.Minimum:
				sType = "min";
				break;
			case AscCommonExcel.ECfvoType.Number:
				sType = "num";
				break;
			case AscCommonExcel.ECfvoType.Percent:
				sType = "percent";
				break;
			case AscCommonExcel.ECfvoType.Percentile:
				sType = "percentile";
				break;
			case AscCommonExcel.ECfvoType.AutoMin:
				sType = "autoMin";
				break;
			case AscCommonExcel.ECfvoType.AutoMax:
				sType = "autoMax";
				break;
		}

		return sType;
	}
	function FromXML_ST_CfvoType(sType)
	{
		var nType = -1;
		switch (sType)
		{
			case "formula":
				nType = AscCommonExcel.ECfvoType.Formula;
				break;
			case "max":
				nType = AscCommonExcel.ECfvoType.Maximum;
				break;
			case "min":
				nType = AscCommonExcel.ECfvoType.Minimum;
				break;
			case "num":
				nType = AscCommonExcel.ECfvoType.Number;
				break;
			case "percent":
				nType = AscCommonExcel.ECfvoType.Percent;
				break;
			case "percentile":
				nType = AscCommonExcel.ECfvoType.Percentile;
				break;
			case "autoMin":
				nType = AscCommonExcel.ECfvoType.AutoMin;
				break;
			case "autoMax":
				nType = AscCommonExcel.ECfvoType.AutoMax;
				break;
		}

		return nType;
	}

	function ToXML_IconSetType(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.EIconSetType.Arrows3:
				sType = "3Arrows";
				break;
			case Asc.EIconSetType.Arrows3Gray:
				sType = "3ArrowsGray";
				break;
			case Asc.EIconSetType.Flags3:
				sType = "3Flags";
				break;
			case Asc.EIconSetType.Signs3:
				sType = "3Signs";
				break;
			case Asc.EIconSetType.Symbols3:
				sType = "3Symbols";
				break;
			case Asc.EIconSetType.Symbols3_2:
				sType = "3Symbols2";
				break;
			case Asc.EIconSetType.Traffic3Lights1:
				sType = "3TrafficLights1";
				break;
			case Asc.EIconSetType.Traffic3Lights2:
				sType = "3TrafficLights2";
				break;
			case Asc.EIconSetType.Arrows4:
				sType = "4Arrows";
				break;
			case Asc.EIconSetType.Arrows4Gray:
				sType = "4ArrowsGray";
				break;
			case Asc.EIconSetType.Rating4:
				sType = "4Rating";
				break;
			case Asc.EIconSetType.RedToBlack4:
				sType = "4RedToBlack";
				break;
			case Asc.EIconSetType.Traffic4Lights:
				sType = "4TrafficLights";
				break;
			case Asc.EIconSetType.Arrows5:
				sType = "5Arrows";
				break;
			case Asc.EIconSetType.Arrows5Gray:
				sType = "5ArrowsGray";
				break;
			case Asc.EIconSetType.Quarters5:
				sType = "5Quarters";
				break;
			case Asc.EIconSetType.Rating5: 
				sType = "5Rating";
				break;
			case Asc.EIconSetType.Triangles3:
				sType = "3Triangles";
				break;
			case Asc.EIconSetType.Stars3:
				sType = "3Stars";
				break;
			case Asc.EIconSetType.Boxes5:
				sType = "5Boxes";
				break;
			case Asc.EIconSetType.NoIcons:
				sType = "NoIcons";
				break;
		}

		return sType;
	}
	function FromXML_IconSetType(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "3Arrows":
				nType = Asc.EIconSetType.Arrows3;
				break;
			case "3ArrowsGray":
				nType = Asc.EIconSetType.Arrows3Gray;
				break;
			case "3Flags":
				nType = Asc.EIconSetType.Flags3;
				break;
			case "3Signs":
				nType = Asc.EIconSetType.Signs3;
				break;
			case "3Symbols":
				nType = Asc.EIconSetType.Symbols3;
				break;
			case "3Symbols2":
				nType = Asc.EIconSetType.Symbols3_2;
				break;
			case "3TrafficLights1":
				nType = Asc.EIconSetType.Traffic3Lights1;
				break;
			case "3TrafficLights2":
				nType = Asc.EIconSetType.Traffic3Lights2;
				break;
			case "4Arrows":
				nType = Asc.EIconSetType.Arrows4;
				break;
			case "4ArrowsGray":
				nType = Asc.EIconSetType.Arrows4Gray;
				break;
			case "4Rating":
				nType = Asc.EIconSetType.Rating4;
				break;
			case "4RedToBlack":
				nType = Asc.EIconSetType.RedToBlack4;
				break;
			case "4TrafficLights":
				nType = Asc.EIconSetType.Traffic4Lights;
				break;
			case "5Arrows":
				nType = Asc.EIconSetType.Arrows5;
				break;
			case "5ArrowsGray":
				nType = Asc.EIconSetType.Arrows5Gray;
				break;
			case "5Quarters":
				nType = Asc.EIconSetType.Quarters5;
				break;
			case "5Rating": 
				nType = Asc.EIconSetType.Rating5;
				break;
			case "3Triangles":
				nType = Asc.EIconSetType.Triangles3;
				break;
			case "3Stars":
				nType = Asc.EIconSetType.Stars3;
				break;
			case "5Boxes":
				nType = Asc.EIconSetType.Boxes5;
				break;
			case "NoIcons":
				nType = Asc.EIconSetType.NoIcons;
				break;
		}

		return nType;
	}

	function ToXML_CFOperatorType(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommonExcel.ECfOperator.Operator_beginsWith:
				sType = "beginsWith";
				break;
			case AscCommonExcel.ECfOperator.Operator_between:
				sType = "between";
				break;
			case AscCommonExcel.ECfOperator.Operator_containsText:
				sType = "containsText";
				break;
			case AscCommonExcel.ECfOperator.Operator_endsWith:
				sType = "endsWith";
				break;
			case AscCommonExcel.ECfOperator.Operator_equal:
				sType = "equal";
				break;
			case AscCommonExcel.ECfOperator.Operator_greaterThan:
				sType = "greaterThan";
				break;
			case AscCommonExcel.ECfOperator.Operator_greaterThanOrEqual:
				sType = "greaterThanOrEqual";
				break;
			case AscCommonExcel.ECfOperator.Operator_lessThan:
				sType = "lessThan";
				break;
			case AscCommonExcel.ECfOperator.Operator_lessThanOrEqual:
				sType = "lessThanOrEqual";
				break;
			case AscCommonExcel.ECfOperator.Operator_notBetween:
				sType = "notBetween";
				break;
			case AscCommonExcel.ECfOperator.Operator_notContains:
				sType = "notContains";
				break;
			case AscCommonExcel.ECfOperator.Operator_notEqual:
				sType = "notEqual";
				break;
		}

		return sType;
	}
	function FromXML_CFOperatorType(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "beginsWith":
				nType = AscCommonExcel.ECfOperator.Operator_beginsWith;
				break;
			case "between":
				nType = AscCommonExcel.ECfOperator.Operator_between;
				break;
			case "containsText":
				nType = AscCommonExcel.ECfOperator.Operator_containsText;
				break;
			case "endsWith":
				nType = AscCommonExcel.ECfOperator.Operator_endsWith;
				break;
			case "equal":
				nType = AscCommonExcel.ECfOperator.Operator_equal;
				break;
			case "greaterThan":
				nType = AscCommonExcel.ECfOperator.Operator_greaterThan;
				break;
			case "greaterThanOrEqual":
				nType = AscCommonExcel.ECfOperator.Operator_greaterThanOrEqual;
				break;
			case "lessThan":
				nType = AscCommonExcel.ECfOperator.Operator_lessThan;
				break;
			case "lessThanOrEqual":
				nType = AscCommonExcel.ECfOperator.Operator_lessThanOrEqual;
				break;
			case "notBetween":
				nType = AscCommonExcel.ECfOperator.Operator_notBetween;
				break;
			case "notContains":
				nType = AscCommonExcel.ECfOperator.Operator_notContains;
				break;
			case "notEqual":
				nType = AscCommonExcel.ECfOperator.Operator_notEqual;
				break;
		}

		return nType;
	}

	function ToXML_ST_TimePeriod(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommonExcel.ST_TimePeriod.last7Days:
				sType = "last7Days";
				break;
			case AscCommonExcel.ST_TimePeriod.lastMonth:
				sType = "lastMonth";
				break;
			case AscCommonExcel.ST_TimePeriod.lastWeek:
				sType = "lastWeek";
				break;
			case AscCommonExcel.ST_TimePeriod.nextMonth:
				sType = "nextMonth";
				break;
			case AscCommonExcel.ST_TimePeriod.nextWeek:
				sType = "nextWeek";
				break;
			case AscCommonExcel.ST_TimePeriod.thisMonth:
				sType = "thisMonth";
				break;
			case AscCommonExcel.ST_TimePeriod.thisWeek:
				sType = "thisWeek";
				break;
			case AscCommonExcel.ST_TimePeriod.today:
				sType = "today";
				break;
			case AscCommonExcel.ST_TimePeriod.tomorrow:
				sType = "tomorrow";
				break;
			case AscCommonExcel.ST_TimePeriod.yesterday:
				sType = "yesterday";
				break;
		}

		return sType;
	}
	function FromXML_ST_TimePeriod(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "last7Days":
				nType = AscCommonExcel.ST_TimePeriod.last7Days;
				break;
			case "lastMonth":
				nType = AscCommonExcel.ST_TimePeriod.lastMonth;
				break;
			case "lastWeek":
				nType = AscCommonExcel.ST_TimePeriod.lastWeek;
				break;
			case "nextMonth":
				nType = AscCommonExcel.ST_TimePeriod.nextMonth;
				break;
			case "nextWeek":
				nType = AscCommonExcel.ST_TimePeriod.nextWeek;
				break;
			case "thisMonth":
				nType = AscCommonExcel.ST_TimePeriod.thisMonth;
				break;
			case "thisWeek":
				nType = AscCommonExcel.ST_TimePeriod.thisWeek;
				break;
			case "today":
				nType = AscCommonExcel.ST_TimePeriod.today;
				break;
			case "tomorrow":
				nType = AscCommonExcel.ST_TimePeriod.tomorrow;
				break;
			case "yesterday":
				nType = AscCommonExcel.ST_TimePeriod.yesterday;
				break;
		}

		return nType;
	}

	function ToXML_CfRuleType(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.ECfType.aboveAverage:
				sType = "aboveAverage";
				break;
			case Asc.ECfType.beginsWith:
				sType = "beginsWith";
				break;
			case Asc.ECfType.cellIs:
				sType = "cellIs";
				break;
			case Asc.ECfType.colorScale:
				sType = "colorScale";
				break;
			case Asc.ECfType.containsBlanks:
				sType = "containsBlanks";
				break;
			case Asc.ECfType.containsErrors:
				sType = "containsErrors";
				break;
			case Asc.ECfType.containsText:
				sType = "containsText";
				break;
			case Asc.ECfType.dataBar:
				sType = "dataBar";
				break;
			case Asc.ECfType.duplicateValues:
				sType = "duplicateValues";
				break;
			case Asc.ECfType.expression:
				sType = "expression";
				break;
			case Asc.ECfType.iconSet:
				sType = "iconSet";
				break;
			case Asc.ECfType.notContainsBlanks:
				sType = "notContainsBlanks";
				break;
			case Asc.ECfType.notContainsErrors:
				sType = "notContainsErrors";
				break;
			case Asc.ECfType.notContainsText:
				sType = "notContainsText";
				break;
			case Asc.ECfType.timePeriod:
				sType = "timePeriod";
				break;
			case Asc.ECfType.top10:
				sType = "top10";
				break;
			case Asc.ECfType.uniqueValues:
				sType = "uniqueValues";
				break;
			case Asc.ECfType.endsWith:
				sType = "endsWith";
				break;
		}

		return sType;
	}
	function FromXML_CfRuleType(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "aboveAverage":
				nType = Asc.ECfType.aboveAverage;
				break;
			case "beginsWith":
				nType = Asc.ECfType.beginsWith;
				break;
			case "cellIs":
				nType = Asc.ECfType.cellIs;
				break;
			case "colorScale":
				nType = Asc.ECfType.colorScale;
				break;
			case "containsBlanks":
				nType = Asc.ECfType.containsBlanks;
				break;
			case "containsErrors":
				nType = Asc.ECfType.containsErrors;
				break;
			case "containsText":
				nType = Asc.ECfType.containsText;
				break;
			case "dataBar":
				nType = Asc.ECfType.dataBar;
				break;
			case "duplicateValues":
				nType = Asc.ECfType.duplicateValues;
				break;
			case "expression":
				nType = Asc.ECfType.expression;
				break;
			case "iconSet":
				nType = Asc.ECfType.iconSet;
				break;
			case "notContainsBlanks":
				nType = Asc.ECfType.notContainsBlanks;
				break;
			case "notContainsErrors":
				nType = Asc.ECfType.notContainsErrors;
				break;
			case "notContainsText":
				nType = Asc.ECfType.notContainsText;
				break;
			case "timePeriod":
				nType = Asc.ECfType.timePeriod;
				break;
			case "top10":
				nType = Asc.ECfType.top10;
				break;
			case "uniqueValues":
				nType = Asc.ECfType.uniqueValues;
				break;
			case "endsWith":
				nType = Asc.ECfType.endsWith;
				break;
		}

		return nType;
	}

	function ToXML_ST_DataValidationErrorStyle(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.EDataValidationErrorStyle.Stop:
				sType = "stop";
				break;
			case Asc.EDataValidationErrorStyle.Warning:
				sType = "warning";
				break;
			case Asc.EDataValidationErrorStyle.Information:
				sType = "information";
				break;
		}

		return sType;
	}
	function FromXML_ST_DataValidationErrorStyle(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "stop":
				nType = Asc.EDataValidationErrorStyle.Stop;
				break;
			case "warning":
				nType = Asc.EDataValidationErrorStyle.Warning;
				break;
			case "information":
				nType = Asc.EDataValidationErrorStyle.Information;
				break;
		}

		return nType;
	}

	function ToXML_ST_DataValidationImeMode(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.EDataValidationImeMode.NoControl:
				sType = "noControl";
				break;
			case Asc.EDataValidationImeMode.Off:
				sType = "off";
				break;
			case Asc.EDataValidationImeMode.On:
				sType = "on";
				break;
			case Asc.EDataValidationImeMode.Disabled:
				sType = "disabled";
				break;
			case Asc.EDataValidationImeMode.Hiragana:
				sType = "hiragana";
				break;
			case Asc.EDataValidationImeMode.FullKatakana:
				sType = "fullKatakana";
				break;
			case Asc.EDataValidationImeMode.HalfKatakana:
				sType = "halfKatakana";
				break;
			case Asc.EDataValidationImeMode.FullAlpha:
				sType = "fullAlpha";
				break;
			case Asc.EDataValidationImeMode.HalfAlpha:
				sType = "halfAlpha";
				break;
			case Asc.EDataValidationImeMode.FullHangul:
				sType = "fullHangul";
				break;
			case Asc.EDataValidationImeMode.HalfHangul:
				sType = "halfHangul";
				break;
		}

		return sType;
	}
	function FromXML_ST_DataValidationImeMode(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "noControl":
				nType = Asc.EDataValidationImeMode.NoControl;
				break;
			case "off":
				nType = Asc.EDataValidationImeMode.Off;
				break;
			case "on":
				nType = Asc.EDataValidationImeMode.On;
				break;
			case "disabled":
				nType = Asc.EDataValidationImeMode.Disabled;
				break;
			case "hiragana":
				nType = Asc.EDataValidationImeMode.Hiragana;
				break;
			case "fullKatakana":
				nType = Asc.EDataValidationImeMode.FullKatakana;
				break;
			case "halfKatakana":
				nType = Asc.EDataValidationImeMode.HalfKatakana;
				break;
			case "fullAlpha":
				nType = Asc.EDataValidationImeMode.FullAlpha;
				break;
			case "halfAlpha":
				nType = Asc.EDataValidationImeMode.HalfAlpha;
				break;
			case "fullHangul":
				nType = Asc.EDataValidationImeMode.FullHangul;
				break;
			case "halfHangul":
				nType = Asc.EDataValidationImeMode.HalfHangul;
				break;
		}

		return nType;
	}
	
	function ToXML_ST_DataValidationOperator(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.EDataValidationOperator.Between:
				sType = "between";
				break;
			case Asc.EDataValidationOperator.NotBetween:
				sType = "notBetween";
				break;
			case Asc.EDataValidationOperator.Equal:
				sType = "equal";
				break;
			case Asc.EDataValidationOperator.NotEqual:
				sType = "notEqual";
				break;
			case Asc.EDataValidationOperator.LessThan:
				sType = "lessThan";
				break;
			case Asc.EDataValidationOperator.LessThanOrEqual:
				sType = "lessThanOrEqual";
				break;
			case Asc.EDataValidationOperator.GreaterThan:
				sType = "greaterThan";
				break;
			case Asc.EDataValidationOperator.GreaterThanOrEqual:
				sType = "greaterThanOrEqual";
				break;
		}

		return sType;
	}
	function FromXML_ST_DataValidationOperator(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "between":
				nType = Asc.EDataValidationOperator.Between;
				break;
			case "notBetween":
				nType = Asc.EDataValidationOperator.NotBetween;
				break;
			case "equal":
				nType = Asc.EDataValidationOperator.Equal;
				break;
			case "notEqual":
				nType = Asc.EDataValidationOperator.NotEqual;
				break;
			case "lessThan":
				nType = Asc.EDataValidationOperator.LessThan;
				break;
			case "lessThanOrEqual":
				nType = Asc.EDataValidationOperator.LessThanOrEqual;
				break;
			case "greaterThan":
				nType = Asc.EDataValidationOperator.GreaterThan;
				break;
			case "greaterThanOrEqual":
				nType = Asc.EDataValidationOperator.GreaterThanOrEqual;
				break;
		}

		return nType;
	}

	function ToXML_ST_DataValidationType(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.EDataValidationType.None:
				sType = "none";
				break;
			case Asc.EDataValidationType.Custom:
				sType = "custom";
				break;
			case Asc.EDataValidationType.Date:
				sType = "date";
				break;
			case Asc.EDataValidationType.Decimal:
				sType = "decimal";
				break;
			case Asc.EDataValidationType.List:
				sType = "list";
				break;
			case Asc.EDataValidationType.TextLength:
				sType = "textLength";
				break;
			case Asc.EDataValidationType.Time:
				sType = "time";
				break;
			case Asc.EDataValidationType.Whole:
				sType = "whole";
				break;
		}

		return sType;
	}
	function FromXML_ST_DataValidationType(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "none":
				nType = Asc.EDataValidationType.None;
				break;
			case "custom":
				nType = Asc.EDataValidationType.Custom;
				break;
			case "date":
				nType = Asc.EDataValidationType.Date;
				break;
			case "decimal":
				nType = Asc.EDataValidationType.Decimal;
				break;
			case "list":
				nType = Asc.EDataValidationType.List;
				break;
			case "textLength":
				nType = Asc.EDataValidationType.TextLength;
				break;
			case "time":
				nType = Asc.EDataValidationType.Time;
				break;
			case "whole":
				nType = Asc.EDataValidationType.Whole;
				break;
		}

		return nType;
	}

	function ToXML_ST_EditAs(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommon.c_oAscCellAnchorType.cellanchorAbsolute:
				sType = "absolute";
				break;
			case AscCommon.c_oAscCellAnchorType.cellanchorOneCell:
				sType = "oneCell";
				break;
			case AscCommon.c_oAscCellAnchorType.cellanchorTwoCell:
				sType = "twoCell";
				break;
		}

		return sType;
	}
	function FromXML_ST_EditAs(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "absolute":
				nType = AscCommon.c_oAscCellAnchorType.cellanchorAbsolute;
				break;
			case "oneCell":
				nType = AscCommon.c_oAscCellAnchorType.cellanchorOneCell;
				break;
			case "twoCell":
				nType = AscCommon.c_oAscCellAnchorType.cellanchorTwoCell;
				break;
		}

		return nType;
	}

	function ToXML_ST_CellComments(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommonExcel.ST_CellComments.asDisplayed:
				sType = "asDisplayed";
				break;
			case AscCommonExcel.ST_CellComments.atEnd:
				sType = "atEnd";
				break;
			case AscCommonExcel.ST_CellComments.none:
				sType = "none";
				break;
		}

		return sType;
	}
	function FromXML_ST_CellComments(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "asDisplayed":
				nType = AscCommonExcel.ST_CellComments.asDisplayed;
				break;
			case "atEnd":
				nType = AscCommonExcel.ST_CellComments.atEnd;
				break;
			case "none":
				nType = AscCommonExcel.ST_CellComments.none;
				break;
		}

		return nType;
	}

	function ToXML_ST_PrintError(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommonExcel.ST_PrintError.blank:
				sType = "blank";
				break;
			case AscCommonExcel.ST_PrintError.dash:
				sType = "dash";
				break;
			case AscCommonExcel.ST_PrintError.displayed:
				sType = "displayed";
				break;
			case AscCommonExcel.ST_PrintError.NA:
				sType = "NA";
				break;
		}

		return sType;
	}
	function FromXML_ST_PrintError(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "blank":
				nType =  AscCommonExcel.ST_PrintError.blank;
				break;
			case "dash":
				nType =  AscCommonExcel.ST_PrintError.dash;
				break;
			case "displayed":
				nType =  AscCommonExcel.ST_PrintError.displayed;
				break;
			case "NA":
				nType =  AscCommonExcel.ST_PrintError.NA;
				break;
		}

		return nType;
	}

	function ToXML_ST_PageOrder(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommonExcel.ST_PageOrder.downThenOver:
				sType = "downThenOver";
				break;
			case AscCommonExcel.ST_PageOrder.overThenDown:
				sType = "overThenDown";
				break;
		}

		return sType;
	}
	function FromXML_ST_PageOrder(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "downThenOver":
				nType = AscCommonExcel.ST_PageOrder.downThenOver;
				break;
			case "overThenDown":
				nType = AscCommonExcel.ST_PageOrder.overThenDown;
				break;
		}

		return nType;
	}

	function To_XML_CryptoAlgorithmName(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case AscCommon.CryptoAlgorithmName.Unknown:
				sVal = "Unknown";
				break;
			case AscCommon.CryptoAlgorithmName.MD2:
				sVal = "MD2";
				break;
			case AscCommon.CryptoAlgorithmName.MD4:
				sVal = "MD4";
				break;
			case AscCommon.CryptoAlgorithmName.MD5:
				sVal = "MD5";
				break;
			case AscCommon.CryptoAlgorithmName.RIPEMD128:
				sVal = "RIPEMD128";
				break;
			case AscCommon.CryptoAlgorithmName.RIPEMD160:
				sVal = "RIPEMD160";
				break;
			case AscCommon.CryptoAlgorithmName.SHA1:
				sVal = "SHA1";
				break;
			case AscCommon.CryptoAlgorithmName.SHA256:
				sVal = "SHA256";
				break;
			case AscCommon.CryptoAlgorithmName.SHA384:
				sVal = "SHA384";
				break;
			case AscCommon.CryptoAlgorithmName.SHA512:
				sVal = "SHA512";
				break;
			case AscCommon.CryptoAlgorithmName.WHIRLPOOL:
				sVal = "WHIRLPOOL";
				break;
		}

		return sVal;
	}
	function From_XML_CryptoAlgorithmName(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "Unknown":
				nVal = AscCommon.CryptoAlgorithmName.Unknown;
				break;
			case "MD2":
				nVal = AscCommon.CryptoAlgorithmName.MD2;
				break;
			case "MD4":
				nVal = AscCommon.CryptoAlgorithmName.MD4;
				break;
			case "MD5":
				nVal = AscCommon.CryptoAlgorithmName.MD5;
				break;
			case "RIPEMD128":
				nVal = AscCommon.CryptoAlgorithmName.RIPEMD128;
				break;
			case "RIPEMD160":
				nVal = AscCommon.CryptoAlgorithmName.RIPEMD160;
				break;
			case "SHA1":
				nVal = AscCommon.CryptoAlgorithmName.SHA1;
				break;
			case "SHA256":
				nVal = AscCommon.CryptoAlgorithmName.SHA256;
				break;
			case "SHA384":
				nVal = AscCommon.CryptoAlgorithmName.SHA384;
				break;
			case "SHA512":
				nVal = AscCommon.CryptoAlgorithmName.SHA512;
				break;
			case "WHIRLPOOL":
				nVal = AscCommon.CryptoAlgorithmName.WHIRLPOOL;
				break;
		}

		return nVal;
	}

	function ToXML_EActivePane(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommonExcel.EActivePane.bottomLeft:
				sType = "bottomLeft";
				break;
			case AscCommonExcel.EActivePane.bottomRight:
				sType = "bottomRight";
				break;
			case AscCommonExcel.EActivePane.topLeft:
				sType = "topLeft";
				break;
			case AscCommonExcel.EActivePane.topRight:
				sType = "topRight";
				break;
		}

		return sType;
	}
	function FromXML_EActivePane(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "bottomLeft":
				nType = AscCommonExcel.EActivePane.bottomLeft;
				break;
			case "bottomRight":
				nType = AscCommonExcel.EActivePane.bottomRight;
				break;
			case "topLeft":
				nType = AscCommonExcel.EActivePane.topLeft;
				break;
			case "topRight":
				nType = AscCommonExcel.EActivePane.topRight;
				break;
		}

		return nType;
	}

	function ToXML_ST_SortMethod(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommonExcel.ESortMethod.sortmethodNone:
				sType = "none";
				break;
			case AscCommonExcel.ESortMethod.sortmethodPinYin:
				sType = "pinYin";
				break;
			case AscCommonExcel.ESortMethod.sortmethodStroke:
				sType = "stroke";
				break;
		}

		return sType;
	}
	function FromXML_ST_SortMethod(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "none":
				nType = AscCommonExcel.ESortMethod.sortmethodNone;
				break;
			case "pinYin":
				nType = AscCommonExcel.ESortMethod.sortmethodPinYin;
				break;
			case "stroke":
				nType = AscCommonExcel.ESortMethod.sortmethodStroke;
				break;
		}

		return nType;
	}

	function ToXML_ETotalsRowFunction(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommonExcel.ETotalsRowFunction.totalrowfunctionAverage:
				sType = "average";
				break;
			case AscCommonExcel.ETotalsRowFunction.totalrowfunctionCount:
				sType = "count";
				break;
			case AscCommonExcel.ETotalsRowFunction.totalrowfunctionCountNums:
				sType = "countNums";
				break;
			case AscCommonExcel.ETotalsRowFunction.totalrowfunctionCustom:
				sType = "custom";
				break;
			case AscCommonExcel.ETotalsRowFunction.totalrowfunctionMax:
				sType = "max";
				break;
			case AscCommonExcel.ETotalsRowFunction.totalrowfunctionMin:
				sType = "min";
				break;
			case AscCommonExcel.ETotalsRowFunction.totalrowfunctionNone:
				sType = "none";
				break;
			case AscCommonExcel.ETotalsRowFunction.totalrowfunctionStdDev:
				sType = "stdDev";
				break;
			case AscCommonExcel.ETotalsRowFunction.totalrowfunctionSum:
				sType = "sum";
				break;
			case AscCommonExcel.ETotalsRowFunction.totalrowfunctionVar:
				sType = "var";
				break;
		}

		return sType;
	}
	function FromXML_ETotalsRowFunction(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "average":
				nType = AscCommonExcel.ETotalsRowFunction.totalrowfunctionAverage;
				break;
			case "count":
				nType = AscCommonExcel.ETotalsRowFunction.totalrowfunctionCount;
				break;
			case "countNums":
				nType = AscCommonExcel.ETotalsRowFunction.totalrowfunctionCountNums;
				break;
			case "custom":
				nType = AscCommonExcel.ETotalsRowFunction.totalrowfunctionCustom;
				break;
			case "max":
				nType = AscCommonExcel.ETotalsRowFunction.totalrowfunctionMax;
				break;
			case "min":
				nType = AscCommonExcel.ETotalsRowFunction.totalrowfunctionMin;
				break;
			case "none":
				nType = AscCommonExcel.ETotalsRowFunction.totalrowfunctionNone;
				break;
			case "stdDev":
				nType = AscCommonExcel.ETotalsRowFunction.totalrowfunctionStdDev;
				break;
			case "sum":
				nType = AscCommonExcel.ETotalsRowFunction.totalrowfunctionSum;
				break;
			case "var":
				nType = AscCommonExcel.ETotalsRowFunction.totalrowfunctionVar;
				break;
		}

		return nType;
	}

	function ToXML_ST_TableType(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommonExcel.ST_TableType.queryTable:
				sType = "queryTable";
				break;
			case AscCommonExcel.ST_TableType.worksheet:
				sType = "worksheet";
				break;
			case AscCommonExcel.ST_TableType.xml:
				sType = "xml";
				break;
		}

		return sType;
	}
	function FromXML_ST_TableType(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "queryTable":
				nType = AscCommonExcel.ST_TableType.queryTable;
				break;
			case "worksheet":
				nType = AscCommonExcel.ST_TableType.worksheet;
				break;
			case "xml":
				nType = AscCommonExcel.ST_TableType.xml;
				break;
		}

		return nType;
	}

	function To_XML_EBorderStyle(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case Asc.EBorderStyle.borderstyleDashDot:
				sVal = "dashDot";
				break;
			case Asc.EBorderStyle.borderstyleDashDotDot:
				sVal = "dashDotDot";
				break;
			case Asc.EBorderStyle.borderstyleDashed:
				sVal = "dashed";
				break;
			case Asc.EBorderStyle.borderstyleDotted:
				sVal = "dotted";
				break;
			case Asc.EBorderStyle.borderstyleDouble:
				sVal = "double";
				break;
			case Asc.EBorderStyle.borderstyleHair:
				sVal = "hair";
				break;
			case Asc.EBorderStyle.borderstyleMedium:
				sVal = "medium";
				break;
			case Asc.EBorderStyle.borderstyleMediumDashDot:
				sVal = "mediumDashDot";
				break;
			case Asc.EBorderStyle.borderstyleMediumDashDotDot:
				sVal = "mediumDashDotDot";
				break;
			case Asc.EBorderStyle.borderstyleMediumDashed:
				sVal = "mediumDashed";
				break;
			case Asc.EBorderStyle.borderstyleNone:
				sVal = "none";
				break;
			case Asc.EBorderStyle.borderstyleSlantDashDot:
				sVal = "slantDashDot";
				break;
			case Asc.EBorderStyle.borderstyleThick:
				sVal = "thick";
				break;
			case Asc.EBorderStyle.borderstyleThin:
				sVal = "thin";
				break;
		}

		return sVal;
	}
	function From_XML_EBorderStyle(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "dashDot":
				nVal = Asc.EBorderStyle.borderstyleDashDot;
				break;
			case "dashDotDot":
				nVal = Asc.EBorderStyle.borderstyleDashDotDot;
				break;
			case "dashed":
				nVal = Asc.EBorderStyle.borderstyleDashed;
				break;
			case "dotted":
				nVal = Asc.EBorderStyle.borderstyleDotted;
				break;
			case "double":
				nVal = Asc.EBorderStyle.borderstyleDouble;
				break;
			case "hair":
				nVal = Asc.EBorderStyle.borderstyleHair;
				break;
			case "medium":
				nVal = Asc.EBorderStyle.borderstyleMedium;
				break;
			case "mediumDashDot":
				nVal = Asc.EBorderStyle.borderstyleMediumDashDot;
				break;
			case "mediumDashDotDot":
				nVal = Asc.EBorderStyle.borderstyleMediumDashDotDot;
				break;
			case "mediumDashed":
				nVal = Asc.EBorderStyle.borderstyleMediumDashed;
				break;
			case "none":
				nVal = Asc.EBorderStyle.borderstyleNone;
				break;
			case "slantDashDot":
				nVal = Asc.EBorderStyle.borderstyleSlantDashDot;
				break;
			case "thick":
				nVal = Asc.EBorderStyle.borderstyleThick;
				break;
			case "thin":
				nVal = Asc.EBorderStyle.borderstyleThin;
				break;
		}

		return nVal;
	}

	function ToXml_ST_PatternType(val) {
		var res = -1;
		if (AscCommonExcel.c_oAscPatternType.None === val) {
			res = "none";
		} else if (AscCommonExcel.c_oAscPatternType.Solid === val) {
			res = "solid";
		} else if (AscCommonExcel.c_oAscPatternType.MediumGray === val) {
			res = "mediumGray";
		} else if (AscCommonExcel.c_oAscPatternType.DarkGray === val) {
			res = "darkGray";
		} else if (AscCommonExcel.c_oAscPatternType.LightGray === val) {
			res = "lightGray";
		} else if (AscCommonExcel.c_oAscPatternType.DarkHorizontal === val) {
			res = "darkHorizontal";
		} else if (AscCommonExcel.c_oAscPatternType.DarkVertical === val) {
			res = "darkVertical";
		} else if (AscCommonExcel.c_oAscPatternType.DarkDown === val) {
			res = "darkDown";
		} else if (AscCommonExcel.c_oAscPatternType.DarkUp === val) {
			res = "darkUp";
		} else if (AscCommonExcel.c_oAscPatternType.DarkGrid === val) {
			res = "darkGrid";
		} else if (AscCommonExcel.c_oAscPatternType.DarkTrellis === val) {
			res = "darkTrellis";
		} else if (AscCommonExcel.c_oAscPatternType.LightHorizontal === val) {
			res = "lightHorizontal";
		} else if (AscCommonExcel.c_oAscPatternType.LightVertical === val) {
			res = "lightVertical";
		} else if (AscCommonExcel.c_oAscPatternType.LightDown === val) {
			res = "lightDown";
		} else if (AscCommonExcel.c_oAscPatternType.LightUp === val) {
			res = "lightUp";
		} else if (AscCommonExcel.c_oAscPatternType.LightGrid === val) {
			res = "lightGrid";
		} else if (AscCommonExcel.c_oAscPatternType.LightTrellis === val) {
			res = "lightTrellis";
		} else if (AscCommonExcel.c_oAscPatternType.Gray125 === val) {
			res = "gray125";
		} else if (AscCommonExcel.c_oAscPatternType.Gray0625 === val) {
			res = "gray0625";
		}
		return res;
	}
	function FromXml_ST_PatternType(val) {
		var res = -1;
		if ("none" === val) {
			res = AscCommonExcel.c_oAscPatternType.None;
		} else if ("solid" === val) {
			res = AscCommonExcel.c_oAscPatternType.Solid;
		} else if ("mediumGray" === val) {
			res = AscCommonExcel.c_oAscPatternType.MediumGray;
		} else if ("darkGray" === val) {
			res = AscCommonExcel.c_oAscPatternType.DarkGray;
		} else if ("lightGray" === val) {
			res = AscCommonExcel.c_oAscPatternType.LightGray;
		} else if ("darkHorizontal" === val) {
			res = AscCommonExcel.c_oAscPatternType.DarkHorizontal;
		} else if ("darkVertical" === val) {
			res = AscCommonExcel.c_oAscPatternType.DarkVertical;
		} else if ("darkDown" === val) {
			res = AscCommonExcel.c_oAscPatternType.DarkDown;
		} else if ("darkUp" === val) {
			res = AscCommonExcel.c_oAscPatternType.DarkUp;
		} else if ("darkGrid" === val) {
			res = AscCommonExcel.c_oAscPatternType.DarkGrid;
		} else if ("darkTrellis" === val) {
			res = AscCommonExcel.c_oAscPatternType.DarkTrellis;
		} else if ("lightHorizontal" === val) {
			res = AscCommonExcel.c_oAscPatternType.LightHorizontal;
		} else if ("lightVertical" === val) {
			res = AscCommonExcel.c_oAscPatternType.LightVertical;
		} else if ("lightDown" === val) {
			res = AscCommonExcel.c_oAscPatternType.LightDown;
		} else if ("lightUp" === val) {
			res = AscCommonExcel.c_oAscPatternType.LightUp;
		} else if ("lightGrid" === val) {
			res = AscCommonExcel.c_oAscPatternType.LightGrid;
		} else if ("lightTrellis" === val) {
			res = AscCommonExcel.c_oAscPatternType.LightTrellis;
		} else if ("gray125" === val) {
			res = AscCommonExcel.c_oAscPatternType.Gray125;
		} else if ("gray0625" === val) {
			res = AscCommonExcel.c_oAscPatternType.Gray0625;
		}
		return res;
	}

	function ToXml_ST_GradientType(val) {
		var res = -1;
		if (Asc.c_oAscFillGradType.GRAD_LINEAR === val) {
			res = "linear";
		} else if (Asc.c_oAscFillGradType.GRAD_PATH === val) {
			res = "path";
		}
		return res;
	}
	function FromXml_ST_GradientType(val) {
		var res = -1;
		if ("linear" === val) {
			res = Asc.c_oAscFillGradType.GRAD_LINEAR;
		} else if ("path" === val) {
			res = Asc.c_oAscFillGradType.GRAD_PATH;
		}
		return res;
	}

	function ToXML_EUnderline(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.EUnderline.underlineDouble:
				sType = "double";
				break;
			case Asc.EUnderline.underlineDoubleAccounting:
				sType = "doubleAccounting";
				break;
			case Asc.EUnderline.underlineNone:
				sType = "none";
				break;
			case Asc.EUnderline.underlineSingle:
				sType = "single";
				break;
			case Asc.EUnderline.underlineSingleAccounting:
				sType = "singleAccounting";
				break;
		}

		return sType;
	}
	function FromXML_EUnderline(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "double":
				nType = Asc.EUnderline.underlineDouble;
				break;
			case "doubleAccounting":
				nType = Asc.EUnderline.underlineDoubleAccounting;
				break;
			case "none":
				nType = Asc.EUnderline.underlineNone;
				break;
			case "single":
				nType = Asc.EUnderline.underlineSingle;
				break;
			case "singleAccounting":
				nType = Asc.EUnderline.underlineSingleAccounting;
				break;
		}

		return nType;
	}

	function ToXML_ST_VerticalAlignRun(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommon.vertalign_SubScript:
				sType = "subscript";
				break;
			case AscCommon.vertalign_SuperScript:
				sType = "superscript";
				break;
		}

		return sType;
	}
	function FromXML_ST_VerticalAlignRun(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "subscript":
				nType = AscCommon.vertalign_SubScript;
				break;
			case "superscript":
				nType = AscCommon.vertalign_SuperScript;
				break;
		}

		return nType;
	}

	function ToXML_EFontScheme(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.EFontScheme.fontschemeMajor:
				sType = "major";
				break;
			case Asc.EFontScheme.fontschemeMinor:
				sType = "minor";
				break;
			case Asc.EFontScheme.fontschemeNone:
				sType = "none";
				break;
		}

		return sType;
	}
	function FromXML_EFontScheme(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "major":
				nType = Asc.EFontScheme.fontschemeMajor;
				break;
			case "minor":
				nType = Asc.EFontScheme.fontschemeMinor;
				break;
			case "none":
				nType = Asc.EFontScheme.fontschemeNone;
				break;
		}

		return nType;
	}

	function ToXML_ST_SparklineType(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.c_oAscSparklineType.Line:
				sType = "line";
				break;
			case Asc.c_oAscSparklineType.Column:
				sType = "column";
				break;
			case Asc.c_oAscSparklineType.Stacked:
				sType = "stacked";
				break;
		}

		return sType;
	}
	function FromXML_ST_SparklineType(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "line":
				nType = Asc.c_oAscSparklineType.Line;
				break;
			case "column":
				nType = Asc.c_oAscSparklineType.Column;
				break;
			case "stacked":
				nType = Asc.c_oAscSparklineType.Stacked;
				break;
		}

		return nType;
	}

	function ToXML_ST_DispBlanksAs(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.c_oAscEDispBlanksAs.Span:
				sType = "span";
				break;
			case Asc.c_oAscEDispBlanksAs.Gap:
				sType = "gap";
				break;
			case Asc.c_oAscEDispBlanksAs.Zero:
				sType = "zero";
				break;
		}

		return sType;
	}
	function FromXML_ST_DispBlanksAs(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case "span":
				sType = Asc.c_oAscEDispBlanksAs.Span;
				break;
			case "gap":
				sType = Asc.c_oAscEDispBlanksAs.Gap;
				break;
			case "zero":
				sType = Asc.c_oAscEDispBlanksAs.Zero;
				break;
		}

		return sType;
	}

	function ToXML_ST_SparklineAxisMinMax(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.c_oAscSparklineAxisMinMax.Individual:
				sType = "individual";
				break;
			case Asc.c_oAscSparklineAxisMinMax.Group:
				sType = "group";
				break;
			case Asc.c_oAscSparklineAxisMinMax.Custom:
				sType = "custom";
				break;
		}

		return sType;
	}
	function FromXML_ST_SparklineAxisMinMax(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "individual":
				nType = Asc.c_oAscSparklineAxisMinMax.Individual;
				break;
			case "group":
				nType = Asc.c_oAscSparklineAxisMinMax.Group;
				break;
			case "custom":
				nType = Asc.c_oAscSparklineAxisMinMax.Custom;
				break;
		}

		return nType;
	}

	function ToXml_ST_Axis(val) {
		var res = "";
		if (Asc.c_oAscAxis.AxisRow === val) {
			res = "axisRow";
		} else if (Asc.c_oAscAxis.AxisCol === val) {
			res = "axisCol";
		} else if (Asc.c_oAscAxis.AxisPage === val) {
			res = "axisPage";
		} else if (Asc.c_oAscAxis.AxisValues === val) {
			res = "axisValues";
		}
		return res;
	}
	function FromXml_ST_Axis(val) {
		var res = -1;
		if ("axisRow" === val) {
			res = Asc.c_oAscAxis.AxisRow;
		} else if ("axisCol" === val) {
			res = Asc.c_oAscAxis.AxisCol;
		} else if ("axisPage" === val) {
			res = Asc.c_oAscAxis.AxisPage;
		} else if ("axisValues" === val) {
			res = Asc.c_oAscAxis.AxisValues;
		}
		return res;
	}

	function ToXml_ST_FieldSortType(val) {
		var res = "";
		if (Asc.c_oAscFieldSortType.Manual === val) {
			res = "manual";
		} else if (Asc.c_oAscFieldSortType.Ascending === val) {
			res = "ascending";
		} else if (Asc.c_oAscFieldSortType.Descending === val) {
			res = "descending";
		}
		return res;
	}
	function FromXml_ST_FieldSortType(val) {
		var res = -1;
		if ("manual" === val) {
			res = Asc.c_oAscFieldSortType.Manual;
		} else if ("ascending" === val) {
			res = Asc.c_oAscFieldSortType.Ascending;
		} else if ("descending" === val) {
			res = Asc.c_oAscFieldSortType.Descending;
		}
		return res;
	}

	function ToXml_ST_AllocationMethod(val) {
		var res = "";
		if (Asc.c_oAscAllocationMethod.EqualAllocation === val) {
			res = "equalAllocation";
		} else if (Asc.c_oAscAllocationMethod.EqualIncrement === val) {
			res = "equalIncrement";
		} else if (Asc.c_oAscAllocationMethod.WeightedAllocation === val) {
			res = "weightedAllocation";
		} else if (Asc.c_oAscAllocationMethod.WeightedIncrement === val) {
			res = "weightedIncrement";
		}
		return res;
	}
	function FromXml_ST_AllocationMethod(val) {
		var res = -1;
		if ("equalAllocation" === val) {
			res = Asc.c_oAscAllocationMethod.EqualAllocation;
		} else if ("equalIncrement" === val) {
			res = Asc.c_oAscAllocationMethod.EqualIncrement;
		} else if ("weightedAllocation" === val) {
			res = Asc.c_oAscAllocationMethod.WeightedAllocation;
		} else if ("weightedIncrement" === val) {
			res = Asc.c_oAscAllocationMethod.WeightedIncrement;
		}
		return res;
	}

	function ToXml_ST_PivotAreaType(val) {
		var res = "";
		if (Asc.c_oAscPivotAreaType.None === val) {
			res = "none";
		} else if (Asc.c_oAscPivotAreaType.Data === val) {
			res = "data";
		} else if (Asc.c_oAscPivotAreaType.All === val) {
			res = "all";
		} else if (Asc.c_oAscPivotAreaType.Origin === val) {
			res = "origin";
		} else if (Asc.c_oAscPivotAreaType.Button === val) {
			res = "button";
		} else if (Asc.c_oAscPivotAreaType.TopEnd === val) {
			res = "topEnd";
		}
		return res;
	}
	function FromXml_ST_PivotAreaType(val) {
		// Normal is default.
		var res = Asc.c_oAscPivotAreaType.Normal;
		if ("none" === val) {
			res = Asc.c_oAscPivotAreaType.None;
		} else if ("data" === val) {
			res = Asc.c_oAscPivotAreaType.Data;
		} else if ("all" === val) {
			res = Asc.c_oAscPivotAreaType.All;
		} else if ("origin" === val) {
			res = Asc.c_oAscPivotAreaType.Origin;
		} else if ("button" === val) {
			res = Asc.c_oAscPivotAreaType.Button;
		} else if ("topEnd" === val) {
			res = Asc.c_oAscPivotAreaType.TopEnd;
		}
		return res;
	}

	function ToXml_ST_ItemType(val) {
		var res = "";
		if (Asc.c_oAscItemType.Data === val) {
			res = "data";
		} else if (Asc.c_oAscItemType.Default === val) {
			res = "default";
		} else if (Asc.c_oAscItemType.Sum === val) {
			res = "sum";
		} else if (Asc.c_oAscItemType.CountA === val) {
			res = "countA";
		} else if (Asc.c_oAscItemType.Avg === val) {
			res = "avg";
		} else if (Asc.c_oAscItemType.Max === val) {
			res = "max";
		} else if (Asc.c_oAscItemType.Min === val) {
			res = "min";
		} else if (Asc.c_oAscItemType.Product === val) {
			res = "product";
		} else if (Asc.c_oAscItemType.Count === val) {
			res = "count";
		} else if (Asc.c_oAscItemType.StdDev === val) {
			res = "stdDev";
		} else if (Asc.c_oAscItemType.StdDevP === val) {
			res = "stdDevP";
		} else if (Asc.c_oAscItemType.Var === val) {
			res = "var";
		} else if (Asc.c_oAscItemType.VarP === val) {
			res = "varP";
		} else if (Asc.c_oAscItemType.Grand === val) {
			res = "grand";
		} else if (Asc.c_oAscItemType.Blank === val) {
			res = "blank";
		}
		return res;
	}
	function FromXml_ST_ItemType(val) {
		var res = -1;
		if ("data" === val) {
			res = Asc.c_oAscItemType.Data;
		} else if ("default" === val) {
			res = Asc.c_oAscItemType.Default;
		} else if ("sum" === val) {
			res = Asc.c_oAscItemType.Sum;
		} else if ("countA" === val) {
			res = Asc.c_oAscItemType.CountA;
		} else if ("avg" === val) {
			res = Asc.c_oAscItemType.Avg;
		} else if ("max" === val) {
			res = Asc.c_oAscItemType.Max;
		} else if ("min" === val) {
			res = Asc.c_oAscItemType.Min;
		} else if ("product" === val) {
			res = Asc.c_oAscItemType.Product;
		} else if ("count" === val) {
			res = Asc.c_oAscItemType.Count;
		} else if ("stdDev" === val) {
			res = Asc.c_oAscItemType.StdDev;
		} else if ("stdDevP" === val) {
			res = Asc.c_oAscItemType.StdDevP;
		} else if ("var" === val) {
			res = Asc.c_oAscItemType.Var;
		} else if ("varP" === val) {
			res = Asc.c_oAscItemType.VarP;
		} else if ("grand" === val) {
			res = Asc.c_oAscItemType.Grand;
		} else if ("blank" === val) {
			res = Asc.c_oAscItemType.Blank;
		}
		return res;
	}

	function ToXml_ST_DataConsolidateFunction(val) {
		var res = "";
		if (Asc.c_oAscDataConsolidateFunction.Average === val) {
			res = "average";
		} else if (Asc.c_oAscDataConsolidateFunction.Count === val) {
			res = "count";
		} else if (Asc.c_oAscDataConsolidateFunction.CountNums === val) {
			res = "countNums";
		} else if (Asc.c_oAscDataConsolidateFunction.Max === val) {
			res = "max";
		} else if (Asc.c_oAscDataConsolidateFunction.Min === val) {
			res = "min";
		} else if (Asc.c_oAscDataConsolidateFunction.Product === val) {
			res = "product";
		} else if (Asc.c_oAscDataConsolidateFunction.StdDev === val) {
			res = "stdDev";
		} else if (Asc.c_oAscDataConsolidateFunction.StdDevp === val) {
			res = "stdDevp";
		} else if (Asc.c_oAscDataConsolidateFunction.Sum === val) {
			res = "sum";
		} else if (Asc.c_oAscDataConsolidateFunction.Var === val) {
			res = "var";
		} else if (Asc.c_oAscDataConsolidateFunction.Varp === val) {
			res = "varp";
		}
		return res;
	}
	function FromXml_ST_DataConsolidateFunction(val) {
		var res = -1;
		if ("average" === val) {
			res = Asc.c_oAscDataConsolidateFunction.Average;
		} else if ("count" === val) {
			res = Asc.c_oAscDataConsolidateFunction.Count;
		} else if ("countNums" === val) {
			res = Asc.c_oAscDataConsolidateFunction.CountNums;
		} else if ("max" === val) {
			res = Asc.c_oAscDataConsolidateFunction.Max;
		} else if ("min" === val) {
			res = Asc.c_oAscDataConsolidateFunction.Min;
		} else if ("product" === val) {
			res = Asc.c_oAscDataConsolidateFunction.Product;
		} else if ("stdDev" === val) {
			res = Asc.c_oAscDataConsolidateFunction.StdDev;
		} else if ("stdDevp" === val) {
			res = Asc.c_oAscDataConsolidateFunction.StdDevp;
		} else if ("sum" === val) {
			res = Asc.c_oAscDataConsolidateFunction.Sum;
		} else if ("var" === val) {
			res = Asc.c_oAscDataConsolidateFunction.Var;
		} else if ("varp" === val) {
			res = Asc.c_oAscDataConsolidateFunction.Varp;
		}
		return res;
	}

	function ToXml_ST_ShowDataAs(val) {
		var res = "";
		if (Asc.c_oAscShowDataAs.Normal === val) {
			res = "normal";
		} else if (Asc.c_oAscShowDataAs.Difference === val) {
			res = "difference";
		} else if (Asc.c_oAscShowDataAs.Percent === val) {
			res = "percent";
		} else if (Asc.c_oAscShowDataAs.PercentDiff === val) {
			res = "percentDiff";
		} else if (Asc.c_oAscShowDataAs.RunTotal === val) {
			res = "runTotal";
		} else if (Asc.c_oAscShowDataAs.PercentOfRow === val) {
			res = "percentOfRow";
		} else if (Asc.c_oAscShowDataAs.PercentOfCol === val) {
			res = "percentOfCol";
		} else if (Asc.c_oAscShowDataAs.PercentOfTotal === val) {
			res = "percentOfTotal";
		} else if (Asc.c_oAscShowDataAs.Index === val) {
			res = "index";
		}
		return res;
	}
	function FromXml_ST_ShowDataAs(val) {
		var res = -1;
		if ("normal" === val) {
			res = Asc.c_oAscShowDataAs.Normal;
		} else if ("difference" === val) {
			res = Asc.c_oAscShowDataAs.Difference;
		} else if ("percent" === val) {
			res = Asc.c_oAscShowDataAs.Percent;
		} else if ("percentDiff" === val) {
			res = Asc.c_oAscShowDataAs.PercentDiff;
		} else if ("runTotal" === val) {
			res = Asc.c_oAscShowDataAs.RunTotal;
		} else if ("percentOfRow" === val) {
			res = Asc.c_oAscShowDataAs.PercentOfRow;
		} else if ("percentOfCol" === val) {
			res = Asc.c_oAscShowDataAs.PercentOfCol;
		} else if ("percentOfTotal" === val) {
			res = Asc.c_oAscShowDataAs.PercentOfTotal;
		} else if ("index" === val) {
			res = Asc.c_oAscShowDataAs.Index;
		}
		return res;
	}
	
	function ToXml_ST_FormatAction(val) {
		var res = "";
		if (Asc.c_oAscFormatAction.Blank === val) {
			res = "blank";
		} else if (Asc.c_oAscFormatAction.Formatting === val) {
			res = "formatting";
		} else if (Asc.c_oAscFormatAction.Drill === val) {
			res = "drill";
		} else if (Asc.c_oAscFormatAction.Formula === val) {
			res = "formula";
		}
		return res;
	}
	function FromXml_ST_FormatAction(val) {
		var res = -1;
		if ("blank" === val) {
			res = Asc.c_oAscFormatAction.Blank;
		} else if ("formatting" === val) {
			res = Asc.c_oAscFormatAction.Formatting;
		} else if ("drill" === val) {
			res = Asc.c_oAscFormatAction.Drill;
		} else if ("formula" === val) {
			res = Asc.c_oAscFormatAction.Formula;
		}
		return res;
	}

	function ToXml_ST_Scope(val) {
		var res = "";
		if (Asc.c_oAscScope.Selection === val) {
			res = "selection";
		} else if (Asc.c_oAscScope.Data === val) {
			res = "data";
		} else if (Asc.c_oAscScope.Field === val) {
			res = "field";
		}
		return res;
	}
	function FromXml_ST_Scope(val) {
		var res = -1;
		if ("selection" === val) {
			res = Asc.c_oAscScope.Selection;
		} else if ("data" === val) {
			res = Asc.c_oAscScope.Data;
		} else if ("field" === val) {
			res = Asc.c_oAscScope.Field;
		}
		return res;
	}

	function ToXml_ST_Type(val) {
		var res = "";
		if (Asc.c_oAscType.None === val) {
			res = "none";
		} else if (Asc.c_oAscType.All === val) {
			res = "all";
		} else if (Asc.c_oAscType.Row === val) {
			res = "row";
		} else if (Asc.c_oAscType.Column === val) {
			res = "column";
		}
		return res;
	}
	function FromXml_ST_Type(val) {
		var res = -1;
		if ("none" === val) {
			res = Asc.c_oAscType.None;
		} else if ("all" === val) {
			res = Asc.c_oAscType.All;
		} else if ("row" === val) {
			res = Asc.c_oAscType.Row;
		} else if ("column" === val) {
			res = Asc.c_oAscType.Column;
		}
		return res;
	}
	
	function ToXml_ST_PivotFilterType(val) {
		var res = "";
		if (Asc.c_oAscPivotFilterType.Unknown === val) {
			res = "unknown";
		} else if (Asc.c_oAscPivotFilterType.Count === val) {
			res = "count";
		} else if (Asc.c_oAscPivotFilterType.Percent === val) {
			res = "percent";
		} else if (Asc.c_oAscPivotFilterType.Sum === val) {
			res = "sum";
		} else if (Asc.c_oAscPivotFilterType.CaptionEqual === val) {
			res = "captionEqual";
		} else if (Asc.c_oAscPivotFilterType.CaptionNotEqual === val) {
			res = "captionNotEqual";
		} else if (Asc.c_oAscPivotFilterType.CaptionBeginsWith === val) {
			res = "captionBeginsWith";
		} else if (Asc.c_oAscPivotFilterType.CaptionNotBeginsWith === val) {
			res = "captionNotBeginsWith";
		} else if (Asc.c_oAscPivotFilterType.CaptionEndsWith === val) {
			res = "captionEndsWith";
		} else if (Asc.c_oAscPivotFilterType.CaptionNotEndsWith === val) {
			res = "captionNotEndsWith";
		} else if (Asc.c_oAscPivotFilterType.CaptionContains === val) {
			res = "captionContains";
		} else if (Asc.c_oAscPivotFilterType.CaptionNotContains === val) {
			res = "captionNotContains";
		} else if (Asc.c_oAscPivotFilterType.CaptionGreaterThan === val) {
			res = "captionGreaterThan";
		} else if (Asc.c_oAscPivotFilterType.CaptionGreaterThanOrEqual === val) {
			res = "captionGreaterThanOrEqual";
		} else if (Asc.c_oAscPivotFilterType.CaptionLessThan === val) {
			res = "captionLessThan";
		} else if (Asc.c_oAscPivotFilterType.CaptionLessThanOrEqual === val) {
			res = "captionLessThanOrEqual";
		} else if (Asc.c_oAscPivotFilterType.CaptionBetween === val) {
			res = "captionBetween";
		} else if (Asc.c_oAscPivotFilterType.CaptionNotBetween === val) {
			res = "captionNotBetween";
		} else if (Asc.c_oAscPivotFilterType.ValueEqual === val) {
			res = "valueEqual";
		} else if (Asc.c_oAscPivotFilterType.ValueNotEqual === val) {
			res = "valueNotEqual";
		} else if (Asc.c_oAscPivotFilterType.ValueGreaterThan === val) {
			res = "valueGreaterThan";
		} else if (Asc.c_oAscPivotFilterType.ValueGreaterThanOrEqual === val) {
			res = "valueGreaterThanOrEqual";
		} else if (Asc.c_oAscPivotFilterType.ValueLessThan === val) {
			res = "valueLessThan";
		} else if (Asc.c_oAscPivotFilterType.ValueLessThanOrEqual === val) {
			res = "valueLessThanOrEqual";
		} else if (Asc.c_oAscPivotFilterType.ValueBetween === val) {
			res = "valueBetween";
		} else if (Asc.c_oAscPivotFilterType.ValueNotBetween === val) {
			res = "valueNotBetween";
		} else if (Asc.c_oAscPivotFilterType.DateEqual === val) {
			res = "dateEqual";
		} else if (Asc.c_oAscPivotFilterType.DateNotEqual === val) {
			res = "dateNotEqual";
		} else if (Asc.c_oAscPivotFilterType.DateOlderThan === val) {
			res = "dateOlderThan";
		} else if (Asc.c_oAscPivotFilterType.DateOlderThanOrEqual === val) {
			res = "dateOlderThanOrEqual";
		} else if (Asc.c_oAscPivotFilterType.DateNewerThan === val) {
			res = "dateNewerThan";
		} else if (Asc.c_oAscPivotFilterType.DateNewerThanOrEqual === val) {
			res = "dateNewerThanOrEqual";
		} else if (Asc.c_oAscPivotFilterType.DateBetween === val) {
			res = "dateBetween";
		} else if (Asc.c_oAscPivotFilterType.DateNotBetween === val) {
			res = "dateNotBetween";
		} else if (Asc.c_oAscPivotFilterType.Tomorrow === val) {
			res = "tomorrow";
		} else if (Asc.c_oAscPivotFilterType.Today === val) {
			res = "today";
		} else if (Asc.c_oAscPivotFilterType.Yesterday === val) {
			res = "yesterday";
		} else if (Asc.c_oAscPivotFilterType.NextWeek === val) {
			res = "nextWeek";
		} else if (Asc.c_oAscPivotFilterType.ThisWeek === val) {
			res = "thisWeek";
		} else if (Asc.c_oAscPivotFilterType.LastWeek === val) {
			res = "lastWeek";
		} else if (Asc.c_oAscPivotFilterType.NextMonth === val) {
			res = "nextMonth";
		} else if (Asc.c_oAscPivotFilterType.ThisMonth === val) {
			res = "thisMonth";
		} else if (Asc.c_oAscPivotFilterType.LastMonth === val) {
			res = "lastMonth";
		} else if (Asc.c_oAscPivotFilterType.NextQuarter === val) {
			res = "nextQuarter";
		} else if (Asc.c_oAscPivotFilterType.ThisQuarter === val) {
			res = "thisQuarter";
		} else if (Asc.c_oAscPivotFilterType.LastQuarter === val) {
			res = "lastQuarter";
		} else if (Asc.c_oAscPivotFilterType.NextYear === val) {
			res = "nextYear";
		} else if (Asc.c_oAscPivotFilterType.ThisYear === val) {
			res = "thisYear";
		} else if (Asc.c_oAscPivotFilterType.LastYear === val) {
			res = "lastYear";
		} else if (Asc.c_oAscPivotFilterType.YearToDate === val) {
			res = "yearToDate";
		} else if (Asc.c_oAscPivotFilterType.Q1 === val) {
			res = "Q1";
		} else if (Asc.c_oAscPivotFilterType.Q2 === val) {
			res = "Q2";
		} else if (Asc.c_oAscPivotFilterType.Q3 === val) {
			res = "Q3";
		} else if (Asc.c_oAscPivotFilterType.Q4 === val) {
			res = "Q4";
		} else if (Asc.c_oAscPivotFilterType.M1 === val) {
			res = "M1";
		} else if (Asc.c_oAscPivotFilterType.M2 === val) {
			res = "M2";
		} else if (Asc.c_oAscPivotFilterType.M3 === val) {
			res = "M3";
		} else if (Asc.c_oAscPivotFilterType.M4 === val) {
			res = "M4";
		} else if (Asc.c_oAscPivotFilterType.M5 === val) {
			res = "M5";
		} else if (Asc.c_oAscPivotFilterType.M6 === val) {
			res = "M6";
		} else if (Asc.c_oAscPivotFilterType.M7 === val) {
			res = "M7";
		} else if (Asc.c_oAscPivotFilterType.M8 === val) {
			res = "M8";
		} else if (Asc.c_oAscPivotFilterType.M9 === val) {
			res = "M9";
		} else if (Asc.c_oAscPivotFilterType.M10 === val) {
			res = "M10";
		} else if (Asc.c_oAscPivotFilterType.M11 === val) {
			res = "M11";
		} else if (Asc.c_oAscPivotFilterType.M12 === val) {
			res = "M12";
		}
		return res;
	}
	function FromXml_ST_PivotFilterType(val) {
		var res = -1;
		if ("unknown" === val) {
			res = Asc.c_oAscPivotFilterType.Unknown;
		} else if ("count" === val) {
			res = Asc.c_oAscPivotFilterType.Count;
		} else if ("percent" === val) {
			res = Asc.c_oAscPivotFilterType.Percent;
		} else if ("sum" === val) {
			res = Asc.c_oAscPivotFilterType.Sum;
		} else if ("captionEqual" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionEqual;
		} else if ("captionNotEqual" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionNotEqual;
		} else if ("captionBeginsWith" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionBeginsWith;
		} else if ("captionNotBeginsWith" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionNotBeginsWith;
		} else if ("captionEndsWith" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionEndsWith;
		} else if ("captionNotEndsWith" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionNotEndsWith;
		} else if ("captionContains" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionContains;
		} else if ("captionNotContains" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionNotContains;
		} else if ("captionGreaterThan" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionGreaterThan;
		} else if ("captionGreaterThanOrEqual" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionGreaterThanOrEqual;
		} else if ("captionLessThan" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionLessThan;
		} else if ("captionLessThanOrEqual" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionLessThanOrEqual;
		} else if ("captionBetween" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionBetween;
		} else if ("captionNotBetween" === val) {
			res = Asc.c_oAscPivotFilterType.CaptionNotBetween;
		} else if ("valueEqual" === val) {
			res = Asc.c_oAscPivotFilterType.ValueEqual;
		} else if ("valueNotEqual" === val) {
			res = Asc.c_oAscPivotFilterType.ValueNotEqual;
		} else if ("valueGreaterThan" === val) {
			res = Asc.c_oAscPivotFilterType.ValueGreaterThan;
		} else if ("valueGreaterThanOrEqual" === val) {
			res = Asc.c_oAscPivotFilterType.ValueGreaterThanOrEqual;
		} else if ("valueLessThan" === val) {
			res = Asc.c_oAscPivotFilterType.ValueLessThan;
		} else if ("valueLessThanOrEqual" === val) {
			res = Asc.c_oAscPivotFilterType.ValueLessThanOrEqual;
		} else if ("valueBetween" === val) {
			res = Asc.c_oAscPivotFilterType.ValueBetween;
		} else if ("valueNotBetween" === val) {
			res = Asc.c_oAscPivotFilterType.ValueNotBetween;
		} else if ("dateEqual" === val) {
			res = Asc.c_oAscPivotFilterType.DateEqual;
		} else if ("dateNotEqual" === val) {
			res = Asc.c_oAscPivotFilterType.DateNotEqual;
		} else if ("dateOlderThan" === val) {
			res = Asc.c_oAscPivotFilterType.DateOlderThan;
		} else if ("dateOlderThanOrEqual" === val) {
			res = Asc.c_oAscPivotFilterType.DateOlderThanOrEqual;
		} else if ("dateNewerThan" === val) {
			res = Asc.c_oAscPivotFilterType.DateNewerThan;
		} else if ("dateNewerThanOrEqual" === val) {
			res = Asc.c_oAscPivotFilterType.DateNewerThanOrEqual;
		} else if ("dateBetween" === val) {
			res = Asc.c_oAscPivotFilterType.DateBetween;
		} else if ("dateNotBetween" === val) {
			res = Asc.c_oAscPivotFilterType.DateNotBetween;
		} else if ("tomorrow" === val) {
			res = Asc.c_oAscPivotFilterType.Tomorrow;
		} else if ("today" === val) {
			res = Asc.c_oAscPivotFilterType.Today;
		} else if ("yesterday" === val) {
			res = Asc.c_oAscPivotFilterType.Yesterday;
		} else if ("nextWeek" === val) {
			res = Asc.c_oAscPivotFilterType.NextWeek;
		} else if ("thisWeek" === val) {
			res = Asc.c_oAscPivotFilterType.ThisWeek;
		} else if ("lastWeek" === val) {
			res = Asc.c_oAscPivotFilterType.LastWeek;
		} else if ("nextMonth" === val) {
			res = Asc.c_oAscPivotFilterType.NextMonth;
		} else if ("thisMonth" === val) {
			res = Asc.c_oAscPivotFilterType.ThisMonth;
		} else if ("lastMonth" === val) {
			res = Asc.c_oAscPivotFilterType.LastMonth;
		} else if ("nextQuarter" === val) {
			res = Asc.c_oAscPivotFilterType.NextQuarter;
		} else if ("thisQuarter" === val) {
			res = Asc.c_oAscPivotFilterType.ThisQuarter;
		} else if ("lastQuarter" === val) {
			res = Asc.c_oAscPivotFilterType.LastQuarter;
		} else if ("nextYear" === val) {
			res = Asc.c_oAscPivotFilterType.NextYear;
		} else if ("thisYear" === val) {
			res = Asc.c_oAscPivotFilterType.ThisYear;
		} else if ("lastYear" === val) {
			res = Asc.c_oAscPivotFilterType.LastYear;
		} else if ("yearToDate" === val) {
			res = Asc.c_oAscPivotFilterType.YearToDate;
		} else if ("Q1" === val) {
			res = Asc.c_oAscPivotFilterType.Q1;
		} else if ("Q2" === val) {
			res = Asc.c_oAscPivotFilterType.Q2;
		} else if ("Q3" === val) {
			res = Asc.c_oAscPivotFilterType.Q3;
		} else if ("Q4" === val) {
			res = Asc.c_oAscPivotFilterType.Q4;
		} else if ("M1" === val) {
			res = Asc.c_oAscPivotFilterType.M1;
		} else if ("M2" === val) {
			res = Asc.c_oAscPivotFilterType.M2;
		} else if ("M3" === val) {
			res = Asc.c_oAscPivotFilterType.M3;
		} else if ("M4" === val) {
			res = Asc.c_oAscPivotFilterType.M4;
		} else if ("M5" === val) {
			res = Asc.c_oAscPivotFilterType.M5;
		} else if ("M6" === val) {
			res = Asc.c_oAscPivotFilterType.M6;
		} else if ("M7" === val) {
			res = Asc.c_oAscPivotFilterType.M7;
		} else if ("M8" === val) {
			res = Asc.c_oAscPivotFilterType.M8;
		} else if ("M9" === val) {
			res = Asc.c_oAscPivotFilterType.M9;
		} else if ("M10" === val) {
			res = Asc.c_oAscPivotFilterType.M10;
		} else if ("M11" === val) {
			res = Asc.c_oAscPivotFilterType.M11;
		} else if ("M12" === val) {
			res = Asc.c_oAscPivotFilterType.M12;
		}
		return res;
	}
	
	function ToXML_ST_olapSlicerCacheSortOrder(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.ST_olapSlicerCacheSortOrder.Natural:
				sType = "natural";
				break;
			case Asc.ST_olapSlicerCacheSortOrder.Ascending:
				sType = "ascending";
				break;
			case Asc.ST_olapSlicerCacheSortOrder.Descending:
				sType = "descending";
				break;
		}

		return sType;
	}
	function FromXML_ST_olapSlicerCacheSortOrder(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "natural":
				nType = Asc.ST_olapSlicerCacheSortOrder.Natural;
				break;
			case "ascending":
				nType = Asc.ST_olapSlicerCacheSortOrder.Ascending;
				break;
			case "descending":
				nType = Asc.ST_olapSlicerCacheSortOrder.Descending;
				break;
		}

		return nType;
	}

	function ToXML_ST_slicerCacheCrossFilter(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.ST_slicerCacheCrossFilter.None:
				sType = "none";
				break;
			case Asc.ST_slicerCacheCrossFilter.ShowItemsWithDataAtTop:
				sType = "showItemsWithDataAtTop";
				break;
			case Asc.ST_slicerCacheCrossFilter.ShowItemsWithNoData:
				sType = "showItemsWithNoData";
				break;
		}

		return sType;
	}
	function FromXML_ST_slicerCacheCrossFilter(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "none":
				nType = Asc.ST_slicerCacheCrossFilter.None;
				break;
			case "showItemsWithDataAtTop":
				nType = Asc.ST_slicerCacheCrossFilter.ShowItemsWithDataAtTop;
				break;
			case "showItemsWithNoData":
				nType = Asc.ST_slicerCacheCrossFilter.ShowItemsWithNoData;
				break;
		}

		return nType;
	}

	function ToXML_ST_tabularSlicerCacheSortOrder(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.ST_tabularSlicerCacheSortOrder.Ascending:
				sType = "ascending";
				break;
			case Asc.ST_tabularSlicerCacheSortOrder.Descending:
				sType = "descending";
				break;
		}

		return sType;
	}
	function FromXML_ST_tabularSlicerCacheSortOrder(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "ascending":
				nType = Asc.ST_tabularSlicerCacheSortOrder.Ascending;
				break;
			case "descending":
				nType = Asc.ST_tabularSlicerCacheSortOrder.Descending;
				break;
		}

		return nType;
	}

	function ToXML_ESortBy(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.ESortBy.sortbyCellColor:
				sType = "cellColor";
				break;
			case Asc.ESortBy.sortbyFontColor:
				sType = "fontColor";
				break;
			case Asc.ESortBy.sortbyIcon:
				sType = "icon";
				break;
			case Asc.ESortBy.sortbyValue:
				sType = "value";
				break;
		}

		return sType;
	}
	function FromXML_ESortBy(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "cellColor":
				nType = Asc.ESortBy.sortbyCellColor;
				break;
			case "fontColor":
				nType = Asc.ESortBy.sortbyFontColor;
				break;
			case "icon":
				nType = Asc.ESortBy.sortbyIcon;
				break;
			case "value":
				nType = Asc.ESortBy.sortbyValue;
				break;
		}

		return nType;
	}

	function ToXML_EDataBarAxisPosition(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommonExcel.EDataBarAxisPosition.automatic:
				sType = "auto";
				break;
			case AscCommonExcel.EDataBarAxisPosition.middle:
				sType = "middle";
				break;
			case AscCommonExcel.EDataBarAxisPosition.none:
				sType = "none";
				break;
		}

		return sType;
	}
	function FromXML_EDataBarAxisPosition(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "auto":
				nType = AscCommonExcel.EDataBarAxisPosition.automatic;
				break;
			case "middle":
				nType = AscCommonExcel.EDataBarAxisPosition.middle;
				break;
			case "none":
				nType = AscCommonExcel.EDataBarAxisPosition.none;
				break;
		}

		return nType;
	}

	function ToXML_EDataBarDirection(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommonExcel.EDataBarDirection.context:
				sType = "context";
				break;
			case AscCommonExcel.EDataBarDirection.leftToRight:
				sType = "leftToRight";
				break;
			case AscCommonExcel.EDataBarDirection.rightToLeft:
				sType = "rightToLeft";
				break;
		}

		return sType;
	}
	function FromXML_EDataBarDirection(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "context":
				nType = AscCommonExcel.EDataBarDirection.context;
				break;
			case "leftToRight":
				nType = AscCommonExcel.EDataBarDirection.leftToRight;
				break;
			case "rightToLeft":
				nType = AscCommonExcel.EDataBarDirection.rightToLeft;
				break;
		}

		return nType;
	}

	function ToXml_ST_SourceType(val) {
		var res = "";
		if (Asc.c_oAscSourceType.Worksheet === val) {
			res = "worksheet";
		} else if (Asc.c_oAscSourceType.External === val) {
			res = "external";
		} else if (Asc.c_oAscSourceType.Consolidation === val) {
			res = "consolidation";
		} else if (Asc.c_oAscSourceType.Scenario === val) {
			res = "scenario";
		}
		return res;
	}
	function FromXml_ST_SourceType(val) {
		var res = -1;
		if ("worksheet" === val) {
			res = Asc.c_oAscSourceType.Worksheet;
		} else if ("external" === val) {
			res = Asc.c_oAscSourceType.External;
		} else if ("consolidation" === val) {
			res = Asc.c_oAscSourceType.Consolidation;
		} else if ("scenario" === val) {
			res = Asc.c_oAscSourceType.Scenario;
		}
		return res;
	}

	function ToXml_ST_GroupBy(val) {
		var res = "";
		if (Asc.c_oAscGroupBy.Range === val) {
			res = "range";
		} else if (Asc.c_oAscGroupBy.Seconds === val) {
			res = "seconds";
		} else if (Asc.c_oAscGroupBy.Minutes === val) {
			res = "minutes";
		} else if (Asc.c_oAscGroupBy.Hours === val) {
			res = "hours";
		} else if (Asc.c_oAscGroupBy.Days === val) {
			res = "days";
		} else if (Asc.c_oAscGroupBy.Months === val) {
			res = "months";
		} else if (Asc.c_oAscGroupBy.Quarters === val) {
			res = "quarters";
		} else if (Asc.c_oAscGroupBy.Years === val) {
			res = "years";
		}
		return res;
	}
	function FromXml_ST_GroupBy(val) {
		var res = -1;
		if ("range" === val) {
			res = Asc.c_oAscGroupBy.Range;
		} else if ("seconds" === val) {
			res = Asc.c_oAscGroupBy.Seconds;
		} else if ("minutes" === val) {
			res = Asc.c_oAscGroupBy.Minutes;
		} else if ("hours" === val) {
			res = Asc.c_oAscGroupBy.Hours;
		} else if ("days" === val) {
			res = Asc.c_oAscGroupBy.Days;
		} else if ("months" === val) {
			res = Asc.c_oAscGroupBy.Months;
		} else if ("quarters" === val) {
			res = Asc.c_oAscGroupBy.Quarters;
		} else if ("years" === val) {
			res = Asc.c_oAscGroupBy.Years;
		}
		return res;
	}

	function ToXML_PivotRecType(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case Asc.c_oAscPivotRecType.Number:
				sType = "n";
				break;
			case Asc.c_oAscPivotRecType.String:
				sType = "s";
				break;
			case Asc.c_oAscPivotRecType.Boolean:
				sType = "b";
				break;
			case Asc.c_oAscPivotRecType.Error:
				sType = "e";
				break;
			case Asc.c_oAscPivotRecType.DateTime:
				sType = "d";
				break;
			case Asc.c_oAscPivotRecType.Missing:
				sType = "m";
				break;
			case Asc.c_oAscPivotRecType.Index:
				sType = "i";
				break;
		}

		return sType;
	}
	function FromXML_PivotRecType(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "n":
				nType = Asc.c_oAscPivotRecType.Number;
				break;
			case "s":
				nType = Asc.c_oAscPivotRecType.String;
				break;
			case "b":
				nType = Asc.c_oAscPivotRecType.Boolean;
				break;
			case "e":
				nType = Asc.c_oAscPivotRecType.Error;
				break;
			case "d":
				nType = Asc.c_oAscPivotRecType.DateTime;
				break;
			case "m":
				nType = Asc.c_oAscPivotRecType.Missing;
				break;
			case "i":
				nType = Asc.c_oAscPivotRecType.Index;
				break;
		}

		return nType;
	}
	
	function To_XML_ST_slicerStyleType(nVal)
	{
		var sVal = undefined;
		switch(nVal)
		{
			case Asc.ST_slicerStyleType.unselectedItemWithData:
				sVal = "unselectedItemWithData";
				break;
			case Asc.ST_slicerStyleType.selectedItemWithData:
				sVal = "selectedItemWithData";
				break;
			case Asc.ST_slicerStyleType.unselectedItemWithNoData:
				sVal = "unselectedItemWithNoData";
				break;
			case Asc.ST_slicerStyleType.selectedItemWithNoData:
				sVal = "selectedItemWithNoData";
				break;
			case Asc.ST_slicerStyleType.hoveredUnselectedItemWithData:
				sVal = "hoveredUnselectedItemWithData";
				break;
			case Asc.ST_slicerStyleType.hoveredSelectedItemWithData:
				sVal = "hoveredSelectedItemWithData";
				break;
			case Asc.ST_slicerStyleType.hoveredUnselectedItemWithNoData:
				sVal = "hoveredUnselectedItemWithNoData";
				break;
			case Asc.ST_slicerStyleType.hoveredSelectedItemWithNoData:
				sVal = "hoveredSelectedItemWithNoData";
				break;
		}

		return sVal;
	}
	function From_XML_ST_slicerStyleType(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "unselectedItemWithData":
				nVal = Asc.ST_slicerStyleType.unselectedItemWithData;
				break;
			case "selectedItemWithData":
				nVal = Asc.ST_slicerStyleType.selectedItemWithData;
				break;
			case "unselectedItemWithNoData":
				nVal = Asc.ST_slicerStyleType.unselectedItemWithNoData;
				break;
			case "selectedItemWithNoData":
				nVal = Asc.ST_slicerStyleType.selectedItemWithNoData;
				break;
			case "hoveredUnselectedItemWithData":
				nVal = Asc.ST_slicerStyleType.hoveredUnselectedItemWithData;
				break;
			case "hoveredSelectedItemWithData":
				nVal = Asc.ST_slicerStyleType.hoveredSelectedItemWithData;
				break;
			case "hoveredUnselectedItemWithNoData":
				nVal = Asc.ST_slicerStyleType.hoveredUnselectedItemWithNoData;
				break;
			case "hoveredSelectedItemWithNoData":
				nVal = Asc.ST_slicerStyleType.hoveredSelectedItemWithNoData;
				break;
		}

		return nVal;
	}

	function From_XML_ETableStyleType(sVal)
	{
		var nVal = undefined;
		switch(sVal)
		{
			case "tablestyletypeBlankRow":
				nVal = Asc.ETableStyleType.tablestyletypeBlankRow;
				break;
			case "tablestyletypeFirstColumn":
				nVal = Asc.ETableStyleType.tablestyletypeFirstColumn;
				break;
			case "tablestyletypeFirstColumnStripe":
				nVal = Asc.ETableStyleType.tablestyletypeFirstColumnStripe;
				break;
			case "tablestyletypeFirstColumnSubheading":
				nVal = Asc.ETableStyleType.tablestyletypeFirstColumnSubheading;
				break;
			case "tablestyletypeFirstHeaderCell":
				nVal = Asc.ETableStyleType.tablestyletypeFirstHeaderCell;
				break;
			case "tablestyletypeFirstRowStripe":
				nVal = Asc.ETableStyleType.tablestyletypeFirstRowStripe;
				break;
			case "tablestyletypeFirstRowSubheading":
				nVal = Asc.ETableStyleType.tablestyletypeFirstRowSubheading;
				break;
			case "tablestyletypeFirstSubtotalColumn":
				nVal = Asc.ETableStyleType.tablestyletypeFirstSubtotalColumn;
				break;
			case "tablestyletypeFirstSubtotalRow":
				nVal = Asc.ETableStyleType.tablestyletypeFirstSubtotalRow;
				break;
			case "tablestyletypeFirstTotalCell":
				nVal = Asc.ETableStyleType.tablestyletypeFirstTotalCell;
				break;
			case "tablestyletypeHeaderRow":
				nVal = Asc.ETableStyleType.tablestyletypeHeaderRow;
				break;
			case "tablestyletypeLastColumn":
				nVal = Asc.ETableStyleType.tablestyletypeLastColumn;
				break;
			case "tablestyletypeLastHeaderCell":
				nVal = Asc.ETableStyleType.tablestyletypeLastHeaderCell;
				break;
			case "tablestyletypeLastTotalCell":
				nVal = Asc.ETableStyleType.tablestyletypeLastTotalCell;
				break;
			case "tablestyletypePageFieldLabels":
				nVal = Asc.ETableStyleType.tablestyletypePageFieldLabels;
				break;
			case "tablestyletypePageFieldValues":
				nVal = Asc.ETableStyleType.tablestyletypePageFieldValues;
				break;
			case "tablestyletypeSecondColumnStripe":
				nVal = Asc.ETableStyleType.tablestyletypeSecondColumnStripe;
				break;
			case "tablestyletypeSecondColumnSubheading":
				nVal = Asc.ETableStyleType.tablestyletypeSecondColumnSubheading;
				break;
			case "tablestyletypeSecondRowStripe":
				nVal = Asc.ETableStyleType.tablestyletypeSecondRowStripe;
				break;
			case "tablestyletypeSecondRowSubheading":
				nVal = Asc.ETableStyleType.tablestyletypeSecondRowSubheading;
				break;
			case "tablestyletypeSecondSubtotalColumn":
				nVal = Asc.ETableStyleType.tablestyletypeSecondSubtotalColumn;
				break;
			case "tablestyletypeSecondSubtotalRow":
				nVal = Asc.ETableStyleType.tablestyletypeSecondSubtotalRow;
				break;
			case "tablestyletypeThirdColumnSubheading":
				nVal = Asc.ETableStyleType.tablestyletypeThirdColumnSubheading;
				break;
			case "tablestyletypeThirdRowSubheading":
				nVal = Asc.ETableStyleType.tablestyletypeThirdRowSubheading;
				break;
			case "tablestyletypeThirdSubtotalColumn":
				nVal = Asc.ETableStyleType.tablestyletypeThirdSubtotalColumn;
				break;
			case "tablestyletypeThirdSubtotalRow":
				nVal = Asc.ETableStyleType.tablestyletypeThirdSubtotalRow;
				break;
			case "tablestyletypeTotalRow":
				nVal = Asc.ETableStyleType.tablestyletypeTotalRow;
				break;
			case "tablestyletypeWholeTable":
				nVal = Asc.ETableStyleType.tablestyletypeWholeTable;
				break;
		}

		return nVal;
	}

	function ToXml_ST_CellFormulaType(val) {
		var res = null;
		switch (val) {
			case window["Asc"].ECellFormulaType.cellformulatypeArray:
				res = "array";
				break;
			case window["Asc"].ECellFormulaType.cellformulatypeShared:
				res = "shared";
				break;
			case window["Asc"].ECellFormulaType.cellformulatypeDataTable:
				res = "dataTable";
				break;
		}
		return res;
	}
	function FromXml_ST_CellFormulaType(val) {
		var res;
		switch (val) {
			case "array":
				res = window["Asc"].ECellFormulaType.cellformulatypeArray;
				break;
			case "shared":
				res = window["Asc"].ECellFormulaType.cellformulatypeShared;
				break;
			case "dataTable":
				res = window["Asc"].ECellFormulaType.cellformulatypeDataTable;
				break;
			default:
				res = window["Asc"].ECellFormulaType.cellformulatypeNormal;
		}
		return res;
	}

	function FromXml_ST_CellValueType(val) {
		var res = undefined;
		switch (val) {
			case "s":
				res = AscCommon.CellValueType.String;
				break;
			case "str":
				res = AscCommon.CellValueType.String;
				break;
			case "n":
				res = AscCommon.CellValueType.Number;
				break;
			case "e":
				res = AscCommon.CellValueType.Error;
				break;
			case "b":
				res =  AscCommon.CellValueType.Bool;
				break;
			case "inlineStr":
				res = AscCommon.CellValueType.String;
				break;
			case "d":
				res = AscCommon.CellValueType.String;
				break;
		}
		return res;
	}

	function ToXml_ST_CellValueType(val) {
		var res = undefined;
		switch (val) {
			case AscCommon.CellValueType.String:
				res = "s";
				break;
			/*case AscCommon.CellValueType.String:
				res = "str";
				break;*/
			case AscCommon.CellValueType.Number:
				res = "n";
				break;
			case AscCommon.CellValueType.Error:
				res = "e";
				break;
			case AscCommon.CellValueType.Bool:
				res = "b";
				break;
			/*case "inlineStr":
				res = AscCommon.CellValueType.String;
				break;
			case "d":
				res = AscCommon.CellValueType.String;
				break;*/
		}
		return res;
	}

    //----------------------------------------------------------export----------------------------------------------------
    window['AscCommon']       = window['AscCommon'] || {};
    window['AscFormat']       = window['AscFormat'] || {};
	
})(window);



